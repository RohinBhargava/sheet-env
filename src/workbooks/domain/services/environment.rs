//! Agent-environment controls: seeded scenarios, reset, a deterministic state hash and task graders.
use crate::{
    domain::{
        schemas::{CheckResult, EnvState, TaskList, TaskSummary, VerifyResult},
        services::{
            sheets::OpenWorkbook,
            workbooks::{clean_path, now_ms},
        },
    },
    error::{AppError, Result},
    persistence::{Access, PostgresRepository},
};
use include_dir::{Dir, include_dir};
use serde::Deserialize;
use sha2::{Digest, Sha256};

/// `seed/<scenario>/files/**.xlsx` are the starting workbooks; `seed/<scenario>/tasks.json` the tasks.
static SEED: Dir<'_> = include_dir!("$CARGO_MANIFEST_DIR/seed");
pub const DEFAULT_SCENARIO: &str = "fund-admin";
const GRADER: Access<'static> = Access { actor: "environment-grader", reason: "verify task" };

#[derive(Debug, Clone, Deserialize)]
struct Task {
    id: String,
    title: String,
    instructions: String,
    file: String,
    checks: Vec<Check>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(untagged)]
enum Expected {
    Number(f64),
    Text(String),
    Flag(bool),
}

#[derive(Debug, Clone, Deserialize)]
struct Check {
    sheet: Option<String>,
    cell: String,
    equals: Option<Expected>,
    approx: Option<f64>,
    tolerance: Option<f64>,
}

#[derive(Deserialize)]
#[serde(untagged)]
enum TaskFile {
    List(Vec<Task>),
    Wrapped { tasks: Vec<Task> },
}

#[derive(Clone)]
pub struct EnvironmentService {
    repository: PostgresRepository,
    control_key: Option<String>,
}

fn scenario_dir(scenario: &str) -> Result<&'static Dir<'static>> {
    SEED.get_dir(scenario).ok_or_else(|| AppError::NotFound(format!("scenario {scenario}")))
}

fn seed_files(scenario: &str) -> Result<Vec<(String, Vec<u8>)>> {
    fn walk(dir: &Dir<'_>, root: &str, out: &mut Vec<(String, Vec<u8>)>) {
        for file in dir.files() {
            let path = file.path().to_string_lossy().replace('\\', "/");
            if let Some(relative) = path.strip_prefix(root)
                && relative.to_ascii_lowercase().ends_with(".xlsx")
            {
                out.push((relative.to_owned(), file.contents().to_vec()));
            }
        }
        for child in dir.dirs() {
            walk(child, root, out);
        }
    }
    let mut out = Vec::new();
    if let Some(files) = scenario_dir(scenario)?.get_dir(format!("{scenario}/files")) {
        walk(files, &format!("{scenario}/files/"), &mut out);
    }
    out.sort_by(|a, b| a.0.cmp(&b.0));
    Ok(out)
}

fn tasks(scenario: &str) -> Result<Vec<Task>> {
    let Some(file) = scenario_dir(scenario)?.get_file(format!("{scenario}/tasks.json")) else {
        return Ok(Vec::new());
    };
    let parsed: TaskFile = serde_json::from_slice(file.contents())
        .map_err(|_| AppError::InvalidConfig(format!("seed/{scenario}/tasks.json is malformed")))?;
    Ok(match parsed {
        TaskFile::List(tasks) | TaskFile::Wrapped { tasks } => tasks,
    })
}

fn same_text(actual: &str, expected: &str) -> bool {
    actual.trim().eq_ignore_ascii_case(expected.trim())
}

impl EnvironmentService {
    pub fn new(repository: PostgresRepository, control_key: Option<String>) -> Self {
        Self { repository, control_key: control_key.filter(|key| !key.is_empty()) }
    }

    /// Seed the default scenario the first time the service starts against an empty database.
    pub async fn boot(&self) -> Result<()> {
        if self.repository.meta().await?.is_none() {
            self.repository.reset(DEFAULT_SCENARIO, &seed_files(DEFAULT_SCENARIO)?, now_ms()).await?;
        }
        Ok(())
    }

    pub async fn scenario(&self) -> Result<String> {
        Ok(self.repository.meta().await?.map_or_else(|| DEFAULT_SCENARIO.to_owned(), |(scenario, _)| scenario))
    }

    pub async fn reset(&self, scenario: Option<&str>, key: Option<&str>) -> Result<EnvState> {
        if let Some(expected) = &self.control_key
            && key != Some(expected.as_str())
        {
            return Err(AppError::Forbidden);
        }
        let scenario = scenario.unwrap_or(DEFAULT_SCENARIO);
        self.repository.reset(scenario, &seed_files(scenario)?, now_ms()).await?;
        self.state().await
    }

    pub async fn state(&self) -> Result<EnvState> {
        let (scenario, generation) =
            self.repository.meta().await?.unwrap_or_else(|| (DEFAULT_SCENARIO.to_owned(), 0));
        let files = self.repository.list().await?;
        let mut hash = Sha256::new();
        for file in &files {
            // Timestamps are left out so identical work after different resets hashes the same.
            hash.update(format!("{}\0{}\0{}\n", file.path, file.version, file.content_sha256));
        }
        let state_hash = hash.finalize().iter().map(|byte| format!("{byte:02x}")).collect();
        Ok(EnvState { scenario, generation: u64::from(generation), state_hash, file_count: files.len() as u64 })
    }

    pub async fn tasks(&self) -> Result<TaskList> {
        let scenario = self.scenario().await?;
        Ok(TaskList {
            tasks: tasks(&scenario)?
                .into_iter()
                .map(|task| TaskSummary { id: task.id, title: task.title, instructions: task.instructions, file: task.file })
                .collect(),
            scenario,
        })
    }

    pub async fn verify(&self, id: &str) -> Result<VerifyResult> {
        let scenario = self.scenario().await?;
        let task = tasks(&scenario)?
            .into_iter()
            .find(|task| task.id == id)
            .ok_or_else(|| AppError::NotFound(format!("task {id}")))?;
        let path = clean_path(&task.file)?;
        let mut workbook = match self.repository.read(&path, GRADER).await? {
            Some(bytes) => Some(OpenWorkbook::open(bytes)?),
            None => None,
        };
        let mut checks = Vec::new();
        for check in &task.checks {
            let sheet = check.sheet.clone().unwrap_or_default();
            let (actual, number) = match workbook.as_mut() {
                Some(workbook) => workbook
                    .cell(check.sheet.as_deref(), &check.cell)
                    .unwrap_or_else(|_| (String::from("#MISSING"), None)),
                None => (String::from("#NO FILE"), None),
            };
            let (expected, passed) = match (&check.equals, check.approx) {
                (_, Some(target)) => {
                    let tolerance = check.tolerance.unwrap_or(0.005).abs();
                    (format!("{target} ± {tolerance}"), number.is_some_and(|n| (n - target).abs() <= tolerance))
                }
                (Some(Expected::Number(target)), None) => {
                    (target.to_string(), number.is_some_and(|n| (n - target).abs() <= 1e-9))
                }
                (Some(Expected::Text(text)), None) => (text.clone(), same_text(&actual, text)),
                (Some(Expected::Flag(flag)), None) => {
                    let text = if *flag { "TRUE" } else { "FALSE" };
                    (text.into(), same_text(&actual, text))
                }
                (None, None) => ("(non-empty)".into(), !actual.is_empty()),
            };
            checks.push(CheckResult { sheet, cell: check.cell.clone(), expected, actual, passed });
        }
        Ok(VerifyResult { task_id: task.id, passed: !checks.is_empty() && checks.iter().all(|c| c.passed), checks })
    }
}

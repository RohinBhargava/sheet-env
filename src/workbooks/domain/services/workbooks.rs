//! Workbook files under the Documents root: validated paths, versioned saves, audited reads.
use crate::{
    domain::schemas::{FileEntry, FileList, VersionEntry, VersionList},
    error::{AppError, Result},
    persistence::{Access, PostgresRepository, StoredFile},
};
use std::time::{SystemTime, UNIX_EPOCH};

/// Uploads are limited by the framework's request body limit (2 MiB); this guards direct callers.
pub const MAX_WORKBOOK_BYTES: usize = 2 * 1024 * 1024;

#[derive(Clone)]
pub struct WorkbookService {
    repository: PostgresRepository,
}

pub fn now_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|elapsed| u64::try_from(elapsed.as_millis()).unwrap_or(u64::MAX))
        .unwrap_or_default()
}

/// Normalise a client path: forward slashes, no empty, `.` or `..` segments, no control characters.
pub fn clean_path(path: &str) -> Result<String> {
    let invalid = || AppError::Invalid("path".into());
    let path = path.trim().replace('\\', "/");
    let segments: Vec<&str> = path.split('/').filter(|segment| !segment.is_empty()).collect();
    if segments.is_empty()
        || segments.iter().any(|segment| {
            matches!(*segment, "." | "..")
                || segment.chars().any(|c| c.is_control() || matches!(c, '<' | '>' | ':' | '"' | '|' | '?' | '*'))
                || segment.ends_with(' ')
                || segment.ends_with('.')
        })
    {
        return Err(invalid());
    }
    let clean = segments.join("/");
    if clean.len() > 400 {
        return Err(invalid());
    }
    Ok(clean)
}

fn entry(file: StoredFile) -> FileEntry {
    FileEntry { path: file.path, size: file.size, modified_ms: file.modified_ms, version: u64::from(file.version) }
}

/// An .xlsx file is a zip package; refuse anything else before it reaches storage.
fn looks_like_workbook(bytes: &[u8]) -> bool {
    bytes.starts_with(b"PK\x03\x04")
}

impl WorkbookService {
    pub fn new(repository: PostgresRepository) -> Self {
        Self { repository }
    }

    pub async fn list(&self) -> Result<FileList> {
        Ok(FileList { files: self.repository.list().await?.into_iter().map(entry).collect() })
    }

    pub async fn read(&self, path: &str, access: Access<'_>) -> Result<Vec<u8>> {
        let path = clean_path(path)?;
        match self.repository.read(&path, access).await? {
            Some(bytes) => Ok(bytes),
            None => Err(AppError::NotFound(path)),
        }
    }

    pub async fn save(&self, path: &str, bytes: &[u8]) -> Result<FileEntry> {
        let path = clean_path(path)?;
        if !path.to_ascii_lowercase().ends_with(".xlsx")
            || bytes.len() > MAX_WORKBOOK_BYTES
            || !looks_like_workbook(bytes)
        {
            return Err(AppError::Invalid("workbook".into()));
        }
        Ok(entry(self.repository.save(&path, bytes, now_ms()).await?))
    }

    pub async fn delete(&self, path: &str) -> Result<()> {
        let path = clean_path(path)?;
        if self.repository.delete(&path, now_ms()).await? { Ok(()) } else { Err(AppError::NotFound(path)) }
    }

    pub async fn versions(&self, path: &str) -> Result<VersionList> {
        let path = clean_path(path)?;
        Ok(VersionList {
            versions: self
                .repository
                .versions(&path)
                .await?
                .into_iter()
                .map(|file| VersionEntry { version: u64::from(file.version), modified_ms: file.modified_ms, size: file.size })
                .collect(),
        })
    }
}

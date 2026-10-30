use std::fmt;
use std::collections::BTreeMap;
use forklaunch_universal_sdk as runtime;
pub use runtime::{SdkError, SdkAuth, HmacSigner, SdkResponseWithHeaders};
pub type SdkRequest = runtime::SdkRequest;
fn deserialize_optional_field<'de, D, T>(deserializer: D) -> Result<Option<T>, D::Error>
where
    D: ::serde::Deserializer<'de>,
    T: ::serde::Deserialize<'de>,
{
    T::deserialize(deserializer).map(::std::option::Option::Some)
}
fn deserialize_required_field<'de, D, T>(deserializer: D) -> Result<T, D::Error>
where
    D: ::serde::Deserializer<'de>,
    T: ::serde::Deserialize<'de>,
{
    T::deserialize(deserializer)
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct CellValue {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "address")]
    pub address: String,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "formula")]
    pub formula: Option<Option<String>>,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "number")]
    pub number: Option<Option<f64>>,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "value")]
    pub value: String,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct CheckResult {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "actual")]
    pub actual: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "cell")]
    pub cell: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "expected")]
    pub expected: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "passed")]
    pub passed: bool,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "sheet")]
    pub sheet: String,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
#[serde(deny_unknown_fields)]
pub struct DeleteFileQuery {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "path")]
    pub path: DeleteFileQueryField0,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "kind")]
    pub kind: Option<DeleteFileQueryField1>,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "app")]
    pub app: Option<DeleteFileQueryField2>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct EndpointFailure {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "error")]
    pub error: String,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct EnvState {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "fileCount")]
    pub file_count: u64,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "generation")]
    pub generation: u64,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "scenario")]
    pub scenario: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "stateHash")]
    pub state_hash: String,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct FileEntry {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "modifiedMs")]
    pub modified_ms: u64,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "path")]
    pub path: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "size")]
    pub size: u64,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "version")]
    pub version: u64,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct FileList {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "files")]
    pub files: Vec<FileEntry>,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct HealthResponse {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "recordCount")]
    pub record_count: u64,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "routes")]
    pub routes: Vec<String>,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "service")]
    pub service: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "status")]
    pub status: String,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
#[serde(deny_unknown_fields)]
pub struct ListVersionsQuery {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "path")]
    pub path: ListVersionsQueryField0,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "kind")]
    pub kind: Option<ListVersionsQueryField1>,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "app")]
    pub app: Option<ListVersionsQueryField2>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
#[serde(deny_unknown_fields)]
pub struct ReadCellsQuery {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "path")]
    pub path: ReadCellsQueryField0,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "sheet")]
    pub sheet: Option<ReadCellsQueryField1>,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "range")]
    pub range: Option<ReadCellsQueryField2>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
#[serde(deny_unknown_fields)]
pub struct ReadFileQuery {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "path")]
    pub path: ReadFileQueryField0,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "kind")]
    pub kind: Option<ReadFileQueryField1>,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "app")]
    pub app: Option<ReadFileQueryField2>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct ResetInput {
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "key")]
    pub key: Option<Option<String>>,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "scenario")]
    pub scenario: Option<Option<String>>,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct SaveInput {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "contentBase64")]
    pub content_base_64: String,
    #[serde(
        default,
        deserialize_with = "deserialize_optional_field",
        skip_serializing_if = "::std::option::Option::is_none"
    )]
    #[serde(rename = "kind")]
    pub kind: Option<Option<String>>,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "path")]
    pub path: String,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct SheetCells {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "cells")]
    pub cells: Vec<CellValue>,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "path")]
    pub path: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "range")]
    pub range: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "sheet")]
    pub sheet: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "sheets")]
    pub sheets: Vec<String>,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct TaskList {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "scenario")]
    pub scenario: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "tasks")]
    pub tasks: Vec<TaskSummary>,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct TaskSummary {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "file")]
    pub file: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "id")]
    pub id: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "instructions")]
    pub instructions: String,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "title")]
    pub title: String,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct VerifyResult {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "checks")]
    pub checks: Vec<CheckResult>,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "passed")]
    pub passed: bool,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "taskId")]
    pub task_id: String,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
#[serde(deny_unknown_fields)]
pub struct VerifyTaskParams {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "id")]
    pub id: VerifyTaskParamsField0,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct VersionEntry {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "modifiedMs")]
    pub modified_ms: u64,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "size")]
    pub size: u64,
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "version")]
    pub version: u64,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
#[derive(
    ::std::fmt::Debug,
    ::std::clone::Clone,
    ::std::default::Default,
    ::std::cmp::PartialEq,
    ::serde::Serialize,
    ::serde::Deserialize
)]
pub struct VersionList {
    #[serde(deserialize_with = "deserialize_required_field")]
    #[serde(rename = "versions")]
    pub versions: Vec<VersionEntry>,
    #[serde(flatten)]
    pub additional_properties: BTreeMap<String, ::serde_json::Value>,
}
pub type DeleteFileQueryField0 = String;
pub type DeleteFileQueryField1 = Option<String>;
pub type DeleteFileQueryField2 = Option<String>;
pub type FileListFilesItem = FileEntry;
pub type ListFilesResponseFilesItem = FileListFilesItem;
pub type ListTasksResponseTasksItem = TaskListTasksItem;
pub type ListVersionsQueryField0 = String;
pub type ListVersionsQueryField1 = Option<String>;
pub type ListVersionsQueryField2 = Option<String>;
pub type ListVersionsResponseVersionsItem = VersionListVersionsItem;
pub type ReadCellsQueryField0 = String;
pub type ReadCellsQueryField1 = Option<String>;
pub type ReadCellsQueryField2 = Option<String>;
pub type ReadCellsResponseCellsItem = SheetCellsCellsItem;
pub type ReadFileQueryField0 = String;
pub type ReadFileQueryField1 = Option<String>;
pub type ReadFileQueryField2 = Option<String>;
pub type ResetEnvRequestBody = ResetInput;
pub type SaveFileRequestBody = SaveInput;
pub type SheetCellsCellsItem = CellValue;
pub type TaskListTasksItem = TaskSummary;
pub type VerifyResultChecksItem = CheckResult;
pub type VerifyTaskParamsField0 = String;
pub type VerifyTaskResponseChecksItem = VerifyResultChecksItem;
pub type VersionListVersionsItem = VersionEntry;
#[derive(Default, Clone, PartialEq, Eq, ::serde::Serialize, ::serde::Deserialize)]
pub struct GetHealthRequest {
    pub headers: BTreeMap<String, String>,
}
impl fmt::Debug for GetHealthRequest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("GetHealthRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum GetHealthResponses {
    Status200(Box<HealthResponse>),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}
#[derive(Default, Clone, PartialEq, Eq, ::serde::Serialize, ::serde::Deserialize)]
pub struct ListFilesRequest {
    pub headers: BTreeMap<String, String>,
}
impl fmt::Debug for ListFilesRequest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ListFilesRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum ListFilesResponses {
    Status200(Box<FileList>),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}
#[derive(Default, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub struct ReadFileRequest {
    pub params: BTreeMap<String, String>,
    pub query: ReadFileQuery,
    pub headers: BTreeMap<String, String>,
}
impl fmt::Debug for ReadFileRequest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ReadFileRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum ReadFileResponses {
    Status200(Vec<u8>),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}
#[derive(Clone, PartialEq, Eq, ::serde::Serialize, ::serde::Deserialize)]
pub struct SaveFileRequest<T = SaveFileRequestBody> {
    pub params: BTreeMap<String, String>,
    pub query: BTreeMap<String, String>,
    pub headers: BTreeMap<String, String>,
    pub body: Option<T>,
}
impl<T> Default for SaveFileRequest<T> {
    fn default() -> Self {
        Self {
            params: Default::default(),
            query: Default::default(),
            headers: Default::default(),
            body: None,
        }
    }
}
impl<T> fmt::Debug for SaveFileRequest<T> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SaveFileRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .field("has_body", &self.body.is_some())
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum SaveFileResponses {
    Status200(Box<FileEntry>),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}
#[derive(Default, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub struct DeleteFileRequest {
    pub params: BTreeMap<String, String>,
    pub query: DeleteFileQuery,
    pub headers: BTreeMap<String, String>,
}
impl fmt::Debug for DeleteFileRequest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DeleteFileRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum DeleteFileResponses {
    Status204(()),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}
#[derive(Default, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub struct ListVersionsRequest {
    pub params: BTreeMap<String, String>,
    pub query: ListVersionsQuery,
    pub headers: BTreeMap<String, String>,
}
impl fmt::Debug for ListVersionsRequest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ListVersionsRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum ListVersionsResponses {
    Status200(Box<VersionList>),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}
#[derive(Default, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub struct ReadCellsRequest {
    pub params: BTreeMap<String, String>,
    pub query: ReadCellsQuery,
    pub headers: BTreeMap<String, String>,
}
impl fmt::Debug for ReadCellsRequest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ReadCellsRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum ReadCellsResponses {
    Status200(Box<SheetCells>),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}
#[derive(Default, Clone, PartialEq, Eq, ::serde::Serialize, ::serde::Deserialize)]
pub struct GetEnvStateRequest {
    pub headers: BTreeMap<String, String>,
}
impl fmt::Debug for GetEnvStateRequest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("GetEnvStateRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum GetEnvStateResponses {
    Status200(Box<EnvState>),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}
#[derive(Clone, PartialEq, Eq, ::serde::Serialize, ::serde::Deserialize)]
pub struct ResetEnvRequest<T = ResetEnvRequestBody> {
    pub params: BTreeMap<String, String>,
    pub query: BTreeMap<String, String>,
    pub headers: BTreeMap<String, String>,
    pub body: Option<T>,
}
impl<T> Default for ResetEnvRequest<T> {
    fn default() -> Self {
        Self {
            params: Default::default(),
            query: Default::default(),
            headers: Default::default(),
            body: None,
        }
    }
}
impl<T> fmt::Debug for ResetEnvRequest<T> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ResetEnvRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .field("has_body", &self.body.is_some())
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum ResetEnvResponses {
    Status200(Box<EnvState>),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}
#[derive(Default, Clone, PartialEq, Eq, ::serde::Serialize, ::serde::Deserialize)]
pub struct ListTasksRequest {
    pub headers: BTreeMap<String, String>,
}
impl fmt::Debug for ListTasksRequest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ListTasksRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum ListTasksResponses {
    Status200(Box<TaskList>),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}
#[derive(Default, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub struct VerifyTaskRequest {
    pub params: VerifyTaskParams,
    pub query: BTreeMap<String, String>,
    pub headers: BTreeMap<String, String>,
}
impl fmt::Debug for VerifyTaskRequest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("VerifyTaskRequest")
            .field("params", &"<redacted>")
            .field("query", &"<redacted>")
            .field("headers", &"<redacted>")
            .finish_non_exhaustive()
    }
}
#[derive(Debug, Clone, PartialEq, ::serde::Serialize, ::serde::Deserialize)]
pub enum VerifyTaskResponses {
    Status200(Box<VerifyResult>),
    Status400(Box<EndpointFailure>),
    Status401(Box<EndpointFailure>),
    Status403(Box<EndpointFailure>),
    Status404(Box<EndpointFailure>),
    Status408(Box<EndpointFailure>),
    Status409(Box<EndpointFailure>),
    Status413(Box<EndpointFailure>),
    Status415(Box<EndpointFailure>),
    Status429(Box<EndpointFailure>),
    Status500(Box<EndpointFailure>),
    Status502(Box<EndpointFailure>),
    Status503(Box<EndpointFailure>),
}

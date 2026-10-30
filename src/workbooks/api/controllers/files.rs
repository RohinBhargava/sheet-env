use crate::{
    api::controllers::endpoint_error,
    domain::schemas::{CellQuery, FileEntry, FileList, FileQuery, SaveInput, SheetCells, VersionList},
    domain::services::sheets::OpenWorkbook,
    persistence::Access,
    registrations::AppState,
};
use base64::{Engine, engine::general_purpose::STANDARD};
use forklaunch_axum::native::{EndpointError, State};
use forklaunch_core::contract::endpoint::{NoFields, RequestFields};

/// The service has no sign-in; every reader is the hosted workbook user.
const OPEN: Access<'static> = Access { actor: "workbook-user", reason: "open workbook" };
const AGENT_READ: Access<'static> = Access { actor: "workbook-user", reason: "read cells" };

pub async fn list_files(State(state): State<AppState>, (): ()) -> Result<FileList, EndpointError> {
    state.workbooks.list().await.map_err(endpoint_error)
}
pub async fn read_file(
    State(state): State<AppState>,
    request: RequestFields<(), FileQuery, NoFields>,
) -> Result<Vec<u8>, EndpointError> {
    state.workbooks.read(&request.query.path, OPEN).await.map_err(endpoint_error)
}
pub async fn save_file(State(state): State<AppState>, input: SaveInput) -> Result<FileEntry, EndpointError> {
    let bytes = STANDARD.decode(input.content_base64.trim()).map_err(|_| EndpointError::InvalidRequest)?;
    state.workbooks.save(&input.path, &bytes).await.map_err(endpoint_error)
}
pub async fn delete_file(
    State(state): State<AppState>,
    request: RequestFields<(), FileQuery, NoFields>,
) -> Result<(), EndpointError> {
    state.workbooks.delete(&request.query.path).await.map_err(endpoint_error)
}
pub async fn list_versions(
    State(state): State<AppState>,
    request: RequestFields<(), FileQuery, NoFields>,
) -> Result<VersionList, EndpointError> {
    state.workbooks.versions(&request.query.path).await.map_err(endpoint_error)
}
pub async fn read_cells(
    State(state): State<AppState>,
    request: RequestFields<(), CellQuery, NoFields>,
) -> Result<SheetCells, EndpointError> {
    let query = request.query;
    let bytes = state.workbooks.read(&query.path, AGENT_READ).await.map_err(endpoint_error)?;
    let read = OpenWorkbook::open(bytes)
        .and_then(|mut workbook| workbook.read(query.sheet.as_deref(), query.range.as_deref()))
        .map_err(endpoint_error)?;
    Ok(SheetCells { path: query.path, sheet: read.sheet, sheets: read.sheets, range: read.range, cells: read.cells })
}

#[forklaunch_core::contract::api]
pub mod contract {
    use crate::domain::schemas::{CellQuery, FileEntry, FileList, FileQuery, SaveInput, SheetCells, VersionList};
    use forklaunch_core::contract::{
        HttpMethod, MediaType, RouteContract,
        endpoint::{BinaryBody, Endpoint, EndpointBody, Fields, JsonBody, NoBody, NoFields},
    };

    /// Build the endpoint as `Endpoint::new` does, but kept out of MCP.
    fn hidden<I: EndpointBody, O: EndpointBody>(
        name: &str,
        summary: &str,
        method: HttpMethod,
        path: &str,
        status: u16,
    ) -> Result<Endpoint<I, O>, &'static str> {
        let mut route = RouteContract::new(name, summary, method, path).with_response(status, O::contract());
        if let Some(body) = I::contract() {
            route = route.with_request_body(body);
        }
        if let Some(fields) = I::path_fields()? {
            route.request.params = fields;
        }
        route.request.body_optional = I::body_optional();
        route.request.query = I::query_fields()?;
        route.request.headers = I::header_fields()?;
        route.request.query_schema = I::query_schema()?;
        route.request.headers_schema = I::headers_schema()?;
        route.options.mcp = false;
        Endpoint::from_contract(route)
    }

    #[access(Public)]
    pub fn list_files() -> Endpoint<NoBody, JsonBody<FileList>> {
        Endpoint::new("list_files", "List every workbook (current versions), in path order", HttpMethod::Get, "/api/files", 200)
            .expect("valid list files endpoint")
    }
    #[access(Public)]
    pub fn read_file() -> Endpoint<Fields<NoBody, FileQuery, NoFields>, BinaryBody> {
        hidden("read_file", "Download a workbook's current .xlsx bytes", HttpMethod::Get, "/api/files/raw", 200)
                .and_then(|endpoint| endpoint.with_response_media(MediaType::Spreadsheet))
                .expect("valid read file endpoint")
    }
    #[access(Public)]
    pub fn save_file() -> Endpoint<JsonBody<SaveInput>, JsonBody<FileEntry>> {
        Endpoint::new(
            "save_file",
            "Save a workbook (.xlsx as base64) as the next version of its path",
            HttpMethod::Post,
            "/api/files/save",
            200,
        )
        .expect("valid save file endpoint")
    }
    #[access(Public)]
    pub fn delete_file() -> Endpoint<Fields<NoBody, FileQuery, NoFields>, NoBody> {
        hidden("delete_file", "Delete a workbook (its history is kept)", HttpMethod::Delete, "/api/files/raw", 204)
                .expect("valid delete file endpoint")
    }
    #[access(Public)]
    pub fn list_versions() -> Endpoint<Fields<NoBody, FileQuery, NoFields>, JsonBody<VersionList>> {
        Endpoint::new("list_versions", "List a workbook's saved versions, newest first", HttpMethod::Get, "/api/files/versions", 200)
            .expect("valid list versions endpoint")
    }
    #[access(Public)]
    pub fn read_cells() -> Endpoint<Fields<NoBody, CellQuery, NoFields>, JsonBody<SheetCells>> {
        Endpoint::new(
            "read_cells",
            "Read the non-empty cells of a worksheet range as last saved: cached values, numbers and formulas (at most 5000 cells)",
            HttpMethod::Get,
            "/api/files/cells",
            200,
        )
        .expect("valid read cells endpoint")
    }
}

use crate::{
    api::controllers::endpoint_error,
    domain::schemas::{EnvState, ResetInput, TaskList, TaskPath, VerifyResult},
    registrations::AppState,
};
use forklaunch_axum::native::{EndpointError, State};
use forklaunch_core::contract::endpoint::{NoFields, RequestFields};

pub async fn get_env_state(State(state): State<AppState>, (): ()) -> Result<EnvState, EndpointError> {
    state.environment.state().await.map_err(endpoint_error)
}
pub async fn reset_env(State(state): State<AppState>, input: ResetInput) -> Result<EnvState, EndpointError> {
    state.environment.reset(input.scenario.as_deref(), input.key.as_deref()).await.map_err(endpoint_error)
}
pub async fn list_tasks(State(state): State<AppState>, (): ()) -> Result<TaskList, EndpointError> {
    state.environment.tasks().await.map_err(endpoint_error)
}
pub async fn verify_task(
    State(state): State<AppState>,
    request: RequestFields<(), NoFields, NoFields, TaskPath>,
) -> Result<VerifyResult, EndpointError> {
    state.environment.verify(&request.path.id).await.map_err(endpoint_error)
}

#[forklaunch_core::contract::api]
pub mod contract {
    use crate::domain::schemas::{EnvState, ResetInput, TaskList, TaskPath, VerifyResult};
    use forklaunch_core::contract::{
        HttpMethod, RouteContract,
        endpoint::{Endpoint, EndpointBody, Fields, JsonBody, NoBody, NoFields},
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
    pub fn get_env_state() -> Endpoint<NoBody, JsonBody<EnvState>> {
        hidden("get_env_state", "Scenario, reset generation and a hash of every current workbook", HttpMethod::Get, "/api/env/state", 200)
                .expect("valid env state endpoint")
    }
    #[access(Public)]
    pub fn reset_env() -> Endpoint<JsonBody<ResetInput>, JsonBody<EnvState>> {
        hidden("reset_env", "Replace every workbook with a scenario's seed files", HttpMethod::Post, "/api/env/reset", 200)
                .expect("valid reset endpoint")
    }
    #[access(Public)]
    pub fn list_tasks() -> Endpoint<NoBody, JsonBody<TaskList>> {
        Endpoint::new("list_tasks", "List the current scenario's tasks and the workbook each one uses", HttpMethod::Get, "/api/env/tasks", 200)
            .expect("valid list tasks endpoint")
    }
    #[access(Public)]
    pub fn verify_task() -> Endpoint<Fields<NoBody, NoFields, NoFields, TaskPath>, JsonBody<VerifyResult>> {
        hidden("verify_task", "Grade a task against the workbook as last saved", HttpMethod::Post, "/api/env/tasks/:id/verify", 200)
                .expect("valid verify endpoint")
    }
}

use crate::{
    api::controllers::environment::{
        contract::{Bindings, GetEnvState, ListTasks, ResetEnv, VerifyTask},
        get_env_state, list_tasks, reset_env, verify_task,
    },
    registrations::AppState,
};
use forklaunch_axum::native::ContractRouter;
pub fn mount(router: ContractRouter<AppState>) -> Result<ContractRouter<AppState>, String> {
    Bindings {
        get_env_state: GetEnvState::bind(get_env_state),
        reset_env: ResetEnv::bind(reset_env),
        list_tasks: ListTasks::bind(list_tasks),
        verify_task: VerifyTask::bind(verify_task),
    }
    .mount(router)
}

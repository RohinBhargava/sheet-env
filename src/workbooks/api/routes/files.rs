use crate::{
    api::controllers::files::{
        contract::{Bindings, DeleteFile, ListFiles, ListVersions, ReadCells, ReadFile, SaveFile},
        delete_file, list_files, list_versions, read_cells, read_file, save_file,
    },
    registrations::AppState,
};
use forklaunch_axum::native::ContractRouter;
pub fn mount(router: ContractRouter<AppState>) -> Result<ContractRouter<AppState>, String> {
    Bindings {
        list_files: ListFiles::bind(list_files),
        read_file: ReadFile::bind(read_file),
        save_file: SaveFile::bind(save_file),
        delete_file: DeleteFile::bind(delete_file),
        list_versions: ListVersions::bind(list_versions),
        read_cells: ReadCells::bind(read_cells),
    }
    .mount(router)
}

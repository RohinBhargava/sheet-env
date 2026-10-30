//! Persisted domain model: classification and protection are part of the stored type.
use forklaunch_core::common::FrameworkResult;
use forklaunch_core::persistence::{
    compliance::ComplianceMetadata,
    protected::{ProtectedField, storage::ProtectedModel},
};
use uuid::Uuid;

/// One sealed slice of a workbook version. Fund workbooks name investors and their
/// commitments, so the bytes are classified and never stored in plaintext.
#[derive(Debug, ComplianceMetadata)]
#[compliance(entity = "FileChunk", persisted)]
pub struct FileChunk {
    #[compliance(level = "none")]
    pub id: Uuid,
    #[compliance(level = "none")]
    pub tenant_id: String,
    #[compliance(level = "none")]
    pub version_id: Uuid,
    #[compliance(level = "none")]
    pub seq: i32,
    #[compliance(level = "pii", purpose = "workbook-content-v1")]
    pub content_ciphertext: ProtectedField,
}
impl FileChunk {
    /// Check declared member bindings before SQL writes and after row hydration.
    pub fn validate(&self) -> FrameworkResult<()> {
        self.validate_protected_storage()
    }
}

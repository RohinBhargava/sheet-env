use std::fmt;
use forklaunch_universal_sdk as runtime;
use super::models::*;
use super::schemas;
type ClientRuntime = runtime::UniversalSdk;
#[derive(Clone)]
pub struct Client {
    runtime: ClientRuntime,
}
impl fmt::Debug for Client {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Client").field("runtime", &self.runtime).finish()
    }
}
impl Client {
    #[must_use]
    pub fn new(base_url: impl Into<String>) -> Self {
        Self {
            runtime: ClientRuntime::new(base_url),
        }
    }
    /// Custom transports must enforce TLS verification and disable credential-bearing redirects.
    #[must_use]
    pub fn with_http_client(base_url: impl Into<String>, http: reqwest::Client) -> Self {
        Self {
            runtime: ClientRuntime::with_http_client(base_url, http),
        }
    }
    /// Local development only: permits plaintext credentials outside loopback.
    #[must_use]
    pub fn with_plaintext_development_transport(mut self) -> Self {
        self.runtime = self.runtime.with_plaintext_development_transport();
        self
    }
    #[must_use]
    pub fn with_default_header(
        mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Self {
        self.runtime = self.runtime.with_default_header(name, value);
        self
    }
    #[must_use]
    pub fn with_bearer_token(mut self, token: impl Into<String>) -> Self {
        self.runtime = self.runtime.with_bearer_token(token);
        self
    }
    #[must_use]
    pub fn with_jwt(self, token: impl Into<String>) -> Self {
        self.with_bearer_token(token)
    }
    #[must_use]
    pub fn with_basic_auth(
        mut self,
        username: impl Into<String>,
        password: impl Into<String>,
    ) -> Self {
        self.runtime = self.runtime.with_basic_auth(username, password);
        self
    }
    /// One fixed `authorization` value; an HMAC credential verifies at most once.
    /// Prefer `with_hmac_secret` for service-to-service calls.
    #[must_use]
    pub fn with_hmac_authorization(mut self, value: impl Into<String>) -> Self {
        self.runtime = self.runtime.with_hmac_authorization(value);
        self
    }
    /// Sign every request with the ForkLaunch request HMAC (`keyId=default`): the final
    /// method, URL path, body bytes, a fresh timestamp and a random nonce.
    #[must_use]
    pub fn with_hmac_secret(mut self, secret: impl Into<Vec<u8>>) -> Self {
        self.runtime = self.runtime.with_hmac_secret(secret);
        self
    }
    /// Sign every request with a named server key.
    #[must_use]
    pub fn with_hmac_signer(mut self, signer: HmacSigner) -> Self {
        self.runtime = self.runtime.with_hmac_signer(signer);
        self
    }
    #[must_use]
    pub fn with_auth_hook<F>(mut self, hook: F) -> Self
    where
        F: Fn(&mut SdkRequest) + Send + Sync + 'static,
    {
        self.runtime = self.runtime.with_auth_hook(hook);
        self
    }
    /// Construct a transport request.
    #[must_use]
    pub fn get_health_request(&self, input: GetHealthRequest) -> SdkRequest {
        let path = runtime::path_with_params(
            "/health",
            &::std::collections::BTreeMap::new(),
        );
        SdkRequest {
            method: "GET",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: ::std::collections::BTreeMap::new(),
            headers: input.headers,
            body: None,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        }
    }
    /// Invoke this operation without request parameters or a body.
    pub async fn get_health(&self) -> Result<GetHealthResponses, SdkError> {
        self.get_health_with_options(GetHealthRequest::default()).await
    }
    /// Invoke this operation with per-request transport headers.
    pub async fn get_health_with_options(
        &self,
        input: GetHealthRequest,
    ) -> Result<GetHealthResponses, SdkError> {
        let request = self.get_health_request(input);
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            200 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS200_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status200(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetHealthResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
    /// Construct a transport request.
    #[must_use]
    pub fn list_files_request(&self, input: ListFilesRequest) -> SdkRequest {
        let path = runtime::path_with_params(
            "/api/files",
            &::std::collections::BTreeMap::new(),
        );
        SdkRequest {
            method: "GET",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: ::std::collections::BTreeMap::new(),
            headers: input.headers,
            body: None,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        }
    }
    /// Invoke this operation without request parameters or a body.
    pub async fn list_files(&self) -> Result<ListFilesResponses, SdkError> {
        self.list_files_with_options(ListFilesRequest::default()).await
    }
    /// Invoke this operation with per-request transport headers.
    pub async fn list_files_with_options(
        &self,
        input: ListFilesRequest,
    ) -> Result<ListFilesResponses, SdkError> {
        let request = self.list_files_request(input);
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            200 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::LIST_FILES_STATUS200_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status200(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListFilesResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
    /// Construct a transport request.
    pub fn read_file_request(
        &self,
        input: ReadFileRequest,
    ) -> Result<SdkRequest, serde_json::Error> {
        let path = runtime::path_with_params("/api/files/raw", &input.params);
        Ok(SdkRequest {
            method: "GET",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: runtime::request_fields(&input.query)?,
            headers: input.headers,
            body: None,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        })
    }
    pub async fn read_file(
        &self,
        input: ReadFileRequest,
    ) -> Result<ReadFileResponses, SdkError> {
        let request = self.read_file_request(input).map_err(SdkError::Encode)?;
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            200 if _content_type
                == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => {
                let body = response.bytes().await?.to_vec();
                Ok(ReadFileResponses::Status200(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadFileResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
    /// Serialize a schema-typed body before constructing the transport request.
    pub fn save_file_request(
        &self,
        input: SaveFileRequest,
    ) -> Result<SdkRequest, serde_json::Error> {
        Ok(
            self
                .save_file_raw_request(SaveFileRequest::<String> {
                    params: input.params,
                    query: input.query,
                    headers: input.headers,
                    body: input
                        .body
                        .map(|body| serde_json::to_string(&body))
                        .transpose()?,
                }),
        )
    }
    /// Execute a schema-typed request. Raw encoded bytes require the explicit `_raw` method.
    pub async fn save_file(
        &self,
        input: SaveFileRequest,
    ) -> Result<SaveFileResponses, SdkError> {
        let input = SaveFileRequest::<String> {
            params: input.params,
            query: input.query,
            headers: input.headers,
            body: input
                .body
                .map(|body| serde_json::to_string(&body))
                .transpose()
                .map_err(SdkError::Encode)?,
        };
        self.save_file_raw(input).await
    }
    /// Construct a transport request from explicitly encoded wire data.
    #[must_use]
    pub fn save_file_raw_request(&self, input: SaveFileRequest<String>) -> SdkRequest {
        let path = runtime::path_with_params("/api/files/save", &input.params);
        SdkRequest {
            method: "POST",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: input.query,
            headers: input.headers,
            body: input.body,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        }
    }
    pub async fn save_file_raw(
        &self,
        input: SaveFileRequest<String>,
    ) -> Result<SaveFileResponses, SdkError> {
        let request = self.save_file_raw_request(input);
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            200 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::SAVE_FILE_STATUS200_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status200(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(SaveFileResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
    /// Construct a transport request.
    pub fn delete_file_request(
        &self,
        input: DeleteFileRequest,
    ) -> Result<SdkRequest, serde_json::Error> {
        let path = runtime::path_with_params("/api/files/raw", &input.params);
        Ok(SdkRequest {
            method: "DELETE",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: runtime::request_fields(&input.query)?,
            headers: input.headers,
            body: None,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        })
    }
    pub async fn delete_file(
        &self,
        input: DeleteFileRequest,
    ) -> Result<DeleteFileResponses, SdkError> {
        let request = self.delete_file_request(input).map_err(SdkError::Encode)?;
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            204 => {
                let bytes = response.bytes().await?;
                let _ = bytes;
                let body = ();
                Ok(DeleteFileResponses::Status204(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(DeleteFileResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
    /// Construct a transport request.
    pub fn list_versions_request(
        &self,
        input: ListVersionsRequest,
    ) -> Result<SdkRequest, serde_json::Error> {
        let path = runtime::path_with_params("/api/files/versions", &input.params);
        Ok(SdkRequest {
            method: "GET",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: runtime::request_fields(&input.query)?,
            headers: input.headers,
            body: None,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        })
    }
    pub async fn list_versions(
        &self,
        input: ListVersionsRequest,
    ) -> Result<ListVersionsResponses, SdkError> {
        let request = self.list_versions_request(input).map_err(SdkError::Encode)?;
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            200 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::LIST_VERSIONS_STATUS200_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status200(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListVersionsResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
    /// Construct a transport request.
    pub fn read_cells_request(
        &self,
        input: ReadCellsRequest,
    ) -> Result<SdkRequest, serde_json::Error> {
        let path = runtime::path_with_params("/api/files/cells", &input.params);
        Ok(SdkRequest {
            method: "GET",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: runtime::request_fields(&input.query)?,
            headers: input.headers,
            body: None,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        })
    }
    pub async fn read_cells(
        &self,
        input: ReadCellsRequest,
    ) -> Result<ReadCellsResponses, SdkError> {
        let request = self.read_cells_request(input).map_err(SdkError::Encode)?;
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            200 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::READ_CELLS_STATUS200_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status200(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ReadCellsResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
    /// Construct a transport request.
    #[must_use]
    pub fn get_env_state_request(&self, input: GetEnvStateRequest) -> SdkRequest {
        let path = runtime::path_with_params(
            "/api/env/state",
            &::std::collections::BTreeMap::new(),
        );
        SdkRequest {
            method: "GET",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: ::std::collections::BTreeMap::new(),
            headers: input.headers,
            body: None,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        }
    }
    /// Invoke this operation without request parameters or a body.
    pub async fn get_env_state(&self) -> Result<GetEnvStateResponses, SdkError> {
        self.get_env_state_with_options(GetEnvStateRequest::default()).await
    }
    /// Invoke this operation with per-request transport headers.
    pub async fn get_env_state_with_options(
        &self,
        input: GetEnvStateRequest,
    ) -> Result<GetEnvStateResponses, SdkError> {
        let request = self.get_env_state_request(input);
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            200 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_ENV_STATE_STATUS200_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status200(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(GetEnvStateResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
    /// Serialize a schema-typed body before constructing the transport request.
    pub fn reset_env_request(
        &self,
        input: ResetEnvRequest,
    ) -> Result<SdkRequest, serde_json::Error> {
        Ok(
            self
                .reset_env_raw_request(ResetEnvRequest::<String> {
                    params: input.params,
                    query: input.query,
                    headers: input.headers,
                    body: input
                        .body
                        .map(|body| serde_json::to_string(&body))
                        .transpose()?,
                }),
        )
    }
    /// Execute a schema-typed request. Raw encoded bytes require the explicit `_raw` method.
    pub async fn reset_env(
        &self,
        input: ResetEnvRequest,
    ) -> Result<ResetEnvResponses, SdkError> {
        let input = ResetEnvRequest::<String> {
            params: input.params,
            query: input.query,
            headers: input.headers,
            body: input
                .body
                .map(|body| serde_json::to_string(&body))
                .transpose()
                .map_err(SdkError::Encode)?,
        };
        self.reset_env_raw(input).await
    }
    /// Construct a transport request from explicitly encoded wire data.
    #[must_use]
    pub fn reset_env_raw_request(&self, input: ResetEnvRequest<String>) -> SdkRequest {
        let path = runtime::path_with_params("/api/env/reset", &input.params);
        SdkRequest {
            method: "POST",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: input.query,
            headers: input.headers,
            body: input.body,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        }
    }
    pub async fn reset_env_raw(
        &self,
        input: ResetEnvRequest<String>,
    ) -> Result<ResetEnvResponses, SdkError> {
        let request = self.reset_env_raw_request(input);
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            200 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_ENV_STATE_STATUS200_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status200(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ResetEnvResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
    /// Construct a transport request.
    #[must_use]
    pub fn list_tasks_request(&self, input: ListTasksRequest) -> SdkRequest {
        let path = runtime::path_with_params(
            "/api/env/tasks",
            &::std::collections::BTreeMap::new(),
        );
        SdkRequest {
            method: "GET",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: ::std::collections::BTreeMap::new(),
            headers: input.headers,
            body: None,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        }
    }
    /// Invoke this operation without request parameters or a body.
    pub async fn list_tasks(&self) -> Result<ListTasksResponses, SdkError> {
        self.list_tasks_with_options(ListTasksRequest::default()).await
    }
    /// Invoke this operation with per-request transport headers.
    pub async fn list_tasks_with_options(
        &self,
        input: ListTasksRequest,
    ) -> Result<ListTasksResponses, SdkError> {
        let request = self.list_tasks_request(input);
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            200 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::LIST_TASKS_STATUS200_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status200(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(ListTasksResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
    /// Construct a transport request.
    pub fn verify_task_request(
        &self,
        input: VerifyTaskRequest,
    ) -> Result<SdkRequest, serde_json::Error> {
        let path = runtime::path_with_params(
            "/api/env/tasks/:id/verify",
            &runtime::request_fields(&input.params)?,
        );
        Ok(SdkRequest {
            method: "POST",
            url: format!("{}{}", self.runtime.base_url(), path),
            query: input.query,
            headers: input.headers,
            body: None,
            body_bytes: None,
            multipart: None,
            content_type: "application/json",
        })
    }
    pub async fn verify_task(
        &self,
        input: VerifyTaskRequest,
    ) -> Result<VerifyTaskResponses, SdkError> {
        let request = self.verify_task_request(input).map_err(SdkError::Encode)?;
        let response = self.runtime.execute(request).await?;
        let status = response.status().as_u16();
        let _content_type = runtime::response_content_type(&response)?;
        match status {
            200 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::VERIFY_TASK_STATUS200_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status200(body))
            }
            400 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status400(body))
            }
            401 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status401(body))
            }
            403 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status403(body))
            }
            404 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status404(body))
            }
            408 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status408(body))
            }
            409 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status409(body))
            }
            413 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status413(body))
            }
            415 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status415(body))
            }
            429 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status429(body))
            }
            500 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status500(body))
            }
            502 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status502(body))
            }
            503 if _content_type == "application/json" => {
                let body = runtime::decode_json_response(
                        response,
                        schemas::GET_HEALTH_STATUS400_BODY,
                        None,
                    )
                    .await?;
                Ok(VerifyTaskResponses::Status503(body))
            }
            _ => Err(SdkError::Contract("undeclared response status or content type")),
        }
    }
}

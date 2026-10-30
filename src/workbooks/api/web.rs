//! The spreadsheet web client, compiled into the binary from `web-dist` (built by `src/web`).
//! Unknown non-API paths fall back to `index.html` so client routes survive a reload.
use axum::{
    http::{StatusCode, Uri, header},
    response::{IntoResponse, Response},
};
use include_dir::{Dir, include_dir};

static WEB: Dir<'_> = include_dir!("$CARGO_MANIFEST_DIR/web-dist");

fn content_type(path: &str) -> &'static str {
    match path.rsplit('.').next().unwrap_or_default().to_ascii_lowercase().as_str() {
        "html" => "text/html; charset=utf-8",
        "js" | "mjs" => "text/javascript; charset=utf-8",
        "css" => "text/css; charset=utf-8",
        "json" | "map" => "application/json",
        "svg" => "image/svg+xml",
        "png" => "image/png",
        "ico" => "image/x-icon",
        "woff2" => "font/woff2",
        "woff" => "font/woff",
        "ttf" => "font/ttf",
        "otf" => "font/otf",
        "wasm" => "application/wasm",
        "txt" | "dic" | "aff" => "text/plain; charset=utf-8",
        _ => "application/octet-stream",
    }
}

pub async fn serve(uri: Uri) -> Response {
    let path = uri.path().trim_start_matches('/');
    if path.starts_with("api/") || path == "mcp" || path.starts_with("mcp/") {
        return StatusCode::NOT_FOUND.into_response();
    }
    let (name, file) = match WEB.get_file(path).filter(|_| !path.is_empty()) {
        Some(file) => (path, file),
        None => match WEB.get_file("index.html") {
            Some(file) => ("index.html", file),
            None => return StatusCode::NOT_FOUND.into_response(),
        },
    };
    ([(header::CONTENT_TYPE, content_type(name))], file.contents()).into_response()
}

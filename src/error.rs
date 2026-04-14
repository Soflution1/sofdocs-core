use thiserror::Error;

#[derive(Debug, Error)]
pub enum SofDocsError {
    #[error("Failed to open ZIP archive: {0}")]
    ZipError(#[from] zip::result::ZipError),

    #[error("XML parsing error: {0}")]
    XmlError(#[from] quick_xml::Error),

    #[error("XML attribute error: {0}")]
    XmlAttrError(#[from] quick_xml::events::attributes::AttrError),

    #[error("IO error: {0}")]
    IoError(#[from] std::io::Error),

    #[error("UTF-8 decoding error: {0}")]
    Utf8Error(#[from] std::str::Utf8Error),

    #[error("Missing required entry in DOCX archive: {0}")]
    MissingEntry(String),

    #[error("Unsupported format: {0}")]
    UnsupportedFormat(String),
}

pub type Result<T> = std::result::Result<T, SofDocsError>;

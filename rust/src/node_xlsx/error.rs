use std::num::{ParseFloatError, ParseIntError};

use rust_xlsxwriter::XlsxError;

pub struct NodeXlsxError {
    message: String,
}

impl NodeXlsxError {
    pub fn new(message: String) -> Self {
        Self { message }
    }
}

impl std::fmt::Display for NodeXlsxError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "{}", self.message)
    }
}

impl std::fmt::Debug for NodeXlsxError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "{}", self.message)
    }
}

impl std::error::Error for NodeXlsxError {}

impl From<XlsxError> for NodeXlsxError {
    fn from(error: XlsxError) -> Self {
        Self::new(error.to_string())
    }
}

impl From<ParseIntError> for NodeXlsxError {
    fn from(error: ParseIntError) -> Self {
        Self::new(error.to_string())
    }
}

impl From<ParseFloatError> for NodeXlsxError {
    fn from(error: ParseFloatError) -> Self {
        Self::new(error.to_string())
    }
}

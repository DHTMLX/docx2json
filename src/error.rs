use std::error::Error;
use std::fmt;

#[derive(Debug)]
pub struct DocError {
    message: String,
}

impl DocError {
    pub fn new(message: &str) -> DocError {
        DocError {
            message: message.to_string(),
        }
    }
}

impl fmt::Display for DocError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{}", self.message)
    }
}

impl Error for DocError {}

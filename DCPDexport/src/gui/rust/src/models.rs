use serde::{Deserialize, Serialize};

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct DCPDData {
    pub timestamp: String,
    pub value: f64,
    pub unit: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct AnalysisResult {
    pub data_points: usize,
    pub min: f64,
    pub max: f64,
    pub average: f64,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ProcessConfig {
    pub input_format: String,
    pub output_format: String,
    pub options: serde_json::Value,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct FileInfo {
    pub name: String,
    pub path: String,
    pub size: u64,
    pub created: String,
}

use serde::{Deserialize, Serialize};
use crate::models::AnalysisResult;
use anyhow::{Result, bail};
use chrono::NaiveDateTime;
use csv::Writer;
use log::info;
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};

#[derive(Debug, Serialize, Deserialize)]
pub struct CommandResult {
    pub success: bool,
    pub message: String,
    pub data: Option<serde_json::Value>,
}

/// 尝试用多种编码读取文件（和 console 共享逻辑）
fn read_file_with_encoding(path: &Path) -> Result<(String, Vec<String>)> {
    let encodings = [
        "GBK",
        "GB2312",
        "GB18030",
        "UTF-8",
        "UTF-16",
        "Latin1",
    ];

    for encoding_name in &encodings {
        if let Some(enc) = encoding_rs::Encoding::for_label(encoding_name.as_bytes()) {
            let data = fs::read(path)?;
            let (decoded, _, had_errors) = enc.decode(&data);
            if !had_errors {
                let lines: Vec<String> = decoded.lines().map(|s| s.to_string()).collect();
                info!("检测到编码: {}", encoding_name);
                return Ok((encoding_name.to_string(), lines));
            }
        }
    }

    // 尝试 UTF-8 作为兜底
    let data = fs::read(path)?;
    if let Some(enc) = encoding_rs::Encoding::for_label(b"utf-8") {
        let (decoded, _, _) = enc.decode(&data);
        let lines: Vec<String> = decoded.lines().map(|s| s.to_string()).collect();
        info!("使用 UTF-8 编码");
        Ok(("UTF-8".to_string(), lines))
    } else {
        bail!("无法读取文件");
    }
}

/// 解析源文件并按秒分组
fn parse_source(path: &Path) -> Result<(HashMap<i64, Vec<(f64, f64, f64, f64)>>, String)> {
    let (encoding, lines) = read_file_with_encoding(path)?;
    let mut data_by_second: HashMap<i64, Vec<(f64, f64, f64, f64)>> = HashMap::new();
    let mut first_ts: Option<NaiveDateTime> = None;

    for line in lines {
        let line = line.trim();
        if line.is_empty() || line.contains("二次平均") {
            continue;
        }
        let parts: Vec<&str> = line.split(',').collect();
        if parts.len() < 5 {
            continue;
        }
        let c1: f64 = match parts[0].trim().parse() {
            Ok(v) => v,
            Err(_) => continue,
        };
        let c2: f64 = match parts[1].trim().parse() {
            Ok(v) => v,
            Err(_) => continue,
        };
        let c3: f64 = match parts[2].trim().parse() {
            Ok(v) => v,
            Err(_) => continue,
        };
        let c4: f64 = match parts[3].trim().parse() {
            Ok(v) => v,
            Err(_) => continue,
        };
        let ts_str = parts[4].trim();
        let datetime = match NaiveDateTime::parse_from_str(ts_str, "%y%m%d%H%M%S") {
            Ok(dt) => dt,
            Err(_) => continue,
        };
        if first_ts.is_none() {
            first_ts = Some(datetime);
        }
        let first = first_ts.unwrap();
        let sec = (datetime - first).num_seconds();
        data_by_second.entry(sec).or_insert_with(Vec::new).push((c1, c2, c3, c4));
    }

    if data_by_second.is_empty() {
        bail!("未提取到有效数据");
    }

    Ok((data_by_second, encoding))
}

/// 生成唯一的输出文件名
fn unique_path(base: &Path) -> PathBuf {
    if !base.exists() {
        return base.to_path_buf();
    }
    let stem = base.file_stem().unwrap().to_string_lossy();
    let ext = base.extension().unwrap_or_default().to_string_lossy();
    let parent = base.parent().unwrap_or_else(|| Path::new("."));
    for i in 1..1000 {
        let new_name = format!("{}-{}.{}", stem, i, ext);
        let new_path = parent.join(&new_name);
        if !new_path.exists() {
            return new_path;
        }
    }
    base.to_path_buf()
}

/// 真正的处理流程，返回输出路径、编码、点数和实验时长
fn do_process(input: &Path, unit: &str) -> Result<(PathBuf, String, usize, f64)> {
    let (data_by_second, encoding) = parse_source(input)?;
    let now = chrono::Local::now().format("%y%m%d%H%M").to_string();
    let stem = input.file_stem().unwrap().to_string_lossy();
    let output_name = format!("{}_parsed_{}_{}.csv", stem, unit, now);
    let output_path = input.parent().unwrap_or_else(|| Path::new(".")).join(&output_name);
    let output_path = unique_path(&output_path);
    let mut sorted_seconds: Vec<i64> = data_by_second.keys().copied().collect();
    sorted_seconds.sort_unstable();
    let duration_secs = *sorted_seconds.last().unwrap_or(&0);
    let duration_value = if unit == "hr" { duration_secs as f64 / 3600.0 } else { duration_secs as f64 };
    let mut writer = Writer::from_path(&output_path)?;
    writer.write_record(&[unit, "V", "Vr", "a", "a/w"])?;
    let mut total_points = 0;
    for second in &sorted_seconds {
        let points = &data_by_second[second];
        let avg_c1 = points.iter().map(|p| p.0).sum::<f64>() / points.len() as f64;
        let avg_c2 = points.iter().map(|p| p.1).sum::<f64>() / points.len() as f64;
        let avg_c3 = points.iter().map(|p| p.2).sum::<f64>() / points.len() as f64;
        let avg_c4 = points.iter().map(|p| p.3).sum::<f64>() / points.len() as f64;
        let t_val = if unit == "hr" { *second as f64 / 3600.0 } else { *second as f64 };
        writer.write_record(&[
            format!("{:.6}", t_val),
            format!("{:.6}", avg_c1),
            format!("{:.6}", avg_c2),
            format!("{:.6}", avg_c3),
            format!("{:.6}", avg_c4),
        ])?;
        total_points += points.len();
    }
    writer.flush()?;
    info!("文件处理完成: {}", output_path.display());
    info!("输出编码: {}", encoding);
    info!("处理的数据点: {}", total_points);
    info!("实验时长: {} {}", duration_value, unit);
    Ok((output_path, encoding, total_points, duration_value))
}

/// 分析DCPD数据
#[tauri::command]
pub fn analyze_dcpd_data(file_path: String) -> CommandResult {
    // 保留原来的 stub
    let result = AnalysisResult {
        data_points: 1024,
        min: 0.0,
        max: 100.0,
        average: 50.5,
    };
    CommandResult {
        success: true,
        message: format!("分析完成: {}", file_path),
        data: Some(serde_json::to_value(result).unwrap_or_default()),
    }
}

/// 处理文件命令
#[tauri::command]
pub fn process_file(input_path: String, time_unit: String) -> CommandResult {
    let input = Path::new(&input_path);
    match do_process(input, &time_unit) {
        Ok((output_path, encoding, points, duration)) => {
            let mut resp = serde_json::Map::new();
            resp.insert("output_path".to_string(), serde_json::json!(output_path.display().to_string()));
            resp.insert("encoding".to_string(), serde_json::json!(encoding));
            resp.insert("points".to_string(), serde_json::json!(points));
            resp.insert("duration".to_string(), serde_json::json!(duration));
            CommandResult {
                success: true,
                message: format!("处理完成: {}", input_path),
                data: Some(serde_json::Value::Object(resp)),
            }
        }
        Err(e) => CommandResult {
            success: false,
            message: format!("处理失败: {}", e),
            data: None,
        },
    }
}

/// 获取系统状态
#[tauri::command]
pub fn get_status() -> CommandResult {
    let info = format!(
        "DCPD GUI v0.1.0 | OS: {} | 系统正常运行",
        std::env::consts::OS
    );

    CommandResult {
        success: true,
        message: info,
        data: Some(serde_json::json!({
            "app": "DCPD GUI",
            "version": "0.1.0",
            "os": std::env::consts::OS,
            "timestamp": chrono::Local::now().to_rfc3339(),
        })),
    }
}

/// 获取使用示例
#[tauri::command]
pub fn get_examples() -> CommandResult {
    let examples = vec![
        "1. 分析数据: 选择CSV/DAT文件 -> 点击「分析」",
        "2. 处理数据: 输入输出路径 -> 选择处理选项 -> 点击「处理」",
        "3. 查看状态: 点击右上角「状态」按钮",
        "4. 批量处理: 支持拖拽多个文件",
    ];

    CommandResult {
        success: true,
        message: "使用示例".to_string(),
        data: Some(serde_json::json!(examples)),
    }
}
 

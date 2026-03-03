use anyhow::{Result, bail};
use chrono::NaiveDateTime;
use csv::Writer;
use log::info;
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};

/// 尝试用多种编码读取文件
fn read_file_with_encoding(path: &Path) -> Result<(String, Vec<String>)> {
    // 尝试多种常见编码
    let encodings = [
        "GBK",
        "GB2312",
        "GB18030",
        "UTF-8",
        "UTF-16",
        "Latin1",
    ];

    for encoding_name in &encodings {
        match encoding_rs::Encoding::for_label(encoding_name.as_bytes()) {
            Some(enc) => {
                let data = fs::read(path)?;
                let (decoded, _, had_errors) = enc.decode(&data);
                if !had_errors {
                    let lines: Vec<String> = decoded.lines().map(|s| s.to_string()).collect();
                    info!("检测到编码: {}", encoding_name);
                    return Ok((encoding_name.to_string(), lines));
                }
            }
            None => continue,
        }
    }

    // 如果所有编码都试过了，归咎于 UTF-8
    let data = fs::read(path)?;
    match encoding_rs::Encoding::for_label(b"utf-8") {
        Some(enc) => {
            let (decoded, _, _) = enc.decode(&data);
            let lines: Vec<String> = decoded.lines().map(|s| s.to_string()).collect();
            info!("使用 UTF-8 编码");
            Ok(("UTF-8".to_string(), lines))
        }
        None => anyhow::bail!("无法读取文件"),
    }
}


/// 解析DCPD数据文件
fn parse_source(path: &Path) -> Result<(HashMap<i64, Vec<(f64, f64, f64, f64)>>, String)> {
    let (encoding, lines) = read_file_with_encoding(path)?;
    
    let mut data_by_second: HashMap<i64, Vec<(f64, f64, f64, f64)>> = HashMap::new();
    let mut first_ts: Option<NaiveDateTime> = None;

    for line in lines {
        let line = line.trim();
        
        // 跳过空行和特殊标记
        if line.is_empty() || line.contains("二次平均") {
            continue;
        }

        let parts: Vec<&str> = line.split(',').collect();
        
        // 需要至少5列
        if parts.len() < 5 {
            continue;
        }

        // 尝试解析前4列为浮点数和第5列为时间戳
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

        // 解析时间戳: yymmddHHMMSS
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

/// 处理DCPD数据文件
pub fn process_file(input_path: &Path, unit: &str) -> Result<(PathBuf, String, usize, f64)> {
    info!("开始处理文件: {}", input_path.display());

    // 解析源文件
    let (data_by_second, encoding) = parse_source(input_path)?;

    // 生成输出文件名
    let now = chrono::Local::now().format("%y%m%d%H%M").to_string();
    let stem = input_path.file_stem().unwrap().to_string_lossy();
    let output_name = format!("{}_parsed_{}_{}.csv", stem, unit, now);
    let output_path = input_path.parent().unwrap_or_else(|| Path::new(".")).join(&output_name);
    let output_path = unique_path(&output_path);

    // 排序秒数
    let mut sorted_seconds: Vec<i64> = data_by_second.keys().copied().collect();
    sorted_seconds.sort_unstable();

    // 计算实验时长（以秒为单位）
    let duration_secs = *sorted_seconds.last().unwrap_or(&0);
    let duration_value = if unit == "hr" {
        duration_secs as f64 / 3600.0
    } else {
        duration_secs as f64
    };

    // 写入CSV文件
    let mut writer = Writer::from_path(&output_path)?;
    
    // 写入表头
    writer.write_record(&[unit, "V", "Vr", "a", "a/w"])?;

    // 写入数据行
    let mut total_points = 0;
    for second in &sorted_seconds {
        let points = &data_by_second[second];
        
        // 计算平均值
        let avg_c1 = points.iter().map(|p| p.0).sum::<f64>() / points.len() as f64;
        let avg_c2 = points.iter().map(|p| p.1).sum::<f64>() / points.len() as f64;
        let avg_c3 = points.iter().map(|p| p.2).sum::<f64>() / points.len() as f64;
        let avg_c4 = points.iter().map(|p| p.3).sum::<f64>() / points.len() as f64;

        // 转换时间值
        let t_val = if unit == "hr" {
            *second as f64 / 3600.0
        } else {
            *second as f64
        };

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

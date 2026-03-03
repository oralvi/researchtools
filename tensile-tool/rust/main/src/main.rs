use anyhow::{anyhow, Context, Result};
use calamine::{open_workbook_auto, Data, Reader};
use chrono::Local;
use csv::ReaderBuilder;
use encoding_rs::{GBK, UTF_8, WINDOWS_1252};
use regex::Regex;
use rust_xlsxwriter::Workbook;
use std::collections::{HashMap, HashSet};
use std::fs;
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::thread;
use std::time::Duration;

const VERSION: &str = "V1.6-rust";
const XLSX_MAX_ROWS: usize = 1_048_576;
const RESULT_HEADER_ROWS: usize = 2;

#[derive(Default)]
struct AppState {
    logs: Vec<String>,
    pre_artifact: Option<PathBuf>,
    result_artifact: Option<PathBuf>,
}

#[derive(Clone)]
struct Specimen {
    name: String,
    col_start: usize,
}

#[derive(Clone)]
struct Params {
    gauge: f64,
    area: f64,
}

struct RawSpecimen {
    load: Vec<String>,
    stroke: Vec<String>,
    ext: Vec<String>,
}

#[derive(Clone)]
struct SpecimenData {
    load_kn: Vec<f64>,
    stroke_mm: Vec<f64>,
    ext_corr: Vec<f64>,
    stress: Vec<f64>,
    strain: Vec<f64>,
    strain_stroke: Vec<f64>,
}

#[derive(Clone)]
struct MechProps {
    e_gpa: f64,
    rp02: f64,
    rm: f64,
    a: f64,
    fit_tag: String,
}

fn log(state: &mut AppState, msg: &str) {
    println!("{}", msg);
    state.logs.push(msg.to_string());
}

fn prompt(msg: &str) -> String {
    print!("{}", msg);
    let _ = io::stdout().flush();
    let mut s = String::new();
    if io::stdin().read_line(&mut s).is_ok() {
        s.trim().to_string()
    } else {
        String::new()
    }
}

fn safe_parent_dir(path: &Path) -> PathBuf {
    if let Some(parent) = path.parent() {
        if !parent.as_os_str().is_empty() {
            return parent.to_path_buf();
        }
    }
    std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."))
}

fn progress_bar(cur: usize, total: usize, prefix: &str) {
    let total = total.max(1);
    let cur = cur.min(total);
    let width = 20;
    let fill = width * cur / total;
    let bar = format!("{}{}", ">".repeat(fill), ".".repeat(width - fill));
    println!("[INFO] {}[{}] {}/{}", prefix, bar, cur, total);
}

fn parse_f64_or_nan(s: &str) -> f64 {
    let t = s.trim();
    if t.is_empty() || t == "-.----" || t == "-" || t.eq_ignore_ascii_case("nan") {
        f64::NAN
    } else {
        t.parse::<f64>().unwrap_or(f64::NAN)
    }
}

fn parse_positive_optional(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    let v = t.parse::<f64>().ok()?;
    if v > 0.0 { Some(v) } else { None }
}

fn app_dir() -> PathBuf {
    std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|x| x.to_path_buf()))
        .unwrap_or_else(|| std::env::current_dir().unwrap_or_else(|_| PathBuf::from(".")))
}

fn write_error_log(dir: &Path, state: &AppState, err: &anyhow::Error) -> Result<PathBuf> {
    let p = dir.join("error.log");
    let mut txt = String::new();
    txt.push_str(&format!("version: {}\n", VERSION));
    txt.push_str(&format!("time: {}\n", Local::now().format("%Y-%m-%d %H:%M:%S")));
    txt.push_str(&format!("error: {:?}\n\n", err));
    txt.push_str("full_log:\n\n");
    for line in &state.logs {
        txt.push_str(line);
        txt.push('\n');
    }
    fs::write(&p, txt)?;
    Ok(p)
}

fn cleanup_artifacts(state: &mut AppState) {
    for (name, p) in [
        ("预处理", state.pre_artifact.clone()),
        ("result", state.result_artifact.clone()),
    ] {
        if let Some(path) = p {
            if path.exists() {
                match fs::remove_file(&path) {
                    Ok(_) => log(state, &format!("[INFO] 已删除{}文件: {}", name, path.display())),
                    Err(_) => {
                        log(state, &format!("[WARN] {}文件删除失败（可能被占用）: {}", name, path.display()));
                        log(state, "[WARN] 请关闭对应Excel文件后手动删除");
                    }
                }
            }
        }
    }
}

fn cleanup_temp_files(state: &mut AppState) {
    let cwd = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    let entries = match fs::read_dir(&cwd) {
        Ok(v) => v,
        Err(_) => return,
    };
    for e in entries.flatten() {
        let p = e.path();
        if !p.is_file() {
            continue;
        }
        let name = match p.file_name().and_then(|s| s.to_str()) {
            Some(v) => v.to_string(),
            None => continue,
        };
        let low = name.to_lowercase();
        let is_pre = low.starts_with("pre-") && low.ends_with(".xlsx");
        let is_tpl = low.starts_with("specimen_params_template") && low.ends_with(".xlsx");
        let is_tmp = low.ends_with(".tmp") || low.ends_with(".temp");
        if !(is_pre || is_tpl || is_tmp) {
            continue;
        }
        match fs::remove_file(&p) {
            Ok(_) => log(state, &format!("[INFO] 已删除临时文件: {}", p.display())),
            Err(_) => {
                log(state, &format!("[WARN] 临时文件删除失败（可能被占用）: {}", p.display()));
                log(state, "[WARN] 请关闭对应Excel文件后手动删除");
            }
        }
    }
}

fn read_raw_csv(path: &Path, _state: &mut AppState) -> Result<Vec<Vec<String>>> {
    let bytes = fs::read(path).with_context(|| format!("读取失败: {}", path.display()))?;
    let encs = ["gbk", "gb18030", "utf-8-sig", "utf-8", "latin-1"];

    for enc in encs {
        let text = match enc {
            "gbk" | "gb18030" => {
                let (c, _, bad) = GBK.decode(&bytes);
                if bad { continue; }
                c.into_owned()
            }
            "utf-8-sig" => {
                let (c, _, bad) = UTF_8.decode(&bytes);
                if bad { continue; }
                c.trim_start_matches('\u{feff}').to_string()
            }
            "utf-8" => {
                let (c, _, bad) = UTF_8.decode(&bytes);
                if bad { continue; }
                c.into_owned()
            }
            "latin-1" => {
                let (c, _, _) = WINDOWS_1252.decode(&bytes);
                c.into_owned()
            }
            _ => continue,
        };

        let mut out = Vec::new();
        let mut rdr = ReaderBuilder::new().has_headers(false).from_reader(text.as_bytes());
        for rec in rdr.records() {
            let rec = rec?;
            let row: Vec<String> = rec.iter().map(|c| c.trim_matches('"').trim().to_string()).collect();
            if row.iter().any(|c| !c.is_empty()) {
                out.push(row);
            }
        }
        return Ok(out);
    }
    Err(anyhow!("无法解码文件：{}", path.display()))
}

fn detect_specimens(rows: &[Vec<String>]) -> Vec<Specimen> {
    if rows.is_empty() { return vec![]; }
    let row0 = &rows[0];
    let mut out = Vec::new();
    let mut c = 0usize;
    while c < row0.len() {
        let name = row0[c].trim();
        if !name.is_empty() && name.to_lowercase() != "nan" {
            out.push(Specimen { name: name.to_string(), col_start: c });
            c += 4;
        } else {
            c += 1;
        }
    }
    out
}

fn extract_specimen(rows: &[Vec<String>], col: usize) -> RawSpecimen {
    let mut load = Vec::new();
    let mut stroke = Vec::new();
    let mut ext = Vec::new();
    for r in rows.iter().skip(3) {
        load.push(r.get(col + 1).cloned().unwrap_or_default());
        stroke.push(r.get(col + 2).cloned().unwrap_or_default());
        ext.push(r.get(col + 3).cloned().unwrap_or_default());
    }
    RawSpecimen { load, stroke, ext }
}

fn find_csv_files(start: &Path, state: &mut AppState) -> Result<Vec<PathBuf>> {
    let dir = safe_parent_dir(start);
    let stem = start.file_stem().and_then(|s| s.to_str()).unwrap_or("");
    let ext = start.extension().and_then(|s| s.to_str()).unwrap_or("csv");
    let re = Regex::new(r"^(.+)-(\d+)$")?;
    if let Some(cap) = re.captures(stem) {
        let prefix = cap.get(1).unwrap().as_str();
        let pat = Regex::new(&format!(r"^{}-(\d+)\.{}$", regex::escape(prefix), regex::escape(ext)))?;
        let mut found: Vec<(usize, PathBuf)> = Vec::new();
        for e in fs::read_dir(&dir)? {
            let e = e?;
            let n = e.file_name().to_string_lossy().to_string();
            if let Some(c) = pat.captures(&n) {
                let idx = c.get(1).unwrap().as_str().parse::<usize>().unwrap_or(0);
                found.push((idx, e.path()));
            }
        }
        if found.is_empty() {
            return Err(anyhow!("未找到 {}-*.{}", prefix, ext));
        }
        found.sort_by_key(|v| v.0);
        let files: Vec<PathBuf> = found.into_iter().map(|v| v.1).collect();
        let mut names: Vec<String> = files
            .iter()
            .filter_map(|p| p.file_name().and_then(|s| s.to_str()).map(|s| s.to_string()))
            .collect();
        let preview = if names.len() > 10 {
            let rest = names.len() - 10;
            names.truncate(10);
            format!("{} ...(+{} 个)", names.join(", "), rest)
        } else {
            names.join(", ")
        };
        log(state, &format!("\n  发现 {} 个分段文件：{}", files.len(), preview));
        Ok(files)
    } else {
        Ok(vec![start.to_path_buf()])
    }
}

fn load_and_merge(files: &[PathBuf], state: &mut AppState) -> Result<(Vec<Vec<String>>, Vec<Specimen>)> {
    let total = files.len().max(1);
    let mut last_stage = 0usize;
    progress_bar(0, 100, "读取合并 ");
    let first = read_raw_csv(&files[0], state)?;
    let mut stage = (100usize / total) / 25 * 25;
    if stage > last_stage {
        last_stage = stage;
        progress_bar(last_stage, 100, "读取合并 ");
    }
    let specs = detect_specimens(&first);
    if files.len() > 10 || (specs.len() > 5 && files.len() >= 5) {
        log(state, "[INFO] 正在读取超大数据集，请耐心等待...");
    }
    if files.len() == 1 {
        return Ok((first, specs));
    }
    let mut merged = first.clone();
    for (i, f) in files.iter().enumerate().skip(1) {
        let rows = read_raw_csv(f, state)?;
        merged.extend(rows.into_iter().skip(3));
        stage = (((i + 1) * 100) / total) / 25 * 25;
        if stage > last_stage {
            last_stage = stage;
            progress_bar(last_stage, 100, "读取合并 ");
        }
    }
    if last_stage < 100 {
        progress_bar(100, 100, "读取合并 ");
    }
    Ok((merged, specs))
}

fn cell_to_text(cell: Option<&Data>) -> String {
    match cell {
        Some(Data::String(s)) => s.trim().to_string(),
        Some(Data::Float(v)) => format!("{}", v),
        Some(Data::Int(v)) => format!("{}", v),
        Some(Data::Bool(v)) => format!("{}", v),
        _ => String::new(),
    }
}

fn safe_path(desired: &Path, state: &mut AppState) -> Result<PathBuf> {
    if !desired.exists() {
        return Ok(desired.to_path_buf());
    }
    let ans = prompt(&format!("[WARN] 已存在：{}，覆盖？[y/N]: ", desired.display()));
    if ans.eq_ignore_ascii_case("y") {
        return Ok(desired.to_path_buf());
    }

    let parent = desired.parent().unwrap_or_else(|| Path::new("."));
    let ext = desired.extension().and_then(|s| s.to_str()).unwrap_or("xlsx");
    let base = desired.file_stem().and_then(|s| s.to_str()).unwrap_or("out");
    let re = Regex::new(r"_\d+$")?;
    let stem = re.replace(base, "").to_string();

    let mut i = 1usize;
    loop {
        let p = parent.join(format!("{}_{}.{}", stem, i, ext));
        if !p.exists() {
            log(state, &format!("  -> 另存为：{}", p.display()));
            return Ok(p);
        }
        i += 1;
    }
}

fn write_safe<F>(mut f: F, desired: &Path, state: &mut AppState) -> Result<PathBuf>
where
    F: FnMut(&Path) -> Result<()>,
{
    let mut path = safe_path(desired, state)?;
    loop {
        match f(&path) {
            Ok(_) => return Ok(path),
            Err(e) => {
                let msg = format!("{:?}", e);
                if msg.contains("Permission denied") {
                    path = safe_path(&path, state)?;
                    continue;
                }
                return Err(e);
            }
        }
    }
}

fn gen_template(specimens: &[Specimen], path: &Path, state: &mut AppState) -> Result<PathBuf> {
    let writer = |target: &Path| -> Result<()> {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        ws.set_name("试样参数")?;
        for (i, h) in ["试样名称", "宽度/直径 (mm)", "厚度 (mm)", "标距 (mm)", "是否圆柱棒(是/否)"]
            .iter()
            .enumerate()
        {
            ws.write(0, i as u16, *h)?;
        }
        for (ri, sp) in specimens.iter().enumerate() {
            let r = (ri + 2) as u32;
            ws.write(r, 0, sp.name.as_str())?;
            ws.write(r, 4, "否")?;
        }
        wb.save(target)?;
        Ok(())
    };
    write_safe(writer, path, state)
}

fn load_template(path: &Path, specimens: &[Specimen], state: &mut AppState) -> Result<HashMap<String, Params>> {
    let mut wb = open_workbook_auto(path).with_context(|| format!("打开模板失败: {}", path.display()))?;
    let range = wb
        .worksheet_range_at(0)
        .ok_or_else(|| anyhow!("模板中没有工作表"))??;

    let mut params = HashMap::new();
    let names: Vec<String> = specimens.iter().map(|s| s.name.clone()).collect();
    let name_set: HashSet<String> = names.iter().cloned().collect();
    let total = names.len();
    let mut loaded = 0usize;
    log(state, "\n  参数文件读取中...");

    for (row_idx, row) in range.rows().enumerate().skip(2) {
        let name = cell_to_text(row.get(0));
        if name.is_empty() || !name_set.contains(&name) || params.contains_key(&name) {
            continue;
        }
        let gauge_txt = cell_to_text(row.get(3));
        let gauge = gauge_txt
            .parse::<f64>()
            .map_err(|_| anyhow!("第{}行 {}: 标距未填写或格式错误", row_idx + 1, name))?;
        let w = parse_positive_optional(&cell_to_text(row.get(1)));
        let t = parse_positive_optional(&cell_to_text(row.get(2)));
        let flag = matches!(
            cell_to_text(row.get(4)).to_lowercase().as_str(),
            "是" | "yes" | "y" | "1" | "true"
        );

        let (width, thickness, cyl) = match (w, t) {
            (None, None) => return Err(anyhow!("第{}行 {}: 宽度/厚度不能同时为空或<=0", row_idx + 1, name)),
            (Some(wv), Some(tv)) => {
                if flag { (wv, wv, true) } else { (wv, tv, false) }
            }
            (Some(d), None) | (None, Some(d)) => {
                log(state, &format!("[INFO] {} 检测到单列直径输入（或另一列为0/空），按圆柱棒处理", name));
                (d, d, true)
            }
        };

        let (area, desc) = calc_area(width, thickness, cyl);
        if area <= 0.0 {
            return Err(anyhow!("第{}行 {}: 截面积<=0", row_idx + 1, name));
        }
        log(state, &format!("    {}：{}，标距={:.4} mm", name, desc, gauge));
        params.insert(name.clone(), Params { gauge, area });
        loaded += 1;
        progress_bar(loaded, total, "读取参数 ");
        if loaded == total {
            break;
        }
    }

    let miss: Vec<String> = names.into_iter().filter(|n| !params.contains_key(n)).collect();
    if !miss.is_empty() {
        return Err(anyhow!("模板中缺少试样参数: {:?}", miss));
    }
    Ok(params)
}

fn calc_area(width: f64, thickness: f64, cyl: bool) -> (f64, String) {
    if cyl {
        let a = width * thickness * std::f64::consts::PI / 4.0;
        (a, format!("圆柱棒 π/4×{:.4}²={:.4} mm²", width, a))
    } else {
        let a = width * thickness;
        (a, format!("矩形 {:.4}×{:.4}={:.4} mm²", width, thickness, a))
    }
}

fn input_one(name: &str, state: &mut AppState) -> Result<Params> {
    log(state, &format!("[INFO] 试样参数: {}", name));
    loop {
        let w_txt = prompt("宽度/直径 (mm，可留空): ");
        let t_txt = prompt("厚度 (mm，可留空): ");
        let g_txt = prompt("标距 (mm): ");
        let gauge = match g_txt.parse::<f64>() {
            Ok(v) => v,
            Err(_) => {
                println!("[ERR] 标距必须是数字");
                continue;
            }
        };
        let w = parse_positive_optional(&w_txt);
        let t = parse_positive_optional(&t_txt);
        let (width, thickness, cyl) = match (w, t) {
            (None, None) => {
                println!("[ERR] 宽度/厚度不能同时为空或<=0");
                continue;
            }
            (Some(wv), Some(tv)) => (wv, tv, false),
            (Some(d), None) | (None, Some(d)) => {
                println!("[INFO] 检测到单列直径输入，按圆柱棒处理");
                (d, d, true)
            }
        };
        let (area, desc) = calc_area(width, thickness, cyl);
        if area <= 0.0 {
            println!("[ERR] 截面积<=0");
            continue;
        }
        log(state, &format!("[INFO] {}，标距={:.4} mm", desc, gauge));
        return Ok(Params { gauge, area });
    }
}

fn get_all_params(specimens: &[Specimen], state: &mut AppState) -> Result<HashMap<String, Params>> {
    if specimens.len() == 1 {
        let p = input_one(&specimens[0].name, state)?;
        let mut m = HashMap::new();
        m.insert(specimens[0].name.clone(), p);
        return Ok(m);
    }

    println!("[INFO] 多试样参数模式，共 {} 个试样", specimens.len());
    println!("[INFO] [1] 生成模板后填写导入（推荐）");
    println!("[INFO] [2] 指定已有参数文件");
    println!("[INFO] [3] 逐个手动输入");
    let c = {
        let s = prompt("选择 [1/2/3]，默认1: ");
        if s.is_empty() { "1".to_string() } else { s }
    };

    if c == "3" {
        let mut out = HashMap::new();
        for sp in specimens {
            out.insert(sp.name.clone(), input_one(&sp.name, state)?);
        }
        return Ok(out);
    }

    if c == "2" {
        loop {
            let p = prompt("参数文件路径（可直接拖入）: ");
            if p.is_empty() {
                println!("[ERR] 未填写内容，请输入参数文件路径");
                continue;
            }
            let path = PathBuf::from(p.trim_matches('"'));
            if !path.exists() {
                println!("[ERR] 文件不存在");
                continue;
            }
            match load_template(&path, specimens, state) {
                Ok(v) => return Ok(v),
                Err(e) => {
                    println!("[ERR] {}", e);
                    if prompt("[INFO] 回车重选，输入 q 退出: ").eq_ignore_ascii_case("q") {
                        return Err(anyhow!("用户取消参数输入"));
                    }
                }
            }
        }
    }

    let tpl_default = std::env::current_dir()?.join("specimen_params_template.xlsx");
    let tpl = gen_template(specimens, &tpl_default, state)?;
    loop {
        println!("[INFO] 参数模板: {}", tpl.display());
        println!("[INFO] 请填写后保存，再按回车继续。");
        let _ = prompt("");
        match load_template(&tpl, specimens, state) {
            Ok(v) => {
                if fs::remove_file(&tpl).is_ok() {
                    log(state, &format!("[INFO] 已删除参数模板: {}", tpl.display()));
                } else {
                    log(state, &format!("[WARN] 参数模板删除失败（可能被Excel占用）: {}", tpl.display()));
                    log(state, "[WARN] 请关闭Excel后手动删除");
                }
                return Ok(v);
            }
            Err(e) => {
                println!("[ERR] {}", e);
                if prompt("[INFO] 回车继续修改模板，输入 q 退出: ").eq_ignore_ascii_case("q") {
                    return Err(anyhow!("用户取消参数输入"));
                }
            }
        }
    }
}

fn process_extensometer(raw: &RawSpecimen, gauge: f64, sp_name: &str) -> Result<(Vec<f64>, f64)> {
    println!("\n  [引伸计] {}", sp_name);
    let ext: Vec<f64> = raw.ext.iter().map(|v| parse_f64_or_nan(v)).collect();
    if !ext.iter().any(|v| !v.is_nan() && *v != 0.0) {
        if !prompt("[WARN] 引伸计无效，继续？[y/N]: ").eq_ignore_ascii_case("y") {
            return Err(anyhow!("{} 引伸计无效且用户取消处理", sp_name));
        }
        return Ok((vec![f64::NAN; ext.len()], 10.0));
    }
    let mut corrected = ext.clone();
    let mut cum = 0.0;
    let mut prev: Option<f64> = None;
    for (i, v) in ext.iter().enumerate() {
        if v.is_nan() { continue; }
        if let Some(p) = prev {
            if v - p < -0.05 {
                cum += p - v;
                println!("[WARN] 滑脱 第{}行：{:.4}->{:.4} 偏移={:.4}", i + 1, p, v, cum);
            }
        }
        corrected[i] = v + cum;
        prev = Some(*v);
    }

    let dash_idx = raw.ext.iter().position(|s| s.trim() == "-.----");
    let mut ext_mm = 10.0_f64;
    let ext_ranges = [(10.0, 0.2), (15.0, 0.3), (25.0, 0.5)];
    if let Some(di) = dash_idx {
        if di > 0 {
            let mut last_i: Option<usize> = None;
            for i in 0..di {
                if !corrected[i].is_nan() {
                    last_i = Some(i);
                }
            }
            if let Some(li) = last_i {
                let last_val = corrected[li];
                let mut best = ext_ranges[0];
                let mut best_diff = (last_val - best.1).abs();
                for r in ext_ranges.iter().skip(1) {
                    let d = (last_val - r.1).abs();
                    if d < best_diff {
                        best = *r;
                        best_diff = d;
                    }
                }
                ext_mm = best.0;
                println!("  -> {:.0}mm 引伸计（末值={:.4}，满量程≈{:.3}）", ext_mm, last_val, best.1);
                let last_stroke = raw.stroke.get(li).map(|s| parse_f64_or_nan(s)).unwrap_or(0.0);
                println!("  -> 行程替代：(行程-{:.4})/{:.4}×{:.0}+{:.4}", last_stroke, gauge, ext_mm, last_val);
                for i in di..corrected.len() {
                    if corrected[i].is_nan() {
                        let s = raw.stroke.get(i).map(|v| parse_f64_or_nan(v)).unwrap_or(f64::NAN);
                        if !s.is_nan() {
                            corrected[i] = (s - last_stroke) / gauge * ext_mm + last_val;
                        }
                    }
                }
            }
        }
    }

    Ok((corrected, ext_mm))
}

fn calc_data(raw: RawSpecimen, ext_corr: Vec<f64>, ext_mm: f64, area: f64, gauge: f64) -> SpecimenData {
    let load_kn: Vec<f64> = raw.load.iter().map(|s| parse_f64_or_nan(s)).collect();
    let stroke_mm: Vec<f64> = raw.stroke.iter().map(|s| parse_f64_or_nan(s)).collect();
    let stress: Vec<f64> = load_kn.iter().map(|v| if v.is_nan() { f64::NAN } else { v / area * 1000.0 }).collect();
    let strain: Vec<f64> = ext_corr.iter().map(|v| if v.is_nan() { f64::NAN } else { v / ext_mm * 100.0 }).collect();
    let first_stroke = stroke_mm.iter().copied().find(|v| !v.is_nan()).unwrap_or(f64::NAN);
    let strain_stroke: Vec<f64> = stroke_mm
        .iter()
        .map(|v| {
            if v.is_nan() || first_stroke.is_nan() || gauge <= 0.0 {
                f64::NAN
            } else {
                (v - first_stroke) / gauge * 100.0
            }
        })
        .collect();
    SpecimenData { load_kn, stroke_mm, ext_corr, stress, strain, strain_stroke }
}

fn linear_fit_r2(x: &[f64], y: &[f64]) -> Option<(f64, f64, f64)> {
    if x.len() < 2 || x.len() != y.len() {
        return None;
    }
    let n = x.len() as f64;
    let sx: f64 = x.iter().sum();
    let sy: f64 = y.iter().sum();
    let sxx: f64 = x.iter().map(|v| v * v).sum();
    let sxy: f64 = x.iter().zip(y.iter()).map(|(a, b)| a * b).sum();
    let denom = n * sxx - sx * sx;
    if denom.abs() < 1e-12 {
        return None;
    }
    let slope = (n * sxy - sx * sy) / denom;
    let intercept = (sy - slope * sx) / n;
    let y_mean = sy / n;
    let mut sst = 0.0;
    let mut sse = 0.0;
    for (xi, yi) in x.iter().zip(y.iter()) {
        let fit = slope * xi + intercept;
        sst += (yi - y_mean) * (yi - y_mean);
        sse += (yi - fit) * (yi - fit);
    }
    let r2 = if sst > 1e-12 { 1.0 - sse / sst } else { 1.0 };
    Some((slope, intercept, r2))
}

fn calc_mech_props(data: &SpecimenData) -> MechProps {
    let mut stress: Vec<f64> = data.stress.iter().copied().filter(|v| !v.is_nan()).collect();
    let mut strain: Vec<f64> = data.strain.iter().copied().filter(|v| !v.is_nan()).collect();
    let n = stress.len().min(strain.len());
    stress.truncate(n);
    strain.truncate(n);
    if n < 10 {
        return MechProps {
            e_gpa: f64::NAN,
            rp02: f64::NAN,
            rm: f64::NAN,
            a: f64::NAN,
            fit_tag: "仅供参考".to_string(),
        };
    }

    let rm = stress.iter().copied().fold(f64::MIN, f64::max);
    let a_ext = data
        .strain
        .iter()
        .rev()
        .copied()
        .find(|v| !v.is_nan())
        .unwrap_or(f64::NAN);
    let a = if a_ext.is_finite() {
        // 优先使用引伸计标定后的应变终值
        a_ext
    } else {
        let a_stroke = data
            .strain_stroke
            .iter()
            .rev()
            .copied()
            .find(|v| !v.is_nan())
            .unwrap_or(f64::NAN);
        if a_stroke.is_finite() {
            println!("[INFO] 引伸计不可用，A按行程终值计算");
        }
        a_stroke
    };
    let sd: Vec<f64> = strain.iter().map(|v| v / 100.0).collect();

    let mut e_gpa = f64::NAN;
    let mut rp02 = f64::NAN;
    let mut fit_tag = "仅供参考".to_string();

    'outer: for strain_cap in [0.010_f64, 0.020_f64] {
        let idx: Vec<usize> = (0..n)
            .filter(|&i| stress[i] >= 0.05 * rm && stress[i] <= 0.70 * rm && sd[i] <= strain_cap)
            .collect();
        if idx.len() < 8 {
            continue;
        }
        let idx = if idx.len() > 5000 {
            let step = idx.len().div_ceil(5000);
            println!("[INFO] 弹性拟合点过多（{}），按步长{}抽样加速", idx.len(), step);
            idx.into_iter().step_by(step).collect::<Vec<_>>()
        } else {
            idx
        };
        let mut best: Option<(f64, f64, f64, usize)> = None;
        for end in 8..=idx.len() {
            let x: Vec<f64> = idx[..end].iter().map(|&i| sd[i]).collect();
            let y: Vec<f64> = idx[..end].iter().map(|&i| stress[i]).collect();
            if let Some((slope, intercept, r2)) = linear_fit_r2(&x, &y) {
                match best {
                    None => best = Some((slope, intercept, r2, end)),
                    Some((_, _, br2, bend)) => {
                        if (r2, -(end as isize)) > (br2, -(bend as isize)) {
                            best = Some((slope, intercept, r2, end));
                        }
                    }
                }
            }
        }

        if let Some((slope, intercept, r2, end)) = best {
            e_gpa = slope / 1000.0;
            fit_tag = if r2 >= 0.9995 { "精确拟合".to_string() } else { "仅供参考".to_string() };
            println!("  -> E={:.2} GPa（{} 点，R2={:.6}，{}）", e_gpa, end, r2, fit_tag);

            let fit_end_idx = idx[end - 1];
            let fit_end_strain = sd[fit_end_idx];
            let mut diff = Vec::with_capacity(n);
            for i in 0..n {
                let offset = slope * (sd[i] - 0.002) + intercept;
                diff.push(stress[i] - offset);
            }
            let mut cross: Option<usize> = None;
            for i in 1..n {
                if sd[i] <= fit_end_strain { continue; }
                if diff[i - 1] < 0.0 && diff[i] >= 0.0 {
                    cross = Some(i);
                    break;
                }
            }
            if let Some(i) = cross {
                let d0 = diff[i - 1];
                let d1 = diff[i];
                let t = if (d1 - d0).abs() > 1e-12 { -d0 / (d1 - d0) } else { 0.0 };
                rp02 = stress[i - 1] + t * (stress[i] - stress[i - 1]);
            } else {
                let mut best_j = 0usize;
                let mut best_abs = f64::MAX;
                for i in 0..n {
                    if sd[i] < fit_end_strain { continue; }
                    let ad = diff[i].abs();
                    if ad < best_abs {
                        best_abs = ad;
                        best_j = i;
                    }
                }
                rp02 = stress[best_j];
            }
            break 'outer;
        }
    }

    MechProps { e_gpa, rp02, rm, a, fit_tag }
}

fn fmt_val(v: f64, unit: &str) -> String {
    if v.is_nan() { "N/A".to_string() } else { format!("{:.3}{}", v, unit) }
}

fn write_result_xlsx(results: &[(String, SpecimenData)], prefix: &str, state: &mut AppState) -> Result<PathBuf> {
    let mut mech: Vec<(String, MechProps)> = Vec::new();
    let logs_snapshot = state.logs.clone();
    log(state, "\n  各试样力学性能：");
    for (name, d) in results {
        if d.stress.len() > 100_000 {
            log(state, &format!("[INFO] {} 数据点较多，正在计算力学性能，请稍候...", name));
        }
        let p = calc_mech_props(d);
        log(
            state,
            &format!(
                "    {}：Rp0.2={}  Rm={}  A={}  E={}（{}）",
                name,
                fmt_val(p.rp02, " MPa"),
                fmt_val(p.rm, " MPa"),
                fmt_val(p.a, " %"),
                fmt_val(p.e_gpa, " GPa"),
                p.fit_tag
            ),
        );
        mech.push((name.clone(), p));
    }

    let desired = std::env::current_dir()?.join(format!("{}_result.xlsx", prefix));
    let writer = |target: &Path| -> Result<()> {
        let mut wb = Workbook::new();
        let ws_log = wb.add_worksheet();
        ws_log.set_name("运行日志")?;
        ws_log.write(0, 0, "程序版本")?;
        ws_log.write(0, 1, VERSION)?;
        ws_log.write(1, 0, "生成时间")?;
        ws_log.write(1, 1, Local::now().format("%Y-%m-%d %H:%M:%S").to_string())?;
        for (i, line) in logs_snapshot.iter().enumerate() {
            ws_log.write((i + 3) as u32, 0, line.as_str())?;
        }

        let ws_sum = wb.add_worksheet();
        ws_sum.set_name("力学性能汇总")?;
        for (i, h) in ["试样", "Rp0.2 (MPa)", "Rm (MPa)", "A (%)", "E (GPa)"].iter().enumerate() {
            ws_sum.write(0, i as u16, *h)?;
        }
        for (i, (name, p)) in mech.iter().enumerate() {
            let r = (i + 1) as u32;
            ws_sum.write(r, 0, name.as_str())?;
            ws_sum.write(r, 1, if p.rp02.is_nan() { "N/A".to_string() } else { format!("{:.3}", p.rp02) })?;
            ws_sum.write(r, 2, if p.rm.is_nan() { "N/A".to_string() } else { format!("{:.3}", p.rm) })?;
            ws_sum.write(r, 3, if p.a.is_nan() { "N/A".to_string() } else { format!("{:.3}", p.a) })?;
            ws_sum.write(r, 4, if p.e_gpa.is_nan() { "N/A".to_string() } else { format!("{:.3}", p.e_gpa) })?;
        }

        let cap = XLSX_MAX_ROWS - RESULT_HEADER_ROWS;
        let total_units: usize = results.iter().map(|(_, d)| d.stress.len().max(1)).sum::<usize>().max(1);
        let mut done = 0usize;
        let mut last = 0usize;
        progress_bar(0, 100, "保存进度 ");
        for (name, d) in results {
            let nrows = d.stress.len();
            let parts = ((nrows + cap - 1) / cap).max(1);
            for pi in 1..=parts {
                let ws = wb.add_worksheet();
                let sheet_name = if parts == 1 { name.clone() } else { format!("{}_p{}", name, pi) };
                let sname: String = sheet_name.chars().take(31).collect();
                ws.set_name(&sname)?;
                for (i, h) in ["载荷", "行程", "引伸计（修正）", "应力", "应变"].iter().enumerate() {
                    ws.write(0, i as u16, *h)?;
                }
                for (i, u) in ["kN", "mm", "mm", "MPa", "%"].iter().enumerate() {
                    ws.write(1, i as u16, *u)?;
                }
                let st = (pi - 1) * cap;
                let ed = (st + cap).min(nrows);
                for r in st..ed {
                    let rr = (r - st + 2) as u32;
                    for (c, v) in [d.load_kn[r], d.stroke_mm[r], d.ext_corr[r], d.stress[r], d.strain[r]].iter().enumerate() {
                        if v.is_nan() { ws.write(rr, c as u16, "")?; } else { ws.write(rr, c as u16, *v)?; }
                    }
                    done += 1;
                    let pct = done * 75 / total_units;
                    let stage = (pct / 25) * 25;
                    if stage > last {
                        last = stage;
                        progress_bar(stage, 100, "保存进度 ");
                    }
                }
            }
        }
        println!("[INFO] 结果数据写入完成，正在保存xlsx文件...");
        wb.save(target)?;
        progress_bar(100, 100, "保存进度 ");
        Ok(())
    };

    let out = write_safe(writer, &desired, state)?;
    log(state, &format!("\n  结果：{}", out.display()));
    Ok(out)
}

fn process_one(start: &Path, state: &mut AppState) -> Result<()> {
    std::env::set_current_dir(safe_parent_dir(start))?;
    let file_name = start.file_name().and_then(|s| s.to_str()).ok_or_else(|| anyhow!("无效路径"))?.to_string();
    let stem = Path::new(&file_name).file_stem().and_then(|s| s.to_str()).unwrap_or("out").to_string();
    let prefix = Regex::new(r"^(.+)-(\d+)$")?.captures(&stem).and_then(|c| c.get(1).map(|m| m.as_str().to_string())).unwrap_or(stem);

    log(state, "[INFO] 开始查找分段文件...");
    let files = find_csv_files(Path::new(&file_name), state)?;
    log(state, "[INFO] 开始读取并合并CSV...");
    let (rows, specimens) = load_and_merge(&files, state)?;
    log(state, &format!("  {} 个试样：{:?}", specimens.len(), specimens.iter().map(|s| s.name.clone()).collect::<Vec<_>>()));

    log(state, "[INFO] 预处理文件阶段...");
    let pre = std::env::current_dir()?.join(format!("pre-{}.xlsx", prefix));
    state.pre_artifact = Some(pre.clone());
    log(state, &format!("[INFO] 预处理文件路径: {}", pre.display()));

    log(state, "[INFO] 试样参数输入阶段...");
    let params = get_all_params(&specimens, state)?;

    log(state, "[INFO] 引伸计修正与力学计算阶段...");
    let mut results = Vec::new();
    for sp in &specimens {
        let raw = extract_specimen(&rows, sp.col_start);
        let p = params.get(&sp.name).ok_or_else(|| anyhow!("缺少参数: {}", sp.name))?;
        let (ext_corr, ext_mm) = process_extensometer(&raw, p.gauge, &sp.name)?;
        let d = calc_data(raw, ext_corr, ext_mm, p.area, p.gauge);
        results.push((sp.name.clone(), d));
    }

    log(state, "[INFO] 开始写出结果文件...");
    let result = write_result_xlsx(&results, &prefix, state)?;
    state.result_artifact = Some(result);
    Ok(())
}

fn main() {
    let mut state = AppState::default();
    log(&mut state, &format!("[INFO] 拉伸试验数据处理 {}", VERSION));
    loop {
        let input = prompt("\n[INFO] CSV文件路径（可直接拖入）: ");
        if input.is_empty() {
            log(&mut state, "[ERR] 请输入CSV文件路径");
            continue;
        }
        let start = PathBuf::from(input.trim_matches('"'));
        if !start.exists() {
            log(&mut state, "[ERR] 文件不存在，请检查路径");
            continue;
        }
        state.logs.clear();
        state.pre_artifact = None;
        state.result_artifact = None;
        log(&mut state, &format!("[INFO] 开始处理: {}", start.display()));
        if let Err(e) = process_one(&start, &mut state) {
            let dir = app_dir();
            if let Ok(p) = write_error_log(&dir, &state, &e) {
                log(&mut state, &format!("[ERR] 本次处理失败，详情见错误日志: {}", p.display()));
            }
            cleanup_artifacts(&mut state);
            cleanup_temp_files(&mut state);
            log(&mut state, "[INFO] 程序将在7秒后自动退出");
            thread::sleep(Duration::from_secs(7));
            break;
        }
        if !prompt("\n[INFO] 还要继续处理其他文件吗？[y/N]: ").eq_ignore_ascii_case("y") {
            break;
        }
    }
    log(&mut state, "[INFO] 全部任务已完成。");
}

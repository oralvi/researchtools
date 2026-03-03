use anyhow::Result;
use log::info;
use std::io::Write;

mod commands;

fn main() -> Result<()> {
    // 初始化日志
    env_logger::Builder::from_default_env()
        .filter_level("info".parse()?)
        .format(|buf, record| writeln!(buf, "[{}] {}", record.level(), record.args()))
        .init();

    info!("DCPD Console 启动");

    println!("\n╔════════════════════════════════════════╗");
    println!("║      DCPD 数据导出工具 (Rust 版)       ║");
    println!("╚════════════════════════════════════════╝\n");

    // 如果命令行提供参数，则批处理模式
    let mut args = std::env::args().skip(1);
    if let Some(path) = args.next() {
        let unit_str = if let Some(u) = args.next() { u } else { "sec".to_string() };
        let unit = unit_str.as_str();
        let src = std::path::Path::new(&path).canonicalize()?;
        match commands::process_file(&src, unit) {
            Ok((output_file, enc, point_count, duration)) => {
                println!();
                println!("[成功] 输入文件: {}", src.display());
                println!("[成功] 输入编码: {}", enc);
                println!("[成功] 输出文件: {}", output_file.display());
                println!("[成功] 数据点数: {}", point_count);
                println!(
                    "[成功] 实验时长: {:.3} {}",
                    duration,
                    if unit == "hr" { "小时" } else { "秒" }
                );
            }
            Err(e) => {
                println!("[错误] 处理失败: {}", e);
            }
        }
        return Ok(());
    }

    // 交互式循环
    loop {
        println!("请选择操作:");
        println!("1. 处理数据文件");
        println!("2. 退出");
        print!("\n请输入选择 (1 或 2): ");
        std::io::stdout().flush()?;

        let mut choice = String::new();
        std::io::stdin().read_line(&mut choice)?;

        match choice.trim() {
            "1" => {
                if let Err(e) = process_loop() {
                    println!("[错误] {}", e);
                }
            }
            "2" => {
                println!("[提示] 程序已退出。");
                break;
            }
            _ => {
                println!("[警告] 无效选择，请重试。\n");
            }
        }
    }

    Ok(())
}

fn process_loop() -> Result<()> {
    loop {
        println!();
        print!("请输入数据文件路径（例如 D:\\data\\test.txt）: ");
        std::io::stdout().flush()?;

        let mut file_path = String::new();
        std::io::stdin().read_line(&mut file_path)?;
        let file_path = file_path.trim();

        if file_path.is_empty() {
            println!("[警告] 路径不能为空，请重新输入。\n");
            continue;
        }

        let src = std::path::Path::new(file_path).canonicalize();
        let src = match src {
            Ok(p) => p,
            Err(_) => {
                println!("[警告] 文件不存在: {}\n", file_path);
                continue;
            }
        };

        if !src.is_file() {
            println!("[警告] 路径不是有效的文件: {}\n", src.display());
            continue;
        }

        println!();
        println!("请选择第一列的时间单位:");
        println!("1. sec - 秒");
        println!("2. hr  - 小时");
        print!("请输入选择 (1 或 2，默认为 1): ");
        std::io::stdout().flush()?;

        let mut unit_choice = String::new();
        std::io::stdin().read_line(&mut unit_choice)?;
        let unit = match unit_choice.trim() {
            "2" => "hr",
            _ => "sec",
        };

        // 处理文件
        match commands::process_file(&src, unit) {
            Ok((output_file, enc, point_count, duration)) => {
                println!();
                println!("[成功] 输入文件: {}", src.display());
                println!("[成功] 输入编码: {}", enc);
                println!("[成功] 输出文件: {}", output_file.display());
                println!("[成功] 数据点数: {}", point_count);
                println!(
                    "[成功] 实验时长: {:.3} {}",
                    duration,
                    if unit == "hr" { "小时" } else { "秒" }
                );
            }
            Err(e) => {
                println!("[错误] 处理失败: {}\n", e);
                continue;
            }
        }

        println!();
        print!("是否继续处理其他文件? (y/n，默认为n): ");
        std::io::stdout().flush()?;

        let mut again = String::new();
        std::io::stdin().read_line(&mut again)?;

        if again.trim().to_lowercase() != "y" {
            println!("[提示] 处理完成。\n");
            break;
        }
    }

    Ok(())
}

# DCPD Console - Rust CLI

纯 Rust 实现的 DCPD 命令行工具。轻快高效，无需依赖GUI框架。

## 项目结构

```
console/rust/
├── src/
│   ├── main.rs            # CLI入口，命令定义
│   ├── commands.rs        # 命令实现
│   └── models.rs          # 数据模型
├── Cargo.toml             # Rust依赖配置
└── README.md              # 本文件
```

## 快速开始

### 开发模式编译
```bash
cd DCPDexport/src/console/rust
cargo build
```

### 运行CLI
```bash
# 查看帮助
./target/debug/dcpd-console --help

# 分析数据
./target/debug/dcpd-console analyze data.csv

# 处理数据
./target/debug/dcpd-console process input.csv output.csv

# 检查状态
./target/debug/dcpd-console status

# 查看示例
./target/debug/dcpd-console examples
```

### 构建发布版本
```bash
cargo build --release
```

生成的可执行文件：
```
target/release/dcpd-console.exe  (Windows)
target/release/dcpd-console      (Linux/macOS)
```

## 命令说明

### analyze - 分析数据文件
```bash
dcpd-console analyze <FILE> [--format <FORMAT>]
```

参数：
- `FILE`: 输入文件路径
- `--format`: 输出格式（json|csv，默认json）

示例：
```bash
dcpd-console analyze data.csv
dcpd-console analyze data.csv --format json
```

### process - 处理数据文件
```bash
dcpd-console process <INPUT> <OUTPUT> [--options <OPTIONS>]
```

参数：
- `INPUT`: 输入文件路径
- `OUTPUT`: 输出文件路径
- `--options`: 处理选项（JSON格式）

示例：
```bash
dcpd-console process input.csv output.csv
dcpd-console process input.csv output.csv --options '{}'
```

### status - 检查系统状态
```bash
dcpd-console status
```

显示应用运行状态和系统信息。

### examples - 显示使用示例
```bash
dcpd-console examples
```

## 依赖

- `clap` - 命令行参数解析
- `tokio` - 异步运行时
- `serde` - 序列化/反序列化
- `log` - 日志记录
- `anyhow` - 错误处理

## 构建优化

`Cargo.toml` 中的 `[profile.release]` 已配置为最小化二进制文件大小：

```toml
[profile.release]
opt-level = "z"          # 最小化大小
lto = true               # 链接时优化
strip = true             # 移除调试符号
codegen-units = 1        # 单线程编译（更小的文件）
```

## 日志

设置日志级别：
```bash
dcpd-console --log-level debug analyze data.csv
dcpd-console --log-level info status
```

支持的日志级别：
- `error` - 只显示错误
- `warn` - 显示警告和错误
- `info` - 显示一般信息（默认）
- `debug` - 显示调试信息
- `trace` - 显示所有信息

## 开发指南

### 添加新命令

1. 在 `main.rs` 的 `Commands` 枚举中添加：
```rust
#[derive(Subcommand)]
enum Commands {
    MyCommand {
        /// 命令描述
        #[arg(value_name = "FILE")]
        file: PathBuf,
    },
}
```

2. 在 `main.rs` 的 `match` 中处理：
```rust
Commands::MyCommand { file } => {
    commands::my_command(&file).await?;
}
```

3. 在 `commands.rs` 中实现：
```rust
pub async fn my_command(file: &Path) -> Result<()> {
    // 你的实现
    Ok(())
}
```

### 修改数据模型

在 `models.rs` 中定义数据结构，使用 `#[derive(Serialize, Deserialize)]` 以支持 JSON 序列化。

## 性能

- ✅ 编译产物极小（< 10MB release）
- ✅ 启动速度快（无GUI框架开销）
- ✅ 内存占用低
- ✅ 适合服务器和自动化脚本

## 打包

### Windows 可执行文件
```bash
cargo build --release
# 输出: target/release/dcpd-console.exe
```

### 跨平台发布
使用 cargo-dist 工具：
```bash
cargo add cargo-dist
cargo dist
```

## 相关链接

- [Clap文档](https://docs.rs/clap/)
- [Tokio文档](https://tokio.rs/)
- [Rust Book](https://doc.rust-lang.org/book/)

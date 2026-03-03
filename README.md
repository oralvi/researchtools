# researchtools

本仓库包含两个主要工具包，均用于材料测试数据的处理与导出。

---

## DCPDexport

这是一个用于解析 DCPD（直流电压降）原始数据并输出 CSV 的工具集。

> **License:** MIT – 详见 [LICENSE](LICENSE)
> 
> **变更日志:** 请参阅 [CHANGELOG.md](CHANGELOG.md) 了解项目历次迭代。

### 目录结构

- `src/core.py` – 核心数据处理库，负责编码检测、数据解析和输出。
- `python/` – Python 工具脚本

### 功能

- 直接传参模式
- 交互式模式  
- 自动编码识别与多次尝试
- 覆盖提示和唯一输出文件名

---

## tensile-tool

用于处理拉伸测试数据。结构为：

- `python/`：几个演化版本的处理脚本 (`process_tensile_data.py` … `v5`)。
- `rust/`：Rust 实现。

---

## 忽略规则

根目录 `.gitignore` 已配置，忽略：Python 缓存/虚拟环境、构建/打包输出、测试/样本文件、Rust `target/`、IDE/OS 临时文件等，适合直接 Git 管理。

---

### 使用提示

- 为了运行脚本，请安装依赖：
  ```bash
  pip install chardet
  ```
- 命令行示例：
  ```bash
  python src/core.py data.txt
  ```

以上说明提供了项目的整体结构与各个工具的功能说明。
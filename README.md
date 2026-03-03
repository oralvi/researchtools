# researchtools

本仓库包含两个主要工具包，均用于材料测试数据的处理与导出。

---

## DCPDexport

这是一个用于解析 DCPD（直流电压降）原始数据并输出 CSV 的工具集。目录结构及各版本概览：

> **License:** MIT – 详见 [LICENSE](LICENSE)
> 
> **变更日志:** 请参阅 [CHANGELOG.md](CHANGELOG.md) 了解项目历次迭代。



- `src/` – 核心代码及两种前端。
  - `core.py` – 统一的处理库，负责编码检测、数据解析和输出。
  - `console/python/dcpd_console.py` – 命令行工具，支持：
    - 直接传参模式 `python -m DCPDexport.src.console.python.dcpd_console input.txt`
    - 交互式模式 `--interactive`（询问文件路径、时间单位）
    - 自动编码识别与多次尝试、覆盖提示和唯一输出文件名。
  - `gui/python/dcpd_gui.py` – 简洁 PySide6 窗口版。功能等价于 console，但添加文件选择、进度条和日志面板。GUI 默认调用共享 `core`。

- `250109gui104.py` – 遗留 GUI 实现，带有丰富样式和背景设置；保留仅作历史参考。
- `260109new.py` – 早期交互式控制台脚本，使用 `input()` 循环；已被整合进 `console`、`core`。

版本演进说明：

当前项目结构反映了以下演进过程：
1. **初期版本**：存在多个单文件实现（如 `250109gui104.py` 的富样式 GUI、`260109new.py` 的交互式控制台等）
2. **模块化重构**：将数据处理逻辑提炼为 `src/core.py`、控制台前端为 `src/console/python/dcpd_console.py`、GUI 为 `src/gui/python/dcpd_gui.py`
3. **当前状态**：统一的包结构，支持
   - **命令行调用**：通过 `-m` 或直接脚本执行
   - **交互式模式**：`--interactive` 选项  
   - **图形界面**：简洁的 PySide6 窗口版本
   - **共享基础**：编码检测、数据解析、文件输出逻辑统一在 core 中，避免重复

---

## tensile-tool

用于处理拉伸测试数据。结构为：

- `python/`：几个演化版本的处理脚本 (`process_tensile_data.py` … `v5`)。
- `rust/`：两个 Rust 实现目录，分别标记为 `v15` 和 `legacy`。

---

## 忽略规则

根目录 `.gitignore` 已配置，忽略：Python 缓存/虚拟环境、构建/打包输出、测试/样本文件、Rust `target/`、IDE/OS 临时文件等，适合直接 Git 管理。

---

### 使用提示

- 为了运行 GUI 或控制台，请安装依赖：
  ```bash
  pip install chardet PySide6
  ```
- 命令行示例：
  ```bash
  python -m DCPDexport.src.console.python.dcpd_console data.txt --unit hr
  ```
- GUI 启动：
  ```bash
  python -m DCPDexport.src.gui.python.dcpd_gui
  ```

以上说明提供了项目的整体结构与各个版本的功能历史，便于后续迭代。
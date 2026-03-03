# Change Log

项目迭代历史记录，按时间顺序排列主要变更。

## 2026-03-03  完成最终版本
- Console 全面重写为交互式循环，添加实验时长输出；新增命令行批处理 `console <path> [unit]` 支持。
- GUI 前端改造为拖放界面并实现真实文件处理逻辑。
- 后端共用解析/编码检测代码，支持多编码自动识别。
- 输出自动命名 CSV；处理过程支持秒/小时单位。
- 生成的 exe 文件分别拷贝至 `dist/` 目录。
- 发布 MSI/NSIS 安装包。
- 添加许可 MIT，创建 LICENSE 文件。
- 清理根目录多余脚本，新增 CHANGELOG 和合并文档。

## 2026-02-xx  GUI/Console 初次重构
- 从 Python 版本迁移至 Rust + Tauri。
- 实现带有命令行子命令的 Console 和基本 GUI。
- 初步尝试图标嵌入与 tauri.conf.json 修复。

## 2026-01-xx  项目结构规划
- 设计 `DCPDexport` 包含 `console` 与 `gui` 两种前端。
- 编写 Python 原型和测试样本。

> 此日志从最初阶段开始，早期更改详情见各个版本文档。

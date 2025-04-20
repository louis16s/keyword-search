# 📊 Excel 关键词搜索工具（Terminal GUI 版）

一个支持 `.xls` / `.xlsx` 的关键词批量搜索工具，适用于大批量 Excel 文件搜索，支持终端界面操作、多进程处理、结果导出为 Excel，并具备自动检查更新功能。

## 🚀 功能特性

- ✅ 批量搜索 Excel 文件中的关键词（支持 .xls / .xlsx）
- ✅ 使用 `config.ini` 配置搜索路径与关键词表
- ✅ 多进程并发处理，大幅提高搜索效率
- ✅ 输出结果为 `搜索结果.xlsx`
- ✅ 终端界面美观，支持菜单选择操作
- ✅ 自动检查 GitHub 最新版本更新

## 🛠 使用方法

### 1. 安装依赖

```bash
pip install openpyxl xlrd tqdm rich requests

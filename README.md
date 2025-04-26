# Excel 关键词搜索工具 v2.2

一个终端 GUI 风格的 Excel 批量关键词搜索工具，支持 `.xls` 和 `.xlsx` 文件，支持多进程并行、关键词配置、自动检查更新。

## 🌟 功能特性

- 支持 `.xlsx` 与 `.xls` 批量搜索
- 支持多进程加速搜索
- 支持配置文件修改（关键词Excel、搜索路径）
- 支持检查 GitHub 最新版本
- 美观的终端界面（rich 实现）
- 支持自动下载更新
## 📦 使用方法

1. 安装依赖：
    ```bash
    pip install openpyxl xlrd rich tqdm requests packaging
    ```

2. 运行程序：
    ```bash
    python main.py
    ```

3. 根据提示完成配置与搜索。

## 📁 配置文件示例 `config.ini`

```ini
[Settings]
search_directory = ./data
excel_file_path = ./keywords.xlsx
github_token = token

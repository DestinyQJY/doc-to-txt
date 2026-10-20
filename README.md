# doc-to-txt

## 📖 项目简介

将 **`.doc`、`.docx`** 文件转换为 **UTF-8 编码** 的 **`.txt`** 文件，支持使用**拖拽**批量处理。

> **注意**：仅支持 **Windows**，需已安装 **Microsoft Word 桌面版**。

## 📝 使用方法

### 使用 EXE（推荐）

1. 前往 [**GitHub Releases**](https://github.com/DestinyQJY/doc-to-txt/releases/latest)，下载最新版 **`doc_to_txt.exe`**。

2. 选中**一个或多个** **`.doc`**、**`.docx`** 文件，将其**拖拽**到 **`doc_to_txt.exe`** 上。

3. 程序会在源文件所在目录依次生成同名 **`.txt`** 文件。

> ⚠️ **注意**：若同名 **`.txt`** 文件已存在，将被**覆盖**，请提前备份。

### 使用命令行

在 **`doc_to_txt.exe`** 所在目录打开终端，然后运行：

```powershell
.\doc_to_txt.exe "D:\Documents\example.docx"
```

> ⚠️ **注意**：必须传入**文件路径**，路径包含空格时需用引号包裹，不支持传入文件夹。

### 使用源码

- 下载 **`doc_to_txt.py`**，将其放入项目目录，并在该目录中创建虚拟环境、安装依赖：

    ```powershell
    py -m venv .venv
    .\.venv\Scripts\Activate.ps1
    pip install pywin32
    ```

- 运行：

    ```powershell
    py doc_to_txt.py "D:\Documents\example.docx"
    ```

- 重新打包：

    ```powershell
    pip install pyinstaller
    pyinstaller --onefile doc_to_txt.py
    ```

## 📌 注意事项

- 若转换失败，建议先用 Word 手动打开并检查 **`.doc`、`.docx`** 源文件是否可正常读取。

- TXT 仅保留纯文本，无法保留图片、字体、排版等 Word 格式；表格等复杂内容以 Word 的纯文本导出结果为准。

## 💬 维护与反馈

本项目已满足作者的日常使用需求，仓库主要用于备份与分享，不承诺固定更新频率。

如果在使用过程中遇到问题，或有功能改进建议，欢迎通过 [GitHub Issues](https://github.com/DestinyQJY/doc-to-txt/issues) 反馈。
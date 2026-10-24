# 新手教程（Word 版）

按顺序读：

1. `01-新手配置教程.docx` —— 从零装环境、改第一页、用编辑器和 AI。
2. `02-单页制作流程.docx` —— 领页到上线的十步。
3. `03-页面清单.docx` —— 谁负责哪一页。
4. `04-项目约定.docx` —— 每次让 AI 写页面前先贴的规则。
5. `05-变更记录.docx` —— 共享的变更日志。
6. `06-开发方案与操作手册.docx` —— 技术方案与完整手册。

## 说明

- Word 是发给同事阅读的版本，**内容以 `docs/*.md` 为准**。
- 改了 Markdown 之后，用内置 Python 重新导出：

```powershell
python tools\build_docx.py
```

如果系统 Python 没有 `python-docx`，用 DSH 自带的那份：

```powershell
& "C:\Users\cjx\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe" tools\build_docx.py
```

- `07-交接日志` 是本地内部文档，不在这个文件夹里。

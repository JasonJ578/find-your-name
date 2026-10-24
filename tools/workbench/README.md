# 本地可视化编辑器

这是从 `@vibe-x/dsh-html-workbench` v0.3.1 中取出的独立运行版本（内置 html-workbench 2.2.0，MIT 许可）。

- `scripts/workbench.py`：本地服务，只用 Python 标准库，不需要 pip 安装任何东西。
- `assets/workbench.html`：编辑器界面。
- 首次启动会下载 GrapesJS 到 `%LOCALAPPDATA%\html-workbench`，之后离线可用。

## 用法

把要编辑的 HTML 文件拖到项目根目录的 `editor.cmd` 上，松手。

或者手动：

```powershell
python tools/workbench/scripts/workbench.py open "weibo/index.html" --editor-root . --port 4317
```

浏览器打开 `http://127.0.0.1:4317/`。

## 注意

- 它只认页面内的 `<style>`，不加载外链 CSS。所以页面要先跑 `tools/inline-css.py`。
- 不要改页面里的 `<style>` 块；改 `assets/css/` 后重新同步。
- 服务只绑定 `127.0.0.1`，不会暴露到局域网。

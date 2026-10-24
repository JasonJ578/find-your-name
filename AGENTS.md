# 给 AI 的项目规则

1. **新对话第一件事**：读 `docs/交接-本地.md` 的**最后一篇**（如果存在）。
2. 写页面前，先读 `docs/项目约定.md`。
3. 新页面以 `templates/page.html` 为骨架。
4. 不要手改页面里的 `<style>` 块；改 `assets/css/` 后运行 `python tools/inline-css.py --force`。
5. **收工前**：
   - 在 `docs/交接-本地.md` 末尾追加一篇交接；
   - 在 `docs/变更记录.md` 末尾追加一条，写清修改人和时间。
6. 永不删除交接历史。

"""把 assets/css 里的样式同步进每个页面。

用法：
    python tools/inline-css.py            # 只处理还带 <link> 的页面
    python tools/inline-css.py --force    # 重写所有页面的 <style> 块（改主题后用）

为什么需要它：
    HTML Workbench 编辑器不加载外链 CSS，只认页面内的 <style>。
    所以以 assets/css 为样式源，脚本负责把样式写进每个页面。

规则：
    - 页面里的 <style> 块由本脚本管理，不要手改。
    - 改全站颜色、字号、间距：改 assets/css，然后 --force。
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

THEME = (ROOT / "assets" / "css" / "theme.css").read_text(encoding="utf-8")
LAYOUT = (ROOT / "assets" / "css" / "layout.css").read_text(encoding="utf-8")

STYLE_BLOCK = (
    "<!-- INLINE-CSS: 由 assets/css 生成，改样式请改源文件后运行 tools/inline-css.py --force -->\n"
    "<style>\n"
    + THEME.strip()
    + "\n\n"
    + LAYOUT.strip()
    + "\n</style>"
)

LINK_PATTERN = re.compile(
    r'[ \t]*<link[^>]*href="[^"]*assets/css/(?:theme|layout)\.css"[^>]*>\s*\n?'
)
INLINE_PATTERN = re.compile(r"<!-- INLINE-CSS.*?-->.*?</style>", re.S)

SKIP_PARTS = {".git", "docs", "tools", "templates", "node_modules"}


def iter_pages():
    for path in ROOT.rglob("*.html"):
        if any(part in SKIP_PARTS for part in path.parts):
            continue
        yield path


def process(path: Path, force: bool) -> str:
    html = path.read_text(encoding="utf-8")
    has_inline = INLINE_PATTERN.search(html) is not None
    has_link = LINK_PATTERN.search(html) is not None

    if has_inline and not force:
        return "skip"

    if has_inline:
        html = INLINE_PATTERN.sub(STYLE_BLOCK, html, count=1)
    elif has_link:
        html = LINK_PATTERN.sub("", html)
        if "</head>" not in html:
            return "error: no </head>"
        html = html.replace("</head>", STYLE_BLOCK + "\n</head>", 1)
    else:
        return "skip"

    path.write_text(html, encoding="utf-8")
    return "done"


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--force",
        action="store_true",
        help="重写所有页面的 <style> 块；改过主题后使用",
    )
    args = parser.parse_args()

    counts = {"done": 0, "skip": 0, "error": 0}
    for page in iter_pages():
        result = process(page, args.force)
        key = "error" if result.startswith("error") else result
        counts[key] += 1
        print(f"{result:>12}  {page.relative_to(ROOT)}")

    print(f"\n完成：写入 {counts['done']} 页，跳过 {counts['skip']} 页，失败 {counts['error']} 页。")


if __name__ == "__main__":
    main()

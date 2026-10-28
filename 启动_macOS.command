#!/bin/bash
# macOS Finder 双击；使用终端输入，避免对 tkinter 的依赖。
cd "$(dirname "$0")" || exit 1
printf '\n论文合集 Word 格式审校工具（macOS）\n\n'
if ! command -v python3 >/dev/null 2>&1; then
  echo '未发现 python3，请先安装 Python 3.9 或更新版本。'
  read -r -p '按回车退出... ' _
  exit 1
fi
python3 - <<'PY'
import sys, shlex
from pathlib import Path
try:
    from word_audit import audit_dir
except Exception as exc:
    print('加载失败：', exc)
    input('按回车退出... ')
    raise SystemExit(1)
if sys.version_info < (3, 9):
    print('需要 Python 3.9 或更新版本，当前版本：', sys.version.split()[0])
    input('按回车退出... ')
    raise SystemExit(1)
print('将存放 .docx 的文件夹拖到这个终端窗口，然后按回车。')
raw=input('稿件文件夹：').strip()
if not raw:
    print('未输入文件夹。')
    input('按回车退出... ')
    raise SystemExit(1)
try:
    tokens=shlex.split(raw)
    source=Path(tokens[0] if len(tokens)==1 else raw).expanduser()
    out, items=audit_dir(source)
    print(f'\n已检查 {len(items)} 份稿件。')
    print('报告目录：',out)
    print('请打开「审校报告.html」查看；亦可用 Excel/Numbers 打开两个 CSV 文件。')
except Exception as exc:
    print('运行失败：',exc)
input('\n按回车退出... ')
PY
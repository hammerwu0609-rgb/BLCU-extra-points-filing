# 离开 WorkBuddy 还能用吗？

**结论：4 个脚本是纯 Python 命令行工具，不依赖任何 AI 平台，拷走就能用；
SKILL.md 离开 WorkBuddy 后退化为一份普通 Markdown 说明书（人可读、任何会读 Markdown 的 agent 可读），
自动触发、AskUserQuestion、present_files 等专有机制不再可用。**

## 平台矩阵（哪些部分在哪能用）

| 能力 | Windows + Excel | Windows 无 Excel | macOS / Linux |
|---|---|---|---|
| `scan_sources.py` 全量检索 | ✅ | ✅ | ✅ |
| `fill_template.py` 填表/公式/一页打印 | ✅ | ✅（加 `--skip-recalc`） | ✅（加 `--skip-recalc`） |
| `fill_template.py` Excel 重算回写缓存 | ✅ | ❌ | ❌ |
| `final_check.py` 数值/截图自检 | ✅ | ✅（`"verify_pages": false`） | ✅（`"verify_pages": false`） |
| `final_check.py` 打印页数检查 | ✅ | ❌ | ❌ |
| `make_range_shots.py` 来源截图 | ✅ | ❌ | ❌ |
| `cleanup_excel.py` 清残留 Excel 进程 | ✅ | ✅（无进程可清） | 不适用 |

依赖项分两层：

```bash
# 所有平台都要（检索 + 填表 + 自检）
pip install -r requirements.txt        # openpyxl xlrd pillow pymupdf

# 仅 Windows + 装有 Microsoft Excel 时才需要（截图 / 重算 / 页数检查）
pip install pywin32
```
> 说明：`python-docx` **不在**依赖里 —— `scan_sources.py` 读 docx 用的是标准库
> `zipfile + xml.etree` 直接解 `word/document.xml`，不需要第三方库。

## 各脚本独立用法（任何电脑的命令行都能跑）

```bash
# 1) 全量检索：强/弱/班级三档标签，输出 A/B/C/D 四区
python scripts/scan_sources.py \
    --root "D:/附加分来源" \
    --name "张三" --sid "202400000001" \
    --clazz "24xx" --out hits.txt

# 2) 填表（跨平台；Windows+Excel 下再去掉 --skip-recalc 可回写公式缓存值）
python scripts/fill_template.py --config form_config.json --skip-recalc

# 3) 逐条来源截图（仅 Windows + Excel）
python scripts/make_range_shots.py --config shots_config.json

# 4) 交付自检（跨平台；页数检查可关）
python scripts/final_check.py --config check_config.json

# 5) 收工清理：只结束「无可见窗口」的 Excel 残留，不误伤你自己开着的 Excel
python scripts/cleanup_excel.py --dry-run   # 先看会杀谁
python scripts/cleanup_excel.py
```
配置写法看 `examples/*.example.json`，字段含义见各脚本头部 docstring。

## SKILL.md 里哪些内容是 WorkBuddy 专属的

| WorkBuddy 专属 | 通用 |
|---|---|
| frontmatter 自动触发 / `agent_created` | Markdown 正文流程骨架（§1–§9） |
| `AskUserQuestion` 交互提问 | 「哪些点必须停下来问用户」清单本身 |
| `present_files` 交付展示 | §5.1 的交付目录结构约定 |
| PowerShell 工具不回显 stdout、需写日志再 Read | 通用 Windows 命令行知识 |
| 沙箱拦截 `shutil.rmtree`/`os.remove`（safe-delete） | Python 本身的删除写法 |
| `~/.workbuddy/binaries/python/...` 托管解释器路径 | 换成你自己的 `python` 即可 |

## macOS / Linux 上做「来源截图」的替代思路

`make_range_shots.py` 依赖 Excel COM（pywin32），这两套系统没有直接等价物。可选：
1. **LibreOffice headless**：先用 openpyxl 造一张只含「表头 + 目标行」的临时表，再
   `soffice --headless --convert-to pdf --outdir out/ temp.xlsx`，最后 PyMuPDF 转 PNG。
   （不能直接转原表，否则会导出全部几千行。）
2. **macOS + Microsoft Excel**：用 JXA / AppleScript 完成同样的「删行 → 标注 → 导 PDF」。
3. **最省事**：把截图这一步留在 Windows 机器上跑，其它步骤跨平台。

## 三条与平台无关的硬提醒（换环境也成立）

1. Excel 不可见时 `Range.CopyPicture` 导出的是空白图 → 必须走「删行 → 导 PDF → PyMuPDF 转 PNG」。
2. pywin32 只 `xl.Quit()` 不会结束 Excel 进程，**每次跑完都要查残留**：
   Windows `Get-Process EXCEL`，`MainWindowTitle` 为空的才是自动化残留。
3. 中间产物写 `tempfile.gettempdir()`，别写进交付包，否则会多出 `.workbuddy/`、`~$xxx.xlsx` 杂质。

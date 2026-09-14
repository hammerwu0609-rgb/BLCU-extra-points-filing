# -*- coding: utf-8 -*-
"""多标签全量检索某位学生在本地「附加分来源表」中的全部记录。

用法:
  python scan_sources.py --root "D:/附加分来源" --name 张三 --sid 202400000001 \
      [--clazz 24xx] [--extra 张三,李四] [--out hits.txt]

核心设计 —— 标签分三档，避免短片段造成假阳性:
  强标签（命中即认定是本人）: 姓名全名、学号全号、--extra 里给的名字
  弱标签（命中只算「疑似」，需人工判断）: 姓名片段、学号后4位/后3位
  班级标签（命中只进「候选集体条目」）: --clazz

输出四块内容:
  A. 个人命中清单（强标签）
  B. 疑似命中（仅弱标签，多为序号列/他人学号巧合，须逐条人工排除）
  C. 候选集体条目（仅班级命中，如「24xx 五四评优 优秀团支部 0.4」）
  D. XML 兜底校验（sharedStrings 计数 vs 逐行命中数，两者不一致说明有漏检）

支持: .xlsx（openpyxl）、.xls（xlrd）、.docx（说明/通知类文件文本扫描）
"""
import argparse, glob, os, re, zipfile
from xml.etree import ElementTree as ET

try:
    from openpyxl import load_workbook
except ImportError:
    print('缺少 openpyxl，请先: pip install openpyxl'); raise SystemExit(1)
try:
    import xlrd
except ImportError:
    xlrd = None

W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def dedup(seq):
    seen, out = set(), []
    for t in seq:
        if t and t not in seen:
            seen.add(t); out.append(t)
    return out


def build_tags(name, sid, clazz, extra):
    strong, weak = [], []
    if name:
        n = str(name).strip()
        strong.append(n)
        if len(n) >= 3:
            weak += [n[1:], n[:-1]]                 # 去姓 / 去名
    if sid:
        s = str(sid).strip()
        strong.append(s)
        if len(s) >= 6:
            weak += [s[-4:], s[-3:]]                # 学号后4位 / 后3位
    for x in extra or []:
        x = x.strip()
        if x:
            strong.append(x)                        # 补充的名字按强标签处理
    return dedup(strong), dedup(weak), (clazz or '').strip()


def row_text_xlsx(row, max_col=40):
    return ' | '.join('%s=%s' % (c.coordinate, c.value) for c in row[:max_col] if c.value is not None)


def scan_xlsx(path, tags):
    hits = []
    wb = load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(max_col=40):
            txt = row_text_xlsx(row)
            hit = [t for t in tags if t in txt]
            if hit:
                hits.append((ws.title, row[0].row, hit, txt))
    return hits


def scan_xls(path, tags):
    if xlrd is None:
        return [('SKIP', 0, ['xlrd 未安装'], '无法读取 .xls: %s' % os.path.basename(path))]
    hits = []
    wb = xlrd.open_workbook(path)
    for sh in wb.sheets():
        for r in range(sh.nrows):
            cells = ['%s%d=%s' % (xlrd.formula.colname(c), r + 1, sh.cell_value(r, c))
                     for c in range(sh.ncols) if str(sh.cell_value(r, c)).strip()]
            txt = ' | '.join(cells)
            hit = [t for t in tags if t in txt]
            if hit:
                hits.append((sh.name, r + 1, hit, txt))
    return hits


def scan_docx(path, tags):
    hits = []
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    text = '\n'.join((t.text or '') for t in root.iter(W_NS + 't'))
    for t in tags:
        for m in re.finditer(re.escape(t), text):
            s = max(0, m.start() - 40)
            snippet = text[s:m.end() + 60].replace('\n', ' / ')
            lineno = text[:m.start()].count('\n') + 1
            hits.append(('正文', lineno, [t], snippet))
    return hits


def xml_probe(path, tags):
    """直接在 xlsx 内部 XML 里数标签出现次数，用于发现逐行扫描的漏检。"""
    found = {}
    try:
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n.endswith('.xml') and ('sharedStrings' in n or 'sheet' in n.lower()):
                    data = z.read(n).decode('utf8', 'ignore')
                    for t in tags:
                        c = data.count(t)
                        if c:
                            found[t] = found.get(t, 0) + c
    except Exception as e:
        found['<读取失败>'] = str(e)
    return found


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--root', required=True, help='来源文件夹')
    ap.add_argument('--name', default='', help='姓名')
    ap.add_argument('--sid', default='', help='学号')
    ap.add_argument('--clazz', default='', help='班级（用于候选集体条目）')
    ap.add_argument('--extra', default='', help='补充强标签，逗号分隔（别名、曾用名等）')
    ap.add_argument('--out', default='', help='结果另存路径')
    a = ap.parse_args()

    if not (a.name or a.sid):
        raise SystemExit('必须至少提供 --name 或 --sid（身份是硬性前提）')

    strong, weak, clazz = build_tags(a.name, a.sid, a.clazz, [x for x in a.extra.split(',') if x.strip()])
    all_tags = dedup(strong + weak + ([clazz] if clazz else []))

    files = sorted(glob.glob(os.path.join(a.root, '**', '*.xlsx'), recursive=True)) + \
            sorted(glob.glob(os.path.join(a.root, '**', '*.xls'), recursive=True)) + \
            sorted(glob.glob(os.path.join(a.root, '**', '*.docx'), recursive=True))

    personal, suspect, classonly, probes = [], [], [], []
    for f in files:
        if '~$' in os.path.basename(f):
            continue
        ext = f.lower().rsplit('.', 1)[-1]
        try:
            hits = scan_xlsx(f, all_tags) if ext == 'xlsx' else \
                   scan_xls(f, all_tags) if ext == 'xls' else \
                   scan_docx(f, all_tags)
        except Exception as e:
            personal.append((os.path.relpath(f, a.root), 'ERROR', 0, [str(e)], ''))
            continue
        for (sheet, rno, hit, txt) in hits:
            rec = (os.path.relpath(f, a.root), sheet, rno, hit, txt)
            if any(t in strong for t in hit):
                personal.append(rec)
            elif clazz and clazz in hit:
                classonly.append(rec)
            else:
                suspect.append(rec)
        if ext == 'xlsx':
            p = xml_probe(f, all_tags)
            if p:
                probes.append((os.path.relpath(f, a.root), p))

    lines = ['=' * 100,
             '强标签(命中即本人): %s' % ' / '.join(strong),
             '弱标签(命中仅疑似): %s' % ' / '.join(weak),
             '班级标签(仅候选集体): %s' % clazz,
             '=' * 100]

    lines.append('\n【A. 个人命中清单】共 %d 条' % len(personal))
    for (f, sheet, rno, hit, txt) in personal:
        lines.append('  %s :: %s :: 第%s行 :: 命中[%s]\n      %s' % (f, sheet, rno, ','.join(hit), txt[:300]))
    if not personal:
        lines.append('  （无命中 —— 检查姓名/学号是否正确，或用 --extra 增加别名）')

    lines.append('\n【B. 疑似命中（仅弱标签，多为序号列或他人学号巧合，须逐条排除）】共 %d 条' % len(suspect))
    for (f, sheet, rno, hit, txt) in suspect[:120]:
        lines.append('  %s :: %s :: 第%s行 :: 命中[%s]\n      %s' % (f, sheet, rno, ','.join(hit), txt[:260]))
    if len(suspect) > 120:
        lines.append('  ...（其余 %d 条略）' % (len(suspect) - 120))

    lines.append('\n【C. 候选集体条目（仅班级命中，需用户判断是否归属本人）】共 %d 条' % len(classonly))
    for (f, sheet, rno, hit, txt) in classonly[:200]:
        lines.append('  %s :: %s :: 第%s行\n      %s' % (f, sheet, rno, txt[:260]))
    if len(classonly) > 200:
        lines.append('  ...（其余 %d 条略）' % (len(classonly) - 200))

    lines.append('\n【D. XML 兜底校验】（内部出现次数；为 0 但 A 区也无命中 → 该文件确实没有）')
    for f, p in probes:
        lines.append('  %s -> %s' % (f, p))
    if not probes:
        lines.append('  （xlsx 内部未发现任何标签）')

    lines.append('\n【结论提示】')
    lines.append('  · A 区 = 直接可申报的记录；')
    lines.append('  · C 区 = 「全班都有」的集体荣誉，按学校规则判断是否折半；')
    lines.append('  · B 区 = 必须逐条人工排除，不要默认忽略；')
    lines.append('  · 把 A+C 的行号抄进 shots_config.json 做来源截图。')

    out = '\n'.join(lines)
    print(out)
    if a.out:
        with open(a.out, 'w', encoding='utf-8') as fh:
            fh.write(out + '\n')
        print('\n[已保存] %s' % a.out)


if __name__ == '__main__':
    main()

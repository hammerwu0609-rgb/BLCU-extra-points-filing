# -*- coding: utf-8 -*-
"""把来源表中的「指定几行」裁剪成 PNG 截图：保留原表样式，橙底红框标注目标行，底部注明出处。

用法:
  python make_range_shots.py --config shots_config.json

shots_config.json 示例:
{
  "base_dir": "C:/Users/xx/Desktop/xx综测附加分申报",
  "source_subdir": "附加分来源",
  "out_dir": "附加分来源截图",
  "dpi": 200,
  "col_min_width": 9,
  "col_max_width": 46,
  "tasks": [
    {
      "src": "【xx学院】2025-2026学年校级及以上级别综测附加分信息统计表.xlsx",
      "sheet": "荣誉表彰",
      "last": 1261,                    // 原表总行数（可用 openpyxl ws.max_row 预先查好）
      "keep": [1, 2, 3, 86],           // 要保留的行：表头行 + 目标行
      "hl":   [86],                    // 要标注的目标行（原表行号）
      "hdr":  3,                       // 列头所在行（其上方的说明行会做换行处理）
      "note": "对应申报表「各级荣誉」校级优秀团支部",
      "name": "03_各级荣誉_五四评优校级优秀团支部_校级表-荣誉表彰.png"
    }
  ]
}

原理与坑:
- Excel 不可见时 Range.CopyPicture 导出的是空白图 → 必须走 ExportAsFixedFormat 导 PDF，再用 PyMuPDF 转 PNG。
- 删行必须「从下往上」，否则行号漂移；标注行新行号 = 原行号 - 上方被删行数。
- 源表标题行常自带黄色底纹，标注改用「浅橙底纹 + 红框」才能区分。
- 合并说明行必须 WrapText + 抬高行高，否则长文本会被左右同时裁掉。
- 只用 xl.Quit() 不会结束 Excel 进程，finally 里要 del 子对象 + gc.collect() + CoUninitialize()。
"""
import argparse, json, math, os, shutil, tempfile, time
from PIL import Image, ImageChops

try:
    import win32com.client as win32
except ImportError:
    print('缺少 pywin32，请先: pip install pywin32'); raise SystemExit(1)
try:
    import pymupdf
except ImportError:
    print('缺少 pymupdf，请先: pip install pymupdf'); raise SystemExit(1)

RGB = lambda rgb: rgb[2] * 65536 + rgb[1] * 256 + rgb[0]     # Excel Color 为 BGR


def bgr(rgb):
    return rgb[2] * 65536 + rgb[1] * 256 + rgb[0]


def crop_white(im, pad=10):
    bbox = ImageChops.difference(im, Image.new('RGB', im.size, (255, 255, 255))).getbbox()
    if bbox:
        im = im.crop((max(0, bbox[0] - pad), max(0, bbox[1] - pad),
                      min(im.width, bbox[2] + pad), min(im.height, bbox[3] + pad)))
    return im


def run(cfg):
    base = cfg['base_dir']
    src_dir = os.path.join(base, cfg.get('source_subdir', ''))
    out_dir = os.path.join(base, cfg.get('out_dir', '附加分来源截图'))
    # 临时目录默认用系统临时目录；要留在项目里就配 "tmp_dir"。
    # 千万别默认写进交付包，否则包里会多出 .workbuddy/ 这类杂质目录。
    tmp = cfg.get('tmp_dir') or os.path.join(tempfile.gettempdir(),
                                             'extra_points_shots_%d' % int(time.time()))
    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(tmp, exist_ok=True)

    dpi = cfg.get('dpi', 200)
    cmin = cfg.get('col_min_width', 9)
    cmax = cfg.get('col_max_width', 46)
    fill_c = bgr(cfg.get('highlight_rgb', [255, 214, 153]))
    line_c = bgr(cfg.get('border_rgb', [255, 0, 0]))
    max_col = cfg.get('max_col', 10)

    xl = win32.DispatchEx('Excel.Application')
    xl.DisplayAlerts = False
    xl.Visible = False
    xl.ScreenUpdating = False

    lines = []
    try:
        for t in cfg['tasks']:
            work = os.path.join(tmp, t['src'])
            shutil.copyfile(os.path.join(src_dir, t['src']), work)
            wb = xl.Workbooks.Open(work)
            for i in range(wb.Worksheets.Count, 0, -1):
                if wb.Worksheets.Item(i).Name != t['sheet']:
                    wb.Worksheets.Item(i).Delete()
            ws = wb.Worksheets.Item(1)
            ws.Activate()

            # 1) 需要删除的行区间（升序）
            ranges, prev = [], 0
            for k in sorted(t['keep']):
                if k - 1 >= prev + 1:
                    ranges.append((prev + 1, k - 1))
                prev = k
            if t['last'] > prev:
                ranges.append((prev + 1, t['last']))

            # 2) 标注行删除后的新行号
            new_rows = []
            for h in t['hl']:
                shift = sum(r1 - r0 + 1 for (r0, r1) in ranges if r1 < h)
                new_rows.append(h - shift)

            # 3) 从下往上删
            for (r0, r1) in sorted(ranges, key=lambda x: -x[0]):
                ws.Rows('%d:%d' % (r0, r1)).Delete()

            # 4) 标注：橙底 + 红框 + 加粗
            for nr in new_rows:
                rg = ws.Range('A%d:%s%d' % (nr, chr(64 + max_col), nr))
                rg.Interior.Color = fill_c
                rg.Font.Bold = True
                for edge in (7, 8, 9, 10):          # 左/上/下/右
                    rg.Borders(edge).LineStyle = 1
                    rg.Borders(edge).Weight = 3
                    rg.Borders(edge).Color = line_c

            # 5) 列宽 + 居中
            last_row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            try:
                ws.Range('A1', ws.Cells(last_row, max_col)).EntireColumn.AutoFit()
            except Exception:
                pass
            for c in range(1, max_col + 1):
                col = ws.Columns(c)
                if col.ColumnWidth > cmax:
                    col.ColumnWidth = cmax
                if col.ColumnWidth < cmin:
                    col.ColumnWidth = cmin
                col.HorizontalAlignment = -4108     # 居中

            # 6) 说明行自动换行，防截断
            ws.Rows(1).RowHeight = max(ws.Rows(1).RowHeight, 27)
            for r in range(1, t.get('hdr', 3)):
                txt = str(ws.Cells(r, 1).Value or '')
                try:
                    ws.Range('A%d:%s%d' % (r, chr(64 + max_col), r)).WrapText = True
                except Exception:
                    pass
                if r > 1 and len(txt) > 40:
                    ws.Rows(r).RowHeight = max(ws.Rows(r).RowHeight,
                                               18 * math.ceil(len(txt) / 58.0))
            hdr = t.get('hdr', 3)
            try:
                ws.Range('A%d:%s%d' % (hdr, chr(64 + max_col), hdr)).WrapText = True
            except Exception:
                pass

            # 7) 底部出处说明
            ann = last_row + 2
            ws.Range('A%d:%s%d' % (ann, chr(64 + max_col), ann)).Merge()
            text = ('来源：%s 〉 工作表「%s」〉 原表第 %s 行（橙色底纹＋红框标注）　｜　%s'
                    % (t['src'], t['sheet'], '、'.join(str(x) for x in t['hl']), t.get('note', '')))
            ws.Cells(ann, 1).Value = text
            ws.Cells(ann, 1).Font.Size = 10
            ws.Cells(ann, 1).Font.Italic = True
            ws.Cells(ann, 1).Font.Color = 5911370
            ws.Cells(ann, 1).HorizontalAlignment = -4131    # 左对齐
            try:
                ws.Range('A%d:%s%d' % (ann, chr(64 + max_col), ann)).WrapText = True
            except Exception:
                pass
            ws.Rows(ann).RowHeight = max(22, 17 * math.ceil(len(text) / 68.0))

            # 8) 导出 PDF → PNG
            ws.PageSetup.Orientation = 2
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = 1
            ws.PageSetup.LeftMargin = ws.PageSetup.RightMargin = 10
            ws.PageSetup.TopMargin = ws.PageSetup.BottomMargin = 10
            ws.PageSetup.CenterHeader = ''
            ws.PageSetup.CenterFooter = ''
            ws.PageSetup.PrintArea = 'A1:%s%d' % (chr(64 + max_col), ann)

            pdf = os.path.join(tmp, os.path.splitext(t['name'])[0] + '.pdf')
            wb.ExportAsFixedFormat(0, pdf)
            wb.Close(False)

            doc = pymupdf.open(pdf)
            raw = os.path.join(tmp, 'raw.png')
            doc[0].get_pixmap(dpi=dpi, alpha=False).save(raw)
            doc.close()

            im = crop_white(Image.open(raw).convert('RGB'))
            im.save(os.path.join(out_dir, t['name']))
            lines.append('OK  %-12s rows=%s  %s  %s' % (t['sheet'], new_rows, t['name'], im.size))
    finally:
        try:
            xl.Quit()
        finally:
            try:
                del ws, wb
            except Exception:
                pass
            import gc
            gc.collect()
            try:
                win32.pythoncom.CoUninitialize()
            except Exception:
                pass
            del xl
            gc.collect()

    report = '\n'.join(lines) + '\nDONE (%d 张)' % len(lines)
    with open(os.path.join(tmp, 'shots.log'), 'w', encoding='utf-8') as f:
        f.write(report + '\n')
    print(report)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--config', required=True)
    a = ap.parse_args()
    run(json.load(open(a.config, encoding='utf-8')))


if __name__ == '__main__':
    main()

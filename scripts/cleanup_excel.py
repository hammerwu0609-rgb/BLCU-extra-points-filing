# -*- coding: utf-8 -*-
"""清理 pywin32 自动化残留的「无窗口 EXCEL.EXE」进程（Windows 专用，无需第三方依赖）。

原理：pywin32 只调用 xl.Quit() 经常结束不掉 Excel 进程，每跑一次脚本就可能残留一个，
慢慢把内存吃光。但**用户自己开着的 Excel 绝不能误杀**，所以用 Win32 API 判定：
只结束「没有任何可见顶层窗口」的 EXCEL.EXE。

用法:
  python cleanup_excel.py            # 实际清理
  python cleanup_excel.py --dry-run  # 只列出将要结束的 PID，不动手
"""
import argparse, csv, ctypes, io, subprocess, sys

IMAGENAME = 'EXCEL.EXE'


def excel_pids():
    """tasklist 拿到所有 EXCEL.EXE 的 PID。"""
    out = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq %s' % IMAGENAME, '/FO', 'CSV', '/NH'],
                         capture_output=True).stdout.decode('gbk', 'ignore')
    pids = []
    for row in csv.reader(io.StringIO(out)):
        if len(row) >= 2 and row[0].strip().upper() == IMAGENAME:
            try:
                pids.append(int(row[1]))
            except ValueError:
                pass
    return pids


def pids_with_visible_window():
    """有可见顶层窗口的进程 PID 集合（用户手开的 Excel 一定在这里）。"""
    user32 = ctypes.windll.user32
    visible = set()

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

    def cb(hwnd, lparam):
        if user32.IsWindowVisible(hwnd):
            pid = ctypes.c_ulong()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value:
                visible.add(pid.value)
        return True

    user32.EnumWindows(WNDENUMPROC(cb), 0)
    return visible


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dry-run', action='store_true', help='只列出将要结束的 PID，不实际结束')
    a = ap.parse_args()

    if sys.platform != 'win32':
        print('本工具仅用于 Windows（依赖 tasklist/taskkill 与 Win32 API）。')
        return

    all_pids = excel_pids()
    if not all_pids:
        print('当前没有 %s 进程，无需清理。' % IMAGENAME)
        return

    visible = pids_with_visible_window()
    orphans = [p for p in all_pids if p not in visible]
    kept = [p for p in all_pids if p in visible]

    print('发现 %d 个 %s：' % (len(all_pids), IMAGENAME))
    for p in all_pids:
        tag = '保留（有可见窗口，可能是用户自己开的）' if p in visible else '清理（无可见窗口，自动化残留）'
        print('  PID %-8d %s' % (p, tag))

    if a.dry_run:
        print('\n[dry-run] 未做任何改动。待清理 PID: %s' % orphans)
        return
    if not orphans:
        print('\n没有需要清理的残留进程。')
        return

    r = subprocess.run(['taskkill', '/F'] + sum([['/PID', str(p)] for p in orphans], []),
                       capture_output=True)
    print('\n结束结果: %s' % r.stdout.decode('gbk', 'ignore').strip())
    left = excel_pids()
    print('剩余 %s 进程: %d 个 %s' % (IMAGENAME, len(left), left))


if __name__ == '__main__':
    main()

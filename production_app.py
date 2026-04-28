# -*- coding: utf-8 -*-
"""
생산관리 시스템 - SEO JIN PRECISION CO.
MTO (Make-to-Order) Production Management System
기계부품 주문생산 관리
"""

import tkinter as tk
from tkinter import ttk, messagebox

# matplotlib (그래프) - 설치되어 있지 않으면 그래프 메뉴만 비활성
try:
    import matplotlib
    matplotlib.use('TkAgg')
    import matplotlib.pyplot as plt
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
    # 한글 폰트: macOS / Windows / Linux 자동 감지
    import platform as _plat
    _sys = _plat.system()
    if _sys == 'Darwin':
        _kfonts = ['AppleGothic', 'Apple SD Gothic Neo', 'NanumGothic']
    elif _sys == 'Windows':
        _kfonts = ['Malgun Gothic', 'NanumGothic', 'Gulim']
    else:
        _kfonts = ['NanumGothic', 'UnDotum', 'DejaVu Sans']
    # 실제 설치된 폰트 중 첫 번째 사용
    from matplotlib import font_manager as _fm
    _avail = {f.name for f in _fm.fontManager.ttflist}
    for _f in _kfonts:
        if _f in _avail:
            matplotlib.rcParams['font.family'] = _f
            break
    matplotlib.rcParams['axes.unicode_minus'] = False
    HAS_MPL = True
except Exception:
    HAS_MPL = False
import sqlite3
import hashlib
import os
import sys
import tempfile
import webbrowser
from datetime import datetime, timedelta

# Windows 인코딩 처리
if sys.platform == "win32":
    import io
    if sys.stdout is not None:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if sys.stderr is not None:
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# ============================================================
# 상수
# ============================================================
COMPANY = "SEO JIN PRECISION CO."
ORDER_PREFIX = "SJ"
PROCESSES = ["일반CNC", "복합CNC"]

C = {
    'primary':      '#00695C',  # 진한 청록 (생산관리 컬러)
    'primary_dark': '#004D40',
    'primary_light':'#26A69A',
    'accent':       '#FF6F00',  # 주황 (강조)
    'secondary':    '#37474F',
    'success':      '#2E7D32',
    'warning':      '#E65100',
    'danger':       '#C62828',
    'urgent':       '#D32F2F',  # 납기 임박
    'bg':           '#ECEFF1',
    'white':        '#FFFFFF',
    'text':         '#212121',
    'border':       '#B0BEC5',
    'sidebar_bg':   '#ECEFF1',
    'sidebar_text': '#ECEFF1',
    'sidebar_sel':  '#00695C',
    'header_bg':    '#00695C',
    'row_even':     '#E0F2F1',
    'pass_bg':      '#E8F5E9',
    'fail_bg':      '#FFEBEE',
    'urgent_bg':    '#FFF3E0',
}

# ============================================================
# 데이터베이스
# ============================================================
class DB:
    def __init__(self):
        # PyInstaller로 묶인 EXE 실행 시 __file__은 임시 추출 폴더를 가리킴 → 종료 시 사라짐
        # → 사용자 데이터를 영구 보존하려면 EXE 옆 또는 사용자 폴더에 저장해야 함
        if getattr(sys, 'frozen', False):
            # Windows EXE: EXE 파일과 같은 폴더 (사용자가 옮기지 않는 한 영속)
            base = os.path.dirname(sys.executable)
        else:
            # Python 스크립트: 스크립트 폴더
            base = os.path.dirname(os.path.abspath(__file__))
        self.path = os.path.join(base, "production.db")
        # 디버그용으로 경로 표시
        print(f"[DB] {self.path}")
        self.conn = sqlite3.connect(self.path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._create()
        self._seed()

    def _create(self):
        c = self.conn.cursor()
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            name TEXT NOT NULL,
            team TEXT NOT NULL,
            role TEXT NOT NULL,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            contact TEXT DEFAULT '',
            phone TEXT DEFAULT '',
            email TEXT DEFAULT '',
            active INTEGER DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_no TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            spec TEXT DEFAULT '',
            material TEXT DEFAULT '',
            unit_weight REAL DEFAULT 0,
            customer_id INTEGER,
            active INTEGER DEFAULT 1,
            FOREIGN KEY(customer_id) REFERENCES customers(id)
        );
        CREATE TABLE IF NOT EXISTS equipments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            process TEXT NOT NULL,
            spec TEXT DEFAULT '',
            status TEXT DEFAULT '가동',
            active INTEGER DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_no TEXT UNIQUE NOT NULL,
            customer_id INTEGER NOT NULL,
            item_id INTEGER NOT NULL,
            quantity INTEGER NOT NULL,
            order_date TEXT NOT NULL,
            due_date TEXT NOT NULL,
            status TEXT DEFAULT '접수',
            memo TEXT DEFAULT '',
            created_by INTEGER,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(customer_id) REFERENCES customers(id),
            FOREIGN KEY(item_id) REFERENCES items(id)
        );
        CREATE TABLE IF NOT EXISTS work_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            wo_no TEXT UNIQUE NOT NULL,
            order_id INTEGER NOT NULL,
            process TEXT NOT NULL,
            equipment_id INTEGER,
            worker_id INTEGER,
            plan_qty INTEGER NOT NULL,
            done_qty INTEGER DEFAULT 0,
            defect_qty INTEGER DEFAULT 0,
            plan_start TEXT DEFAULT '',
            plan_end TEXT DEFAULT '',
            actual_start TEXT DEFAULT '',
            actual_end TEXT DEFAULT '',
            status TEXT DEFAULT '대기',
            seq INTEGER DEFAULT 1,
            memo TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(order_id) REFERENCES orders(id),
            FOREIGN KEY(equipment_id) REFERENCES equipments(id),
            FOREIGN KEY(worker_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS production_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            wo_id INTEGER NOT NULL,
            work_date TEXT NOT NULL,
            qty INTEGER NOT NULL,
            defect_qty INTEGER DEFAULT 0,
            worker_id INTEGER,
            equipment_id INTEGER,
            memo TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(wo_id) REFERENCES work_orders(id)
        );
        CREATE TABLE IF NOT EXISTS inspections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER NOT NULL,
            inspect_type TEXT DEFAULT '최종검사',
            inspect_date TEXT NOT NULL,
            sample_qty INTEGER DEFAULT 0,
            result TEXT NOT NULL,
            defect_qty INTEGER DEFAULT 0,
            defect_reason TEXT DEFAULT '',
            inspector_id INTEGER,
            memo TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(order_id) REFERENCES orders(id)
        );
        CREATE TABLE IF NOT EXISTS shipments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER NOT NULL,
            ship_no TEXT UNIQUE NOT NULL,
            ship_date TEXT NOT NULL,
            quantity INTEGER NOT NULL,
            shipped_by INTEGER,
            memo TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(order_id) REFERENCES orders(id)
        );
        """)
        self.conn.commit()

    def _seed(self):
        c = self.conn.cursor()
        # 사용자
        users = [
            ('admin',      'admin123', '관리자',     '관리',   'admin'),
            ('sales',      'sales123', '영업담당',   '영업팀', 'sales'),
            ('production', 'prod123',  '생산담당',   '생산팀', 'production'),
            ('quality',    'qual123',  '품질담당',   '품질팀', 'quality'),
            ('shipping',   'ship123',  '출하담당',   '출하팀', 'shipping'),
        ]
        for u, p, n, t, r in users:
            pw = hashlib.sha256(p.encode()).hexdigest()
            c.execute("INSERT OR IGNORE INTO users(username,password,name,team,role) VALUES(?,?,?,?,?)",
                      (u, pw, n, t, r))

        # 고객사
        customers = [
            ('CUS001', '(주)현대중공업',   '김부장', '02-1234-5678', 'kim@hd.com'),
            ('CUS002', '삼성테크윈(주)',   '이과장', '031-234-5678', 'lee@stw.com'),
            ('CUS003', '두산기계',         '박차장', '055-234-5678', 'park@dms.com'),
            ('CUS004', 'LG전자 부품사업부', '최대리', '02-9999-1111', 'choi@lge.com'),
        ]
        for row in customers:
            c.execute("INSERT OR IGNORE INTO customers(code,name,contact,phone,email) VALUES(?,?,?,?,?)", row)

        # 품목
        items = [
            ('SHF-001', '구동축',       'Dia30 x L500',  'SCM440', 2.5,  1),
            ('GER-002', '평기어',       'M2 x Z40',      'SCM415', 0.8,  1),
            ('BRK-003', '브라켓',       't10 x 100x80',  'SS400',  0.6,  2),
            ('FLG-004', '플랜지',       'Dia80 x t15',   'SUS304', 0.7,  2),
            ('PIN-005', '핀',           'Dia6 x L40',    'SUS316', 0.01, 3),
            ('CAM-006', '캠 부품',      'Dia60 x L80',   'SCM440', 1.8,  4),
        ]
        for row in items:
            c.execute("INSERT OR IGNORE INTO items(part_no,name,spec,material,unit_weight,customer_id) VALUES(?,?,?,?,?,?)", row)

        # 설비 (실제 보유 설비) - (설비명, 생산품목)
        general_cnc = [
            ('LYNX200G-S10',     'SDC80 C/BOTTOM'),
            ('E-160A-1',         'ADJUSTER류'),
            ('E-160A-2',         'ADJUSTER류'),
            ('LYNX220MA20',      'SDC80 C/BOTTOM'),
            ('LYNX220A-NT10-1',  'SDC80 C/BOTTOM'),
            ('LYNX220MA30-1',    'SDC50 CASE'),
            ('LYNX220MA30-2',    'SDC50 CASE'),
            ('LYNX2100M-1',      'SDC80 CASE'),
            ('LYNX2100M-2',      'SDC80 CASE'),
            ('LYNX2100M-3',      'SDC80 CASE'),
            ('LYNX220A20',       'VC CORE'),
            ('LYNX220A-NT10-2',  'VC CORE'),
            ('PUMATW2100M-GL',   'SDC80 C/BOTTOM'),
        ]
        compound_cnc = [
            ('SR32J-1',     'SEPERATOR'),
            ('XD-26II',     'ROD B'),
            ('XD-20II-1',   'ROD B'),
            ('XD-20II-2',   'ROD A'),
            ('XD-20II-3',   'ROD A'),
            ('XD-38II-H-1', 'CORE'),
            ('XD-38II-H-2', 'CORE'),
            ('SR32J-2',     'SEPERATOR'),
            ('SR32JN',      'SEPERATOR'),
            ('SB20R-1',     'PLUNGER'),
            ('SB20R-2',     'PLUNGER'),
            ('SB20R-3',     'PLUNGER'),
            ('SB20R-4',     'PLUNGER'),
            ('SB20R-5',     'PLUNGER'),
            ('SR38',        'CORE'),
        ]

        # 1회성 마이그레이션: 옛 샘플 데이터(DOOSAN/MAZAK/HYUNDAI/DMG)만 삭제
        # ⚠ 사용자가 추가한 설비/수정한 데이터는 절대 삭제하지 않음
        old = c.execute("""SELECT COUNT(*) FROM equipments
            WHERE name LIKE 'DOOSAN%' OR name LIKE 'MAZAK%'
               OR name LIKE 'HYUNDAI%' OR name LIKE 'DMG%'""").fetchone()[0]
        if old > 0:
            c.execute("""DELETE FROM equipments
                WHERE name LIKE 'DOOSAN%' OR name LIKE 'MAZAK%'
                   OR name LIKE 'HYUNDAI%' OR name LIKE 'DMG%'""")

        for i, (name, item) in enumerate(general_cnc, 1):
            code = f"CNC-{i:02d}"
            c.execute("INSERT OR IGNORE INTO equipments(code,name,process,spec) VALUES(?,?,?,?)",
                      (code, name, '일반CNC', item))
        for i, (name, item) in enumerate(compound_cnc, 1):
            code = f"MCT-{i:02d}"
            c.execute("INSERT OR IGNORE INTO equipments(code,name,process,spec) VALUES(?,?,?,?)",
                      (code, name, '복합CNC', item))

        self.conn.commit()

    def next_order_no(self):
        today = datetime.now().strftime("%Y%m%d")
        c = self.conn.cursor()
        c.execute("SELECT order_no FROM orders WHERE order_no LIKE ? ORDER BY order_no DESC LIMIT 1",
                  (f"{ORDER_PREFIX}-{today}-%",))
        row = c.fetchone()
        seq = int(row[0].split('-')[-1]) + 1 if row else 1
        return f"{ORDER_PREFIX}-{today}-{seq:03d}"

    def next_wo_no(self, order_no, seq):
        return f"WO-{order_no.replace(ORDER_PREFIX + '-', '')}-{seq:02d}"

    def next_ship_no(self):
        today = datetime.now().strftime("%Y%m%d")
        c = self.conn.cursor()
        c.execute("SELECT ship_no FROM shipments WHERE ship_no LIKE ? ORDER BY ship_no DESC LIMIT 1",
                  (f"SH-{today}-%",))
        row = c.fetchone()
        seq = int(row[0].split('-')[-1]) + 1 if row else 1
        return f"SH-{today}-{seq:03d}"

    def query(self, sql, params=()):
        c = self.conn.cursor()
        c.execute(sql, params)
        return c.fetchall()

    def execute(self, sql, params=()):
        c = self.conn.cursor()
        c.execute(sql, params)
        self.conn.commit()
        return c.lastrowid


# ============================================================
# UI 헬퍼
# ============================================================
def commit_inputs(widget):
    """버튼 클릭 직전 IME/StringVar 강제 동기화 (전 페이지 공용)."""
    try:
        top = widget.winfo_toplevel()
        # 현재 포커스 위젯에서 포커스를 이동시켜 IME 입력 강제 확정
        try:
            cur = top.focus_get()
            if cur is not None:
                top.focus_set()
        except Exception:
            pass
        top.update_idletasks()
        top.update()
    except Exception:
        pass

def make_btn(parent, text, cmd, color=None, **kw):
    def _wrapped():
        commit_inputs(parent)
        cmd()
    return tk.Button(parent, text=text, command=_wrapped,
                     font=('Malgun Gothic', 10, 'bold'),
                     bg=color or C['primary'], fg='white', relief='flat',
                     cursor='hand2', padx=16, pady=6, **kw)

# ─────────────────────────────────────────────
# 통일된 컬러 액션 버튼 (Frame+Label, macOS Aqua 색상 강제)
# ─────────────────────────────────────────────
BTN_THEMES = {
    'save':    ('#2E7D32', '#1B5E20'),  # 녹색 - 등록/저장
    'update':  ('#E65100', '#BF360C'),  # 주황 - 수정
    'delete':  ('#C62828', '#8E0000'),  # 빨강 - 삭제
    'action':  ('#1565C0', '#0D47A1'),  # 파랑 - 일반 액션 (생성/시작/완료)
    'export':  ('#5E35B1', '#311B92'),  # 보라 - 출력/저장
    'neutral': ('#546E7A', '#37474F'),  # 회색 - 보조
}

def color_btn(parent, text, cmd, theme='save', size=14, padx=22, pady=10):
    """통일된 디자인의 컬러 버튼 (모든 페이지 공용).

    theme: save | update | delete | action | export | neutral
    """
    bg, hover = BTN_THEMES.get(theme, BTN_THEMES['save'])
    wrap = tk.Frame(parent, bg=bg, bd=0, highlightthickness=0, cursor='hand2')
    lbl = tk.Label(wrap, text=text,
                   font=('Malgun Gothic', size, 'bold'),
                   bg=bg, fg='white', padx=padx, pady=pady, cursor='hand2')
    lbl.pack(fill='both', expand=True)

    def _click(e=None):
        commit_inputs(wrap); cmd()
    def _enter(e=None):
        wrap.config(bg=hover); lbl.config(bg=hover)
    def _leave(e=None):
        wrap.config(bg=bg); lbl.config(bg=bg)
    for w in (wrap, lbl):
        w.bind('<Button-1>', _click)
        w.bind('<Enter>', _enter)
        w.bind('<Leave>', _leave)
    return wrap

def export_tree_csv(tree, default_name='data', title='데이터 저장'):
    """Treeview 데이터를 CSV로 저장 (모든 페이지 공용 헬퍼)."""
    import os, csv, traceback
    from tkinter import filedialog, messagebox
    from datetime import datetime as _dt

    try:
        items = tree.get_children()
    except Exception as e:
        messagebox.showerror(title, f"트리 접근 실패: {e}"); return

    if not items:
        messagebox.showwarning(title,
            "저장할 데이터가 없습니다.\n\n먼저 조회 또는 등록을 해주세요.")
        return

    try:
        cols = list(tree['columns'])
    except Exception:
        cols = []

    # ID 컬럼(width=0) 제외
    display_cols, display_idx = [], []
    for i, c in enumerate(cols):
        try:
            w = tree.column(c, option='width')
            if w is None or int(w) > 0:
                display_cols.append(str(c))
                display_idx.append(i)
        except Exception:
            display_cols.append(str(c)); display_idx.append(i)

    if not display_cols:
        # 컬럼 정보가 없으면 모든 컬럼 사용
        display_cols = [str(c) for c in cols]
        display_idx = list(range(len(cols)))

    # 기본 파일명: 영문 prefix + 한글 부제 (한글이 시스템에 따라 깨질 수 있어 영문 fallback)
    ts = _dt.now().strftime('%Y%m%d_%H%M%S')
    default_name_safe = str(default_name).replace('/', '_').replace('\\', '_')
    fname = f"{default_name_safe}_{ts}.csv"

    # 기본 저장 경로: ~/Documents (없으면 home)
    init_dir = os.path.expanduser('~/Documents')
    if not os.path.isdir(init_dir):
        init_dir = os.path.expanduser('~')

    try:
        path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV (Excel)', '*.csv'), ('모든 파일', '*.*')],
            initialfile=fname,
            initialdir=init_dir,
            title=title)
    except Exception as e:
        messagebox.showerror(title, f"파일 다이얼로그 오류:\n{e}"); return

    if not path:
        return  # 사용자가 취소

    try:
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow([f"# SEO JIN PRECISION - {title}"])
            w.writerow([f"# 생성: {_dt.now().strftime('%Y-%m-%d %H:%M:%S')}"])
            w.writerow([])
            w.writerow(display_cols)
            count = 0
            for iid in items:
                vals = tree.item(iid)['values']
                if not isinstance(vals, (list, tuple)):
                    vals = [vals]
                row = []
                for i in display_idx:
                    row.append(vals[i] if i < len(vals) else '')
                w.writerow(row)
                count += 1
            w.writerow([])
            w.writerow(['총 행 수', count])
        messagebox.showinfo(title,
            f"✓ CSV 파일로 저장되었습니다.\n\n"
            f"경로: {path}\n"
            f"총 {count}건\n\n"
            f"Excel에서 더블클릭으로 열 수 있습니다.")
    except PermissionError:
        messagebox.showerror(title,
            f"저장 권한이 없습니다:\n{path}\n\n"
            "다른 폴더를 선택하거나 파일이 이미 열려있는지 확인하세요.")
    except Exception as e:
        messagebox.showerror(title,
            f"저장 실패:\n{e}\n\n{traceback.format_exc()[-500:]}")

def make_label(parent, text, bold=False, size=10, color=None, **kw):
    return tk.Label(parent, text=text,
                    font=('Malgun Gothic', size, 'bold' if bold else 'normal'),
                    fg=color or C['text'], **kw)

def make_entry(parent, var, width=18, **kw):
    return tk.Entry(parent, textvariable=var, font=('Malgun Gothic', 10),
                    relief='flat', bd=3, bg='#F5F5F5', width=width, **kw)

def make_combo(parent, var, values, width=18, state='readonly', **kw):
    return ttk.Combobox(parent, textvariable=var, values=values,
                        font=('Malgun Gothic', 10), width=width,
                        state=state, **kw)

def setup_treeview_style():
    style = ttk.Style()
    try: style.theme_use('clam')
    except: pass
    style.configure('PM.Treeview',
                    background=C['white'], foreground=C['text'],
                    rowheight=26, fieldbackground=C['white'],
                    font=('Malgun Gothic', 9))
    style.configure('PM.Treeview.Heading',
                    background=C['primary'], foreground='white',
                    font=('Malgun Gothic', 9, 'bold'), relief='flat')
    style.map('PM.Treeview', background=[('selected', C['primary_light'])])

def make_tree(parent, cols, widths, height=12):
    frame = tk.Frame(parent, bg=C['white'])
    frame.pack(fill='both', expand=True)
    tree = ttk.Treeview(frame, columns=cols, show='headings',
                        style='PM.Treeview', height=height)
    vsb = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    vsb.pack(side='right', fill='y')
    hsb.pack(side='bottom', fill='x')
    tree.pack(fill='both', expand=True)
    for col, w in zip(cols, widths):
        tree.heading(col, text=col)
        tree.column(col, width=w, anchor='center', minwidth=40)
    tree.tag_configure('even', background=C['row_even'])
    tree.tag_configure('pass_tag', background=C['pass_bg'])
    tree.tag_configure('fail_tag', background=C['fail_bg'])
    tree.tag_configure('urgent_tag', background=C['urgent_bg'])
    return tree

def fill_tree(tree, rows, tag_fn=None):
    tree.delete(*tree.get_children())
    for i, row in enumerate(rows):
        tag = tag_fn(i, row) if tag_fn else ('even' if i % 2 else '')
        tree.insert('', 'end', values=list(row), tags=(tag,))

def page_header(parent, title, sub=''):
    wrap = tk.Frame(parent, bg=C['white'])
    wrap.pack(fill='x', padx=20, pady=(18, 0))
    make_label(wrap, title, bold=True, size=16, color=C['primary'], bg=C['white']).pack(side='left', padx=12)
    if sub:
        make_label(wrap, sub, size=9, color=C['secondary'], bg=C['white']).pack(side='left')
    tk.Frame(parent, bg=C['primary'], height=2).pack(fill='x', padx=20, pady=(4, 0))

# ============================================================
# 프린터 유틸 (HP, 삼성 등 Windows 설치된 모든 프린터 지원)
# ============================================================
def list_windows_printers():
    """Windows에 설치된 프린터 목록. HP / 삼성(Samsung) / 기타 모든 프린터 자동 인식."""
    try:
        import win32print
        names = [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        default = win32print.GetDefaultPrinter()
        return names, default
    except Exception:
        return [], None


def print_file_to(printer_name, file_path):
    """특정 프린터로 파일을 직접 출력 (HP / Samsung 등)."""
    try:
        import win32api
        win32api.ShellExecute(0, "printto", file_path, f'"{printer_name}"', ".", 0)
        return True, None
    except Exception as e:
        return False, str(e)


def open_print_preview(content, title="출력"):
    """HTML 미리보기 + 인쇄 버튼. 브라우저 인쇄 대화상자에서 HP/삼성 등 어떤 프린터도 선택 가능."""
    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>{title} - SEO JIN PRECISION</title>
<style>
  @page {{ size: A4; margin: 15mm 12mm; }}
  body{{font-family:'Malgun Gothic','Nanum Gothic',sans-serif;margin:20px;font-size:11px;color:#222}}
  .hdr{{border-bottom:3px solid #00695C;padding-bottom:10px;margin-bottom:14px;
        display:flex;justify-content:space-between;align-items:flex-end}}
  .hdr .co{{font-size:14px;color:#00695C;font-weight:bold;letter-spacing:1px}}
  .hdr .tl{{font-size:22px;font-weight:bold;color:#263238}}
  .hdr .dt{{font-size:10px;color:#666}}
  pre{{white-space:pre-wrap;font-family:'Consolas','D2Coding','Courier New','Malgun Gothic';
       font-size:10.5px;line-height:1.5;background:#fff;padding:10px;border:1px solid #ccc;
       -webkit-print-color-adjust:exact;print-color-adjust:exact}}
  .toolbar{{position:sticky;top:0;background:#fff;padding:10px 0;border-bottom:1px solid #eee;
           margin-bottom:14px;display:flex;gap:8px;flex-wrap:wrap;z-index:10}}
  .pb{{background:#00695C;color:white;border:none;padding:9px 22px;
       font-size:13px;font-weight:bold;cursor:pointer;border-radius:3px}}
  .pb:hover{{background:#004D40}}
  .pb.sec{{background:#FF6F00}}
  .pb.sec:hover{{background:#E65100}}
  .info{{background:#ECEFF1;padding:8px 14px;border-left:4px solid #00695C;
        font-size:11px;color:#455A64;margin-bottom:12px;border-radius:2px}}
  @media print{{
    .toolbar,.info{{display:none !important}}
    body{{margin:0}}
    pre{{border:none;padding:0;background:#fff;font-size:10pt}}
    .hdr{{break-inside:avoid}}
  }}
</style></head><body>
<div class="toolbar">
  <button class="pb" onclick="window.print()">🖨  인쇄 (Ctrl+P)</button>
  <button class="pb sec" onclick="window.close()">✕  닫기</button>
</div>
<div class="info">
  💡 인쇄 대화상자에서 <b>HP</b>, <b>Samsung</b> 등 설치된 프린터를 선택할 수 있습니다.
     용지는 <b>A4</b>로 자동 설정됩니다.
</div>
<div class="hdr">
  <div>
    <div class="co">SEO JIN PRECISION CO.</div>
    <div class="tl">{title}</div>
  </div>
  <div class="dt">출력일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
</div>
<pre>{content}</pre>
</body></html>"""
    tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8')
    tmp.write(html); tmp.close()
    webbrowser.open(f'file://{tmp.name}')
    return tmp.name


def open_printer_dialog(content, title="출력", parent=None):
    """설치된 프린터 목록(HP/삼성 등)에서 선택해 직접 인쇄 — Windows 전용.
    win32print 없으면 기본 브라우저 미리보기로 fallback."""
    html_path = open_print_preview(content, title)

    if sys.platform != 'win32':
        return  # macOS/Linux: 브라우저 인쇄로 처리

    printers, default = list_windows_printers()
    if not printers:
        return  # win32print 미설치 — 브라우저 인쇄만 사용

    # 프린터 선택 다이얼로그
    dlg = tk.Toplevel(parent)
    dlg.title("프린터 선택 — HP / Samsung / 기타")
    dlg.geometry("460x340")
    dlg.configure(bg='white')
    dlg.transient(parent); dlg.grab_set()

    tk.Label(dlg, text="🖨  프린터 선택",
             font=('Malgun Gothic', 14, 'bold'),
             fg='#00695C', bg='white').pack(pady=(18, 4))
    tk.Label(dlg, text=f"출력 문서: {title}",
             font=('Malgun Gothic', 10), fg='#455A64', bg='white').pack()

    # HP / Samsung 구분 아이콘
    def _brand(name):
        low = name.lower()
        if 'hp' in low or 'hewlett' in low or 'laserjet' in low or 'officejet' in low or 'deskjet' in low:
            return '🖨 HP'
        if 'samsung' in low or 'xpress' in low or 'scx' in low or 'clp' in low or 'ml-' in low:
            return '🖨 Samsung'
        return '🖨'

    listf = tk.Frame(dlg, bg='white'); listf.pack(fill='both', expand=True, padx=24, pady=12)
    lb = tk.Listbox(listf, font=('Malgun Gothic', 11), height=8,
                    selectbackground='#00695C', selectforeground='white',
                    relief='flat', highlightthickness=1, highlightbackground='#CFD8DC')
    sb = ttk.Scrollbar(listf, command=lb.yview); lb.config(yscrollcommand=sb.set)
    sb.pack(side='right', fill='y'); lb.pack(side='left', fill='both', expand=True)

    for p in printers:
        mark = '  ★ (기본)' if p == default else ''
        lb.insert('end', f"  {_brand(p)}   {p}{mark}")
    if default and default in printers:
        lb.selection_set(printers.index(default))
        lb.see(printers.index(default))

    btnf = tk.Frame(dlg, bg='white'); btnf.pack(fill='x', padx=24, pady=(0, 18))
    def _do_print():
        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("선택", "프린터를 선택하세요."); return
        pname = printers[sel[0]]
        ok, err = print_file_to(pname, html_path)
        if ok:
            messagebox.showinfo("인쇄 전송", f"'{pname}'(으)로 인쇄를 전송했습니다.")
            dlg.destroy()
        else:
            messagebox.showerror("인쇄 실패", f"프린터: {pname}\n\n{err}")

    tk.Button(btnf, text="인쇄", font=('Malgun Gothic', 11, 'bold'),
              bg='#00695C', fg='white', relief='flat', cursor='hand2',
              padx=24, pady=8, command=_do_print).pack(side='right', padx=4)
    tk.Button(btnf, text="취소", font=('Malgun Gothic', 11),
              bg='#CFD8DC', fg='#263238', relief='flat', cursor='hand2',
              padx=18, pady=8, command=dlg.destroy).pack(side='right', padx=4)


# ============================================================
# 메인 앱
# ============================================================
class ProductionApp:
    def __init__(self):
        self.db = DB()
        # 로그인 없이 관리자 권한으로 자동 진입
        admin_rows = self.db.query("SELECT * FROM users WHERE username='admin'")
        if admin_rows:
            self.user = dict(admin_rows[0])
        else:
            self.user = {'id': 0, 'username': 'admin', 'name': '관리자',
                         'team': '관리부', 'role': 'admin'}
        setup_treeview_style()

        self.root = tk.Tk()
        self.root.title(f"생산관리 시스템 - {COMPANY}")
        self.root.geometry("1340x820")
        self.root.configure(bg=C['bg'])
        self.root.minsize(1100, 680)

        self._show_splash()
        self.root.mainloop()

    # ========================================================
    # 시작 화면 (서진정밀 인트로)
    # ========================================================
    def _show_splash(self):
        for w in self.root.winfo_children(): w.destroy()
        self.root.geometry("720x520")
        self.root.resizable(False, False)
        # 화면 중앙 배치
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth(); sh = self.root.winfo_screenheight()
        x = (sw - 720) // 2; y = (sh - 520) // 2
        self.root.geometry(f"720x520+{x}+{y}")

        bg = tk.Frame(self.root, bg=C['primary_dark']); bg.pack(fill='both', expand=True)

        # 상단 여백
        tk.Frame(bg, bg=C['primary_dark'], height=70).pack()

        # 로고 박스
        logo_box = tk.Frame(bg, bg=C['primary_dark']); logo_box.pack()
        tk.Label(logo_box, text="⚙", font=('Segoe UI Emoji', 84),
                 fg=C['accent'], bg=C['primary_dark']).pack()

        # 회사명 (한글)
        tk.Label(bg, text="서 진 정 밀",
                 font=('Malgun Gothic', 38, 'bold'),
                 fg='white', bg=C['primary_dark']).pack(pady=(14, 4))

        # 회사명 (영문)
        tk.Label(bg, text=COMPANY,
                 font=('Malgun Gothic', 13),
                 fg='#80CBC4', bg=C['primary_dark']).pack()

        # 구분선
        tk.Frame(bg, bg='#80CBC4', height=2, width=320).pack(pady=18)

        # 시스템명
        tk.Label(bg, text="생산관리 시스템",
                 font=('Malgun Gothic', 18, 'bold'),
                 fg='white', bg=C['primary_dark']).pack()
        tk.Label(bg, text="MTO Production Management System",
                 font=('Malgun Gothic', 10),
                 fg='#B0BEC5', bg=C['primary_dark']).pack(pady=(2, 0))

        # 로딩 점
        self._splash_dots = tk.Label(bg, text="●  ○  ○",
                                     font=('Malgun Gothic', 14),
                                     fg=C['accent'], bg=C['primary_dark'])
        self._splash_dots.pack(pady=(28, 6))

        tk.Label(bg, text="Enter 키를 누르거나 화면을 클릭하세요",
                 font=('Malgun Gothic', 11, 'bold'),
                 fg=C['accent'], bg=C['primary_dark']).pack()

        # 하단 시간 / 버전
        bot = tk.Frame(bg, bg=C['primary_dark']); bot.pack(side='bottom', fill='x', pady=14)
        tk.Label(bot, text=datetime.now().strftime("%Y-%m-%d  %H:%M"),
                 font=('Malgun Gothic', 9), fg='#78909C',
                 bg=C['primary_dark']).pack()

        # 점 애니메이션
        self._splash_step = 0
        def _animate():
            patterns = ["●  ○  ○", "○  ●  ○", "○  ○  ●", "○  ●  ○"]
            try:
                self._splash_dots.config(text=patterns[self._splash_step % 4])
            except tk.TclError:
                return  # 위젯이 이미 파괴된 경우
            self._splash_step += 1
            self.root.after(280, _animate)
        _animate()

        # Enter 키 또는 클릭으로 진입 (자동 진입 없음)
        def _enter(e=None):
            self.root.unbind('<Return>')
            self.root.unbind('<KP_Enter>')
            self.root.unbind('<Button-1>')
            self._show_main()
        self.root.bind('<Return>', _enter)
        self.root.bind('<KP_Enter>', _enter)
        self.root.bind('<Button-1>', _enter)
        self.root.focus_force()

    # ========================================================
    # 로그인
    # ========================================================
    def _show_login(self):
        for w in self.root.winfo_children(): w.destroy()
        self.root.geometry("480x600")
        self.root.resizable(False, False)

        bg = tk.Frame(self.root, bg=C['primary_dark']); bg.pack(fill='both', expand=True)

        logo = tk.Frame(bg, bg=C['primary_dark'], pady=36); logo.pack(fill='x')
        tk.Label(logo, text=COMPANY, font=('Malgun Gothic', 13, 'bold'),
                 fg='#80CBC4', bg=C['primary_dark']).pack()
        tk.Label(logo, text="생산관리 시스템", font=('Malgun Gothic', 26, 'bold'),
                 fg='white', bg=C['primary_dark']).pack(pady=6)
        tk.Label(logo, text="MTO Production Management System",
                 font=('Malgun Gothic', 11), fg='#B0BEC5', bg=C['primary_dark']).pack()

        form = tk.Frame(bg, bg='white', padx=36, pady=30); form.pack(fill='x', padx=40)

        for label_text, show_char, attr in [("아이디", '', '_lu'), ("비밀번호", '*', '_lp')]:
            tk.Label(form, text=label_text, font=('Malgun Gothic', 10),
                     bg='white', fg=C['secondary']).pack(anchor='w')
            var = tk.StringVar()
            e = tk.Entry(form, textvariable=var, show=show_char,
                         font=('Malgun Gothic', 13), relief='flat', bg='#F5F5F5', bd=5)
            e.pack(fill='x', pady=(3, 14))
            setattr(self, attr + '_var', var)
            setattr(self, attr + '_entry', e)

        self._lu_var.set('admin'); self._lp_var.set('admin123')

        tk.Button(form, text="로그인", font=('Malgun Gothic', 13, 'bold'),
                  bg=C['primary'], fg='white', relief='flat',
                  pady=12, cursor='hand2', command=self._do_login).pack(fill='x')

        hint = tk.Frame(bg, bg=C['primary_dark'], pady=18); hint.pack(fill='x', padx=40)
        tk.Label(hint,
                 text="기본 계정\nadmin/admin123 | sales/sales123 | production/prod123\nquality/qual123 | shipping/ship123",
                 font=('Malgun Gothic', 9), fg='#78909C', bg=C['primary_dark'],
                 justify='center').pack()

        self._lp_entry.bind('<Return>', lambda e: self._do_login())
        self._lu_entry.bind('<Return>', lambda e: self._lp_entry.focus())

    def _do_login(self):
        u = self._lu_var.get().strip(); p = self._lp_var.get().strip()
        if not u or not p:
            messagebox.showerror("오류", "아이디와 비밀번호를 입력하세요."); return
        pw = hashlib.sha256(p.encode()).hexdigest()
        rows = self.db.query("SELECT * FROM users WHERE username=? AND password=?", (u, pw))
        if rows:
            self.user = dict(rows[0])
            self._show_main()
        else:
            messagebox.showerror("로그인 실패", "아이디 또는 비밀번호가 올바르지 않습니다.")

    # ========================================================
    # 메인 레이아웃
    # ========================================================
    def _show_main(self):
        for w in self.root.winfo_children(): w.destroy()
        self.root.geometry("1340x820")
        self.root.resizable(True, True)

        self._build_header()
        content = tk.Frame(self.root, bg=C['bg']); content.pack(fill='both', expand=True)
        self._build_sidebar(content)

        # ── 스크롤 가능한 페이지 영역 (수직/수평 스크롤바) ──
        scroll_wrap = tk.Frame(content, bg=C['bg'])
        scroll_wrap.pack(side='left', fill='both', expand=True)

        vbar = ttk.Scrollbar(scroll_wrap, orient='vertical')
        hbar = ttk.Scrollbar(scroll_wrap, orient='horizontal')
        canvas = tk.Canvas(scroll_wrap, bg=C['bg'], highlightthickness=0,
                           yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        vbar.config(command=canvas.yview); hbar.config(command=canvas.xview)
        vbar.pack(side='right', fill='y')
        hbar.pack(side='bottom', fill='x')
        canvas.pack(side='left', fill='both', expand=True)

        self.page_area = tk.Frame(canvas, bg=C['bg'])
        self._page_window = canvas.create_window((0, 0), window=self.page_area, anchor='nw')
        self._page_canvas = canvas

        def _on_inner_configure(e):
            canvas.configure(scrollregion=canvas.bbox('all'))
        def _on_canvas_configure(e):
            # 내부 프레임의 폭을 캔버스 폭과 내부 콘텐츠 중 더 큰 값으로 (가로 스크롤 발생 시 콘텐츠 우선)
            inner_w = self.page_area.winfo_reqwidth()
            canvas.itemconfig(self._page_window, width=max(e.width, inner_w))
        self.page_area.bind('<Configure>', _on_inner_configure)
        canvas.bind('<Configure>', _on_canvas_configure)

        # 마우스 휠 (Win/Linux: <MouseWheel>, mac도 <MouseWheel>이지만 delta가 작음)
        def _on_wheel(e):
            delta = -1 if (e.delta > 0) else 1
            if sys.platform == 'darwin':
                delta = -int(e.delta)
            else:
                delta = -int(e.delta / 120)
            canvas.yview_scroll(delta, 'units')
        def _on_wheel_shift(e):
            delta = -int(e.delta / 120) if sys.platform != 'darwin' else -int(e.delta)
            canvas.xview_scroll(delta, 'units')
        canvas.bind_all('<MouseWheel>', _on_wheel)
        canvas.bind_all('<Shift-MouseWheel>', _on_wheel_shift)
        # Linux scroll buttons
        canvas.bind_all('<Button-4>', lambda e: canvas.yview_scroll(-1, 'units'))
        canvas.bind_all('<Button-5>', lambda e: canvas.yview_scroll(1, 'units'))

        self._nav('dashboard')

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=C['header_bg'], height=56); hdr.pack(fill='x'); hdr.pack_propagate(False)
        tk.Label(hdr, text=f"  {COMPANY}  |  생산관리 시스템",
                 font=('Malgun Gothic', 14, 'bold'),
                 fg='white', bg=C['header_bg']).pack(side='left', padx=10)
        tk.Label(hdr, text=f"{self.user['name']}  ({self.user['team']})",
                 font=('Malgun Gothic', 10), fg='#80CBC4', bg=C['header_bg']).pack(side='right', padx=14)
        self._time_lbl = tk.Label(hdr, font=('Malgun Gothic', 9), fg='#B0BEC5', bg=C['header_bg'])
        self._time_lbl.pack(side='right', padx=16)
        self._tick()

    def _tick(self):
        self._time_lbl.config(text=datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))
        self.root.after(1000, self._tick)

    def _build_sidebar(self, parent):
        sb = tk.Frame(parent, bg=C['sidebar_bg'], width=285); sb.pack(side='left', fill='y'); sb.pack_propagate(False)
        # (key, 라벨, 아이콘색, 이모지)
        menus = [
            ('dashboard',  '대시보드',     '#42A5F5', '🏠'),
            None,
            ('orders',     '수주 관리',    '#26A69A', '📋'),
            ('plan',       '생산계획',     '#66BB6A', '📅'),
            ('workorder',  '작업지시',     '#FFA726', '🔧'),
            ('production', '생산실적',     '#EF5350', '⚙️'),
            ('inspection', '품질검사',     '#AB47BC', '🔍'),
            ('shipment',   '출하 관리',    '#FF7043', '🚚'),
            None,
            ('report',     '보고서 / 출력','#FFCA28', '📈'),
            ('statistics', '통계 (일/월/년)','#7E57C2', '📊'),
            ('history',    '일자별 이력',   '#00838F', '📆'),
            None,
        ]
        if self.user['role'] == 'admin':
            menus += [
                ('items',      '품목 관리',    '#5C6BC0', '📦'),
                ('customers',  '고객사 관리',  '#26C6DA', '🏢'),
                ('equipments', '설비 관리',    '#8D6E63', '🏭'),
            ]

        tk.Label(sb, text="MENU", font=('Malgun Gothic', 11, 'bold'),
                 fg='#78909C', bg=C['sidebar_bg']).pack(pady=(10, 4))

        self._sb_btns = {}
        self._sb_meta = {}  # key -> (color, emoji, label)
        for item in menus:
            if item is None:
                tk.Frame(sb, bg='#B0BEC5', height=1).pack(fill='x', padx=14, pady=2); continue
            key, label, color, emoji = item
            self._sb_meta[key] = (color, emoji, label)

            # 컨테이너: 좌측 컬러 바 + 버튼
            row = tk.Frame(sb, bg=C['sidebar_bg'])
            row.pack(fill='x', padx=3, pady=0)

            bar = tk.Frame(row, bg=C['sidebar_bg'], width=5)
            bar.pack(side='left', fill='y')

            btn = tk.Button(row, text=f"  {emoji}  {label}",
                            font=('Malgun Gothic', 14, 'bold'),
                            fg='#000000', bg=C['sidebar_bg'],
                            relief='flat', anchor='w', cursor='hand2', pady=7,
                            activebackground=color, activeforeground='black',
                            command=lambda k=key: self._nav(k))
            btn.pack(side='left', fill='x', expand=True)

            # 호버 효과
            def _on_enter(e, b=btn, c=color, br=bar, k=key):
                if not getattr(self, '_active_key', None) == k:
                    b.config(bg='#CFD8DC', fg='#000000')
                    br.config(bg=c)
            def _on_leave(e, b=btn, br=bar, k=key):
                if not getattr(self, '_active_key', None) == k:
                    b.config(bg=C['sidebar_bg'], fg='#000000')
                    br.config(bg=C['sidebar_bg'])
            btn.bind('<Enter>', _on_enter)
            btn.bind('<Leave>', _on_leave)

            self._sb_btns[key] = (btn, bar)

    def _nav(self, key):
        self._active_key = key
        for k, (b, bar) in self._sb_btns.items():
            color = self._sb_meta[k][0]
            if k == key:
                b.config(bg=color, fg='#000000')
                bar.config(bg='#000000')
            else:
                b.config(bg=C['sidebar_bg'], fg='#000000')
                bar.config(bg=C['sidebar_bg'])
        for w in self.page_area.winfo_children(): w.destroy()
        # 페이지 전환 시 스크롤 위치 초기화
        if hasattr(self, '_page_canvas'):
            self._page_canvas.yview_moveto(0)
            self._page_canvas.xview_moveto(0)
        pages = {
            'dashboard':  self._pg_dashboard,
            'orders':     self._pg_orders,
            'plan':       self._pg_plan,
            'workorder':  self._pg_workorder,
            'production': self._pg_production,
            'inspection': self._pg_inspection,
            'shipment':   self._pg_shipment,
            'report':     self._pg_report,
            'statistics': self._pg_statistics,
            'history':    self._pg_history,
            'items':      self._pg_items,
            'customers':  self._pg_customers,
            'equipments': self._pg_equipments,
            'users':      self._pg_users,
        }
        if key in pages: pages[key]()

    # ========================================================
    # 대시보드
    # ========================================================
    def _pg_dashboard(self):
        p = self.page_area
        page_header(p, "대시보드", "  주문생산 현황 한눈에")

        # 통계 카드
        cards = tk.Frame(p, bg=C['bg']); cards.pack(fill='x', padx=20, pady=14)
        today = datetime.now().strftime("%Y-%m-%d")
        d7 = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")

        stats = [
            ("진행중 수주", f"SELECT COUNT(*) FROM orders WHERE status IN ('접수','진행중')", C['primary']),
            ("작업 대기",    f"SELECT COUNT(*) FROM work_orders WHERE status='대기'",          C['warning']),
            ("작업 진행",    f"SELECT COUNT(*) FROM work_orders WHERE status='진행중'",        C['accent']),
            ("이번주 납기",  f"SELECT COUNT(*) FROM orders WHERE due_date BETWEEN '{today}' AND '{d7}' AND status != '출하완료'", C['urgent']),
            ("출하 완료",    f"SELECT COUNT(*) FROM orders WHERE status='출하완료'",           C['success']),
        ]
        for title, sql, color in stats:
            cnt = self.db.query(sql)[0][0]
            card = tk.Frame(cards, bg=color, padx=16, pady=14); card.pack(side='left', expand=True, fill='both', padx=5)
            tk.Label(card, text=str(cnt), font=('Malgun Gothic', 28, 'bold'),
                     fg='white', bg=color).pack()
            tk.Label(card, text=title, font=('Malgun Gothic', 10),
                     fg='#E0E0E0', bg=color).pack()

        # 납기 임박 수주
        make_label(p, " ⚠ 납기 임박 (7일 이내)", bold=True, size=11, color=C['urgent'], bg=C['bg']).pack(anchor='w', padx=22, pady=(10, 2))
        wrap1 = tk.Frame(p, bg=C['bg']); wrap1.pack(fill='x', padx=20, pady=4)
        cols1 = ('주문번호', '고객사', '생산품명', '수량', '납기일', '상태', '진행률')
        ws1   = (135, 150, 100, 140, 60, 100, 80, 70)
        tree1 = make_tree(wrap1, cols1, ws1, height=5)

        rows1 = self.db.query(f"""
            SELECT o.order_no, c.name, i.name, o.quantity, o.due_date, o.status,
                   COALESCE((SELECT SUM(done_qty) FROM work_orders WHERE order_id=o.id),0) || '/' || o.quantity
            FROM orders o
            LEFT JOIN customers c ON o.customer_id=c.id
            LEFT JOIN items i ON o.item_id=i.id
            WHERE o.due_date BETWEEN '{today}' AND '{d7}' AND o.status != '출하완료'
            ORDER BY o.due_date
        """)
        fill_tree(tree1, rows1, lambda i, r: 'urgent_tag')

        # 진행중 작업
        make_label(p, " 진행중 작업 현황", bold=True, size=11, bg=C['bg']).pack(anchor='w', padx=22, pady=(10, 2))
        wrap2 = tk.Frame(p, bg=C['bg']); wrap2.pack(fill='both', expand=True, padx=20, pady=4)
        cols2 = ('작업번호', '주문번호', '생산품명', '공정', '설비', '계획수량', '실적수량', '불량', '상태')
        ws2   = (130, 130, 100, 80, 90, 80, 80, 60, 80)
        tree2 = make_tree(wrap2, cols2, ws2, height=10)

        rows2 = self.db.query("""
            SELECT wo.wo_no, o.order_no, i.name, wo.process, e.code,
                   wo.plan_qty, wo.done_qty, wo.defect_qty, wo.status
            FROM work_orders wo
            LEFT JOIN orders o ON wo.order_id=o.id
            LEFT JOIN items i ON o.item_id=i.id
            LEFT JOIN equipments e ON wo.equipment_id=e.id
            WHERE wo.status IN ('대기','진행중')
            ORDER BY wo.created_at DESC LIMIT 30
        """)
        def tag_wo(i, r):
            if r[8] == '진행중': return 'pass_tag'
            return 'even' if i%2 else ''
        fill_tree(tree2, rows2, tag_wo)

    # ========================================================
    # 수주 관리
    # ========================================================
    def _pg_orders(self):
        p = self.page_area
        page_header(p, "수주 관리", "  고객 주문 등록 및 관리")

        customers = self.db.query("SELECT id, name FROM customers WHERE active=1")
        items     = self.db.query("SELECT id, part_no, name FROM items WHERE active=1")
        cus_names = [r[1] for r in customers]
        item_disp = [r[2] for r in items]

        # ── 등록 폼 ──
        f = tk.Frame(p, bg='white', padx=22, pady=16); f.pack(fill='x', padx=20, pady=12)
        make_label(f, "신규 수주 등록", bold=True, size=11, color=C['primary'], bg='white').grid(
            row=0, column=0, columnspan=8, sticky='w', pady=(0, 10))

        def _lbl(r, c, t): make_label(f, t, size=9, color=C['secondary'], bg='white').grid(row=r, column=c, sticky='w', padx=6)

        # ── 주문번호 (큰 글씨, 자동 채번) ──
        order_var = tk.StringVar(value=self.db.next_order_no())
        _lbl(1, 0, "주문번호 *")
        on_entry = tk.Entry(f, textvariable=order_var,
                            font=('Malgun Gothic', 13, 'bold'), width=22,
                            fg='#00695C', relief='flat', bd=4)
        on_entry.grid(row=1, column=1, padx=4, pady=4, columnspan=2, sticky='w')

        _lbl(1, 3, "고객사 *")
        cus_var = tk.StringVar()
        cus_combo = make_combo(f, cus_var, cus_names, width=22)
        cus_combo.grid(row=1, column=4, padx=4, columnspan=2, sticky='w')

        _lbl(2, 0, "생산품명 *")
        item_var = tk.StringVar()
        item_combo = make_combo(f, item_var, item_disp, width=30)
        item_combo.grid(row=2, column=1, padx=4, columnspan=3, sticky='w', pady=4)

        _lbl(2, 4, "수량 *")
        qty_var = tk.StringVar()
        qty_entry = make_entry(f, qty_var, 12)
        qty_entry.grid(row=2, column=5, padx=4)

        # ── 날짜 입력 (단일 Entry "YYYY-MM-DD" + 자동 포맷 + 빠른 버튼) ──
        now = datetime.now()

        def _normalize(s):
            """20260428 / 2026.04.28 / 2026/04/28 → 2026-04-28 자동 변환."""
            s = (s or '').strip().replace('.', '-').replace('/', '-').replace(' ', '')
            digits = ''.join(c for c in s if c.isdigit())
            if len(digits) == 8:
                return f"{digits[:4]}-{digits[4:6]}-{digits[6:8]}"
            return s

        def _make_date_box(parent, default_dt):
            """단일 Entry 날짜 입력 (포커스 잃을 때 자동 정규화)."""
            v = tk.StringVar(value=default_dt.strftime('%Y-%m-%d'))
            ent = tk.Entry(parent, textvariable=v,
                           font=('Malgun Gothic', 12, 'bold'),
                           width=12, justify='center',
                           bg='#F5F5F5', fg='#263238',
                           relief='flat', bd=4)
            def _on_focus_out(e=None):
                v.set(_normalize(v.get()))
            ent.bind('<FocusOut>', _on_focus_out)
            ent.bind('<Return>', _on_focus_out)
            return ent, v

        _lbl(3, 0, "수주일 *")
        odate_ent, odate_var = _make_date_box(f, now)
        odate_ent.grid(row=3, column=1, padx=4, pady=6, columnspan=2, sticky='w')

        _lbl(3, 3, "납기일 *")
        ddate_ent, ddate_var = _make_date_box(f, now + timedelta(days=14))
        ddate_ent.grid(row=3, column=4, padx=4, pady=6, columnspan=2, sticky='w')

        def _sync_dates():
            """저장 직전 호출 — 자동 정규화."""
            odate_var.set(_normalize(odate_var.get()))
            ddate_var.set(_normalize(ddate_var.get()))

        # _odate_set / _ddate_set 호환 함수 (행 클릭 시 사용)
        def _odate_set(dt): odate_var.set(dt.strftime('%Y-%m-%d'))
        def _ddate_set(dt): ddate_var.set(dt.strftime('%Y-%m-%d'))

        _lbl(4, 0, "비고")
        memo_var = tk.StringVar()
        make_entry(f, memo_var, 50).grid(row=4, column=1, columnspan=5, padx=4, sticky='w', pady=4)

        # 편집 상태 추적용 (UI 라벨 없음)
        edit_id = [None]
        class _NoLbl:
            def config(self, **kw): pass
        mode_lbl = _NoLbl()

        def _clear_form():
            edit_id[0] = None
            order_var.set(self.db.next_order_no())
            for v in [cus_var, item_var, qty_var, memo_var]: v.set('')
            now2 = datetime.now()
            _odate_set(now2)
            _ddate_set(now2 + timedelta(days=14))
            _sync_dates()
            mode_lbl.config(text="👉 새 수주 입력 중", fg='#2E7D32', bg='#E8F5E9')

        def _read(widget, var):
            """Entry/Combobox에서 값 직접 읽기 (textvariable 미동기화 대비)."""
            try:
                v = widget.get()
                if v: return v.strip()
            except Exception: pass
            return (var.get() or '').strip()

        def _save_new():
            _sync_dates()
            if edit_id[0] is not None:
                if not messagebox.askyesno("확인", "현재 수정 모드입니다.\n신규 등록으로 전환할까요?"):
                    return
                _clear_form()
                return
            ono   = _read(on_entry, order_var)
            cname = _read(cus_combo, cus_var)
            iname = _read(item_combo, item_var)
            qty_s = _read(qty_entry, qty_var)
            od    = _read(odate_ent, odate_var)
            dd    = _read(ddate_ent, ddate_var)

            if not ono or not cname or not iname or not qty_s or not od or not dd:
                messagebox.showerror("입력 오류",
                    "필수 항목(*)을 모두 입력하세요.\n\n"
                    f"  · 주문번호 = '{ono}'\n"
                    f"  · 고객사 = '{cname}'\n"
                    f"  · 생산품명 = '{iname}'\n"
                    f"  · 수량 = '{qty_s}'\n"
                    f"  · 수주일 = '{od}'\n"
                    f"  · 납기일 = '{dd}'")
                return
            try: q = int(qty_s)
            except: messagebox.showerror("오류", f"수량은 정수로 입력하세요. (현재: '{qty_s}')"); return

            # 중복 주문번호 방지
            dup = self.db.query("SELECT id FROM orders WHERE order_no=?", (ono,))
            if dup:
                messagebox.showerror("오류", f"이미 존재하는 주문번호입니다: {ono}\n새 번호를 입력하세요."); return

            # 고객사/품목 매칭 — 드롭다운에 없는 값이면 직접 검색
            cus_id = None
            if cname in cus_names:
                cus_id = customers[cus_names.index(cname)][0]
            else:
                row = self.db.query("SELECT id FROM customers WHERE name=? AND active=1", (cname,))
                if row: cus_id = row[0][0]
            if cus_id is None:
                messagebox.showerror("고객사 오류",
                    f"'{cname}' 고객사를 찾을 수 없습니다.\n\n"
                    "고객사 관리 메뉴에서 먼저 등록하세요.\n"
                    f"등록된 고객사: {', '.join(cus_names) if cus_names else '(없음)'}")
                return

            item_id = None
            if iname in item_disp:
                item_id = items[item_disp.index(iname)][0]
            else:
                row = self.db.query("SELECT id FROM items WHERE name=? AND active=1", (iname,))
                if row: item_id = row[0][0]
            if item_id is None:
                messagebox.showerror("품목 오류",
                    f"'{iname}' 품목을 찾을 수 없습니다.\n\n"
                    "품목 관리 메뉴에서 먼저 등록하세요.\n"
                    f"등록된 품목: {', '.join(item_disp[:8]) if item_disp else '(없음)'}{' ...' if len(item_disp)>8 else ''}")
                return

            try:
                self.db.execute("""
                    INSERT INTO orders(order_no,customer_id,item_id,quantity,order_date,due_date,memo,created_by)
                    VALUES(?,?,?,?,?,?,?,?)
                """, (ono, cus_id, item_id, q, od, dd,
                      memo_var.get(), self.user['id']))
            except Exception as e:
                messagebox.showerror("DB 저장 실패", f"오류: {e}\n\n주문번호: {ono}")
                return
            messagebox.showinfo("등록 완료", f"수주 [{ono}] 등록 완료!")
            _clear_form()
            _load()

        def _update():
            _sync_dates()
            if edit_id[0] is None:
                messagebox.showwarning("수정", "수정할 수주를 목록에서 선택하세요."); return
            ono   = _read(on_entry, order_var)
            cname = _read(cus_combo, cus_var)
            iname = _read(item_combo, item_var)
            qty_s = _read(qty_entry, qty_var)
            od    = _read(odate_ent, odate_var)
            dd    = _read(ddate_ent, ddate_var)
            if not all([cname, iname, qty_s, od, dd]):
                messagebox.showerror("오류", "필수 항목(*)을 모두 입력하세요."); return
            try: q = int(qty_s)
            except: messagebox.showerror("오류", "수량은 정수로 입력하세요."); return
            # 고객사/품목 매칭
            cus_id = None
            if cname in cus_names:
                cus_id = customers[cus_names.index(cname)][0]
            else:
                row = self.db.query("SELECT id FROM customers WHERE name=? AND active=1", (cname,))
                if row: cus_id = row[0][0]
            item_id = None
            if iname in item_disp:
                item_id = items[item_disp.index(iname)][0]
            else:
                row = self.db.query("SELECT id FROM items WHERE name=? AND active=1", (iname,))
                if row: item_id = row[0][0]
            if cus_id is None or item_id is None:
                messagebox.showerror("오류", "고객사 또는 품목을 찾을 수 없습니다."); return
            if not messagebox.askyesno("확인", f"수주 [{ono}] 의 내용을 수정하시겠습니까?"):
                return
            self.db.execute("""
                UPDATE orders SET customer_id=?, item_id=?, quantity=?,
                                  order_date=?, due_date=?, memo=?
                WHERE id=?
            """, (cus_id, item_id, q, od, dd, memo_var.get(), edit_id[0]))
            messagebox.showinfo("완료", f"수주 [{ono}] 수정 완료!")
            _clear_form()
            _load()

        def _delete_selected():
            ono = order_var.get()
            row = self.db.query("SELECT id, status FROM orders WHERE order_no=?", (ono,))
            if not row:
                messagebox.showwarning("삭제", "삭제할 수주를 목록에서 선택하세요."); return
            oid, status = row[0][0], row[0][1]
            wo = self.db.query("SELECT COUNT(*) FROM work_orders WHERE order_id=?", (oid,))[0][0]
            sh = self.db.query("SELECT COUNT(*) FROM shipments WHERE order_id=?", (oid,))[0][0]
            warn = ""
            if wo or sh:
                warn = f"\n\n⚠ 연관 데이터: 작업지시 {wo}건 / 출하 {sh}건\n   함께 삭제됩니다."
            if not messagebox.askyesno("삭제 확인",
                f"수주 [{ono}] (상태: {status}) 를 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다.{warn}"):
                return
            self.db.execute("DELETE FROM inspections WHERE order_id=?", (oid,))
            self.db.execute("DELETE FROM shipments WHERE order_id=?", (oid,))
            self.db.execute("DELETE FROM production_records WHERE wo_id IN (SELECT id FROM work_orders WHERE order_id=?)", (oid,))
            self.db.execute("DELETE FROM work_orders WHERE order_id=?", (oid,))
            self.db.execute("DELETE FROM orders WHERE id=?", (oid,))
            messagebox.showinfo("삭제 완료", f"수주 [{ono}] 가 삭제되었습니다.")
            _clear_form()
            _load()

        # 통일된 컬러 액션 버튼
        btn_row = tk.Frame(f, bg='white')
        btn_row.grid(row=5, column=0, columnspan=8, sticky='e', pady=(12, 0))
        color_btn(btn_row, "수주 등록", _save_new, theme='save').pack(side='right', padx=5)
        color_btn(btn_row, "수주 수정", _update, theme='update').pack(side='right', padx=5)
        color_btn(btn_row, "수주 삭제", _delete_selected, theme='delete').pack(side='right', padx=5)
        # 초기화 — 화면 우측 하단
        _reset_bar = tk.Frame(p, bg=C['bg']); _reset_bar.pack(side='bottom', fill='x', padx=20, pady=8)
        color_btn(_reset_bar, "초기화", _clear_form, theme='neutral', size=9, padx=10, pady=5).pack(side='right')
        # Esc 키로도 초기화
        self.root.bind('<Escape>', lambda e: _clear_form())

        # ── 목록 ──
        make_label(p, " 수주 목록  (행을 클릭하면 수정/삭제 모드로 전환)",
                   bold=True, size=11, bg=C['bg']).pack(anchor='w', padx=22, pady=(4, 2))
        wrap = tk.Frame(p, bg=C['bg']); wrap.pack(fill='both', expand=True, padx=20, pady=4)
        cols = ('주문번호', '고객사', '생산품명', '재질', '수량', '수주일', '납기일', '상태')
        ws   = (135, 150, 100, 140, 80, 60, 100, 100, 80)
        tree = make_tree(wrap, cols, ws, height=14)

        def _load():
            rows = self.db.query("""
                SELECT o.order_no, c.name, i.name, i.material, o.quantity,
                       o.order_date, o.due_date, o.status
                FROM orders o
                LEFT JOIN customers c ON o.customer_id=c.id
                LEFT JOIN items i ON o.item_id=i.id
                ORDER BY o.created_at DESC
            """)
            today = datetime.now().strftime("%Y-%m-%d")
            d7 = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
            def tag(i, r):
                st = r[7]; due = r[6]
                if st == '출하완료': return 'pass_tag'
                if today <= due <= d7: return 'urgent_tag'
                return 'even' if i%2 else ''
            fill_tree(tree, rows, tag)

        def _on_select(e):
            s = tree.selection()
            if not s: return
            v = tree.item(s[0])['values']
            ono = v[0]
            row = self.db.query("""
                SELECT o.id, o.order_no, c.name, i.name,
                       o.quantity, o.order_date, o.due_date, o.memo
                FROM orders o
                LEFT JOIN customers c ON o.customer_id=c.id
                LEFT JOIN items i ON o.item_id=i.id
                WHERE o.order_no=?
            """, (ono,))
            if not row: return
            r = row[0]
            edit_id[0] = r[0]
            order_var.set(r[1])
            cus_var.set(r[2] or '')
            # item_disp 에서 정확 매칭되는 항목을 찾아 셋
            for d in item_disp:
                if d == r[3]:
                    item_var.set(d); break
            else:
                item_var.set(r[3] or '')
            qty_var.set(str(r[4] or ''))
            odate_var.set(r[5] or '')
            ddate_var.set(r[6] or '')
            # Spinbox 날짜 위젯에도 반영
            try:
                if r[5]:
                    od = datetime.strptime(str(r[5])[:10], '%Y-%m-%d')
                    _odate_set(od)
                if r[6]:
                    dd = datetime.strptime(str(r[6])[:10], '%Y-%m-%d')
                    _ddate_set(dd)
            except Exception:
                pass
            memo_var.set(r[7] or '')
            mode_lbl.config(text=f"✏ 기존 수주 [{r[1]}] 수정 중",
                            fg='#E65100', bg='#FFF3E0')
        tree.bind('<<TreeviewSelect>>', _on_select)

        _load()

    # ========================================================
    # 생산계획 (수주 → 작업지시 생성)
    # ========================================================
    def _pg_plan(self):
        p = self.page_area
        page_header(p, "생산계획", "  수주를 작업지시로 변환")

        info = tk.Frame(p, bg='white', padx=20, pady=10); info.pack(fill='x', padx=20, pady=12)
        make_label(info, "수주 선택 → 공정별로 작업지시(WO)를 자동 생성합니다 (일반CNC + 복합CNC 순)",
                   size=10, color=C['secondary'], bg='white').pack(anchor='w')

        wrap = tk.Frame(p, bg=C['bg']); wrap.pack(fill='both', expand=True, padx=20, pady=4)
        cols = ('OID', '주문번호', '고객사', '생산품명', '수량', '납기일', '상태', 'WO수')
        ws   = (0, 135, 150, 100, 150, 60, 100, 80, 60)
        tree = make_tree(wrap, cols, ws, height=14)
        tree.column('OID', width=0, stretch=False)

        def _load():
            rows = self.db.query("""
                SELECT o.id, o.order_no, c.name, i.name, o.quantity, o.due_date, o.status,
                       (SELECT COUNT(*) FROM work_orders WHERE order_id=o.id)
                FROM orders o
                LEFT JOIN customers c ON o.customer_id=c.id
                LEFT JOIN items i ON o.item_id=i.id
                WHERE o.status IN ('접수','진행중')
                ORDER BY o.due_date
            """)
            def tag(i, r):
                if r[7] > 0: return 'pass_tag'
                return 'even' if i%2 else ''
            fill_tree(tree, rows, tag)

        _load()

        # 액션 영역
        af = tk.Frame(p, bg=C['bg']); af.pack(fill='x', padx=20, pady=10)
        make_label(af, "선택한 수주 → 작업지시 생성", size=10, color=C['secondary'], bg=C['bg']).pack(side='left')

        def _generate():
            sel = tree.selection()
            if not sel:
                messagebox.showerror("오류", "수주를 선택하세요."); return
            v = tree.item(sel[0])['values']
            order_id, order_no, qty = v[0], v[1], v[4]
            existing = self.db.query("SELECT COUNT(*) FROM work_orders WHERE order_id=?", (order_id,))[0][0]
            if existing > 0:
                if not messagebox.askyesno("확인", f"이미 {existing}건의 작업지시가 있습니다. 추가로 생성할까요?"):
                    return
            for idx, proc in enumerate(PROCESSES, start=1):
                seq = existing + idx
                wo_no = self.db.next_wo_no(order_no, seq)
                self.db.execute("""
                    INSERT INTO work_orders(wo_no,order_id,process,plan_qty,seq)
                    VALUES(?,?,?,?,?)
                """, (wo_no, order_id, proc, qty, seq))
            self.db.execute("UPDATE orders SET status='진행중' WHERE id=?", (order_id,))
            messagebox.showinfo("완료", f"작업지시 {len(PROCESSES)}건 생성 완료!\n(일반CNC, 복합CNC)")
            _load()

        color_btn(af, "작업지시 자동 생성", _generate, theme='action').pack(side='right', padx=4)

    # ========================================================
    # 작업지시
    # ========================================================
    def _pg_workorder(self):
        p = self.page_area
        page_header(p, "작업지시", "  공정별 작업 배정 / 시작 / 완료")

        equipments = self.db.query("SELECT id, code, name, process FROM equipments WHERE active=1")
        workers    = self.db.query("SELECT id, name FROM users WHERE role IN ('production','admin')")

        wrap = tk.Frame(p, bg=C['bg']); wrap.pack(fill='both', expand=True, padx=20, pady=10)
        cols = ('WID', '작업번호', '주문번호', '생산품명', '공정', '계획수량', '실적', '불량', '설비', '담당', '상태')
        ws   = (0, 135, 130, 100, 130, 75, 70, 60, 50, 90, 80, 75)
        tree = make_tree(wrap, cols, ws, height=15)
        tree.column('WID', width=0, stretch=False)

        def _load():
            rows = self.db.query("""
                SELECT wo.id, wo.wo_no, o.order_no, i.name, wo.process,
                       wo.plan_qty, wo.done_qty, wo.defect_qty,
                       COALESCE(e.code,'-'), COALESCE(u.name,'-'), wo.status
                FROM work_orders wo
                LEFT JOIN orders o ON wo.order_id=o.id
                LEFT JOIN items i ON o.item_id=i.id
                LEFT JOIN equipments e ON wo.equipment_id=e.id
                LEFT JOIN users u ON wo.worker_id=u.id
                ORDER BY wo.created_at DESC
            """)
            def tag(i, r):
                if r[10] == '완료': return 'pass_tag'
                if r[10] == '진행중': return 'urgent_tag'
                return 'even' if i%2 else ''
            fill_tree(tree, rows, tag)

        _load()

        # 액션
        af = tk.Frame(p, bg='white', padx=22, pady=14); af.pack(fill='x', padx=20, pady=8)
        sel_var = tk.StringVar(value="작업을 선택하세요")
        sel_id  = [None]; sel_proc = [None]

        # 모든 설비 목록 (일반CNC + 복합CNC 통합, 공정 표시 포함)
        all_eq_codes = [f"{r[1]} - {r[2]} ({r[3]})" for r in equipments]

        def _on_sel(e):
            s = tree.selection()
            if not s: return
            v = tree.item(s[0])['values']
            sel_id[0] = v[0]; sel_proc[0] = v[4]
            sel_var.set(f"  {v[1]}  ({v[3]} / {v[4]} / 계획 {v[5]})")
            # 설비 콤보: 모든 설비 표시 (사용자가 자유 선택)
            equip_combo['values'] = all_eq_codes

        tree.bind('<<TreeviewSelect>>', _on_sel)

        make_label(af, "작업 배정 / 상태 변경", bold=True, size=11, color=C['primary'], bg='white').grid(
            row=0, column=0, columnspan=6, sticky='w')
        make_label(af, "선택:", size=9, color=C['secondary'], bg='white').grid(row=1, column=0, sticky='w', pady=4)
        tk.Label(af, textvariable=sel_var, font=('Malgun Gothic', 10, 'bold'),
                 fg=C['primary'], bg='white').grid(row=1, column=1, columnspan=5, sticky='w')

        def _lbl(r, c, t): make_label(af, t, size=9, color=C['secondary'], bg='white').grid(row=r, column=c, sticky='w', padx=6, pady=3)

        _lbl(2, 0, "설비 배정")
        eq_var = tk.StringVar()
        equip_combo = make_combo(af, eq_var, [], width=24)
        equip_combo.grid(row=2, column=1, padx=4)

        _lbl(2, 2, "담당자")
        worker_var = tk.StringVar()
        worker_names = [w[1] for w in workers]
        worker_combo = make_combo(af, worker_var, worker_names, width=14)
        worker_combo.grid(row=2, column=3, padx=4)

        # ── 계획수량 수정 ──
        _lbl(3, 0, "계획수량")
        plan_var = tk.StringVar()
        plan_entry = make_entry(af, plan_var, 12)
        plan_entry.grid(row=3, column=1, padx=4, pady=4)

        # 행 선택 시 계획수량도 폼에 채우기
        def _on_sel2(e=None):
            s = tree.selection()
            if not s: return
            v = tree.item(s[0])['values']
            # v[5]=계획수량
            try:
                plan_var.set(str(v[5] or ''))
            except Exception: pass
        tree.bind('<<TreeviewSelect>>', _on_sel2, add='+')

        def _assign():
            if not sel_id[0]:
                messagebox.showerror("오류", "작업을 선택하세요."); return
            eq_id = None
            ev = (equip_combo.get() or eq_var.get() or '').strip()
            if ev:
                code = ev.split(' - ')[0]
                row = self.db.query("SELECT id FROM equipments WHERE code=?", (code,))
                if row: eq_id = row[0][0]
            wid = None
            wv = (worker_combo.get() or worker_var.get() or '').strip()
            if wv:
                row = self.db.query("SELECT id FROM users WHERE name=?", (wv,))
                if row: wid = row[0][0]
            self.db.execute("UPDATE work_orders SET equipment_id=?, worker_id=? WHERE id=?",
                            (eq_id, wid, sel_id[0]))
            messagebox.showinfo("완료", "설비/담당자 배정 완료!")
            _load()

        def _start():
            if not sel_id[0]: messagebox.showerror("오류", "작업을 선택하세요."); return
            self.db.execute("""UPDATE work_orders SET status='진행중',
                               actual_start=? WHERE id=?""",
                            (datetime.now().strftime("%Y-%m-%d %H:%M"), sel_id[0]))
            messagebox.showinfo("완료", "작업 시작 처리 완료!")
            _load()

        def _complete():
            if not sel_id[0]: messagebox.showerror("오류", "작업을 선택하세요."); return
            self.db.execute("""UPDATE work_orders SET status='완료',
                               actual_end=? WHERE id=?""",
                            (datetime.now().strftime("%Y-%m-%d %H:%M"), sel_id[0]))
            messagebox.showinfo("완료", "작업 완료 처리!")
            _load()

        def _wo_modify():
            """WO의 계획수량 수정."""
            if not sel_id[0]:
                messagebox.showerror("오류", "수정할 작업을 선택하세요."); return
            new_plan = (plan_entry.get() or plan_var.get() or '').strip()
            if not new_plan:
                messagebox.showerror("오류", "계획수량을 입력하세요."); return
            try: q = int(new_plan)
            except: messagebox.showerror("오류", "계획수량은 정수로 입력하세요."); return
            if not messagebox.askyesno("확인",
                f"계획수량을 {q}(으)로 수정하시겠습니까?"):
                return
            self.db.execute("UPDATE work_orders SET plan_qty=? WHERE id=?",
                            (q, sel_id[0]))
            messagebox.showinfo("완료", "작업지시 수정 완료!")
            _load()

        def _wo_delete():
            """WO 및 연관 생산실적 삭제."""
            if not sel_id[0]:
                messagebox.showerror("오류", "삭제할 작업을 선택하세요."); return
            pr_cnt = self.db.query("SELECT COUNT(*) FROM production_records WHERE wo_id=?", (sel_id[0],))[0][0]
            warn = ""
            if pr_cnt:
                warn = f"\n\n⚠ 연관 생산실적 {pr_cnt}건도 함께 삭제됩니다."
            if not messagebox.askyesno("삭제 확인",
                f"이 작업지시를 삭제하시겠습니까?{warn}\n이 작업은 되돌릴 수 없습니다."):
                return
            self.db.execute("DELETE FROM production_records WHERE wo_id=?", (sel_id[0],))
            self.db.execute("DELETE FROM work_orders WHERE id=?", (sel_id[0],))
            messagebox.showinfo("삭제 완료", "작업지시가 삭제되었습니다.")
            sel_id[0] = None
            sel_var.set("작업을 선택하세요")
            _load()

        # 5개 버튼: 배정 저장 / 작업 시작 / 작업 완료 / WO 수정 / WO 삭제
        bf = tk.Frame(af, bg='white')
        bf.grid(row=4, column=0, columnspan=6, pady=(14, 4), sticky='e', padx=4)
        def _clear_wo_form():
            sel_id[0] = None
            sel_var.set("작업을 선택하세요")
            eq_var.set(''); worker_var.set(''); plan_var.set('')

        color_btn(bf, "배정 저장", _assign, theme='save').pack(side='left', padx=4)
        color_btn(bf, "작업 시작", _start, theme='action').pack(side='left', padx=4)
        color_btn(bf, "작업 완료", _complete, theme='update').pack(side='left', padx=4)
        color_btn(bf, "WO 수정", _wo_modify, theme='update').pack(side='left', padx=4)
        color_btn(bf, "WO 삭제", _wo_delete, theme='delete').pack(side='left', padx=4)
        # 초기화 — 화면 우측 하단
        _reset_bar = tk.Frame(p, bg=C['bg']); _reset_bar.pack(side='bottom', fill='x', padx=20, pady=8)
        color_btn(_reset_bar, "초기화", _clear_wo_form, theme='neutral', size=9, padx=10, pady=5).pack(side='right')
        self.root.bind('<Escape>', lambda e: _clear_wo_form())

    # ========================================================
    # 생산실적
    # ========================================================
    def _pg_production(self):
        p = self.page_area
        page_header(p, "생산실적", "  공정별 생산 수량 입력 / 수정 / 삭제")

        # ── 1. 작업지시(WO) 선택 트리 ──
        make_label(p, " 1. 작업 선택", bold=True, size=11, bg=C['bg']).pack(anchor='w', padx=22, pady=(4, 2))
        wrap = tk.Frame(p, bg=C['bg']); wrap.pack(fill='x', padx=20, pady=4)
        cols = ('WID', '작업번호', '생산품명', '공정', '설비', '계획', '실적', '불량', '진행률', '상태')
        ws   = (0, 130, 100, 140, 75, 90, 60, 60, 50, 70, 75)
        tree = make_tree(wrap, cols, ws, height=8)
        tree.column('WID', width=0, stretch=False)

        def _load():
            rows = self.db.query("""
                SELECT wo.id, wo.wo_no, i.name, wo.process,
                       COALESCE(e.code,'-'),
                       wo.plan_qty, wo.done_qty, wo.defect_qty,
                       CAST(wo.done_qty * 100.0 / wo.plan_qty AS INTEGER) || '%',
                       wo.status
                FROM work_orders wo
                LEFT JOIN orders o ON wo.order_id=o.id
                LEFT JOIN items i ON o.item_id=i.id
                LEFT JOIN equipments e ON wo.equipment_id=e.id
                WHERE wo.status IN ('진행중','완료')
                ORDER BY wo.actual_start DESC
            """)
            def tag(i, r):
                if r[9] == '완료': return 'pass_tag'
                return 'urgent_tag' if r[9] == '진행중' else ('even' if i%2 else '')
            fill_tree(tree, rows, tag)

        # ── 2. 선택한 작업의 생산실적 이력 트리 ──
        make_label(p, " 2. 실적 이력  (행 클릭 → 수정/삭제 모드)",
                   bold=True, size=11, bg=C['bg']).pack(anchor='w', padx=22, pady=(8, 2))
        rwrap = tk.Frame(p, bg=C['bg']); rwrap.pack(fill='x', padx=20, pady=4)
        rcols = ('PID', '작업일', '양품', '불량', '담당', '비고', '등록시각')
        rws   = (0, 110, 80, 70, 90, 250, 140)
        rtree = make_tree(rwrap, rcols, rws, height=6)
        rtree.column('PID', width=0, stretch=False)

        def _load_records():
            for x in rtree.get_children(): rtree.delete(x)
            if not sel_id[0]: return
            rows = self.db.query("""
                SELECT pr.id, pr.work_date, pr.qty, pr.defect_qty,
                       COALESCE(u.name,'-'), COALESCE(pr.memo,''), pr.created_at
                FROM production_records pr
                LEFT JOIN users u ON pr.worker_id=u.id
                WHERE pr.wo_id=?
                ORDER BY pr.work_date DESC, pr.id DESC
            """, (sel_id[0],))
            fill_tree(rtree, rows, lambda i, r: 'even' if i%2 else '')

        # 입력 폼
        f = tk.Frame(p, bg='white', padx=22, pady=14); f.pack(fill='x', padx=20, pady=8)
        sel_var = tk.StringVar(value="작업을 선택하세요")
        sel_id  = [None]
        edit_pid = [None]   # 수정 대상 PR id

        def _clear_form():
            edit_pid[0] = None
            dt_v.set(datetime.now().strftime("%Y-%m-%d"))
            q_v.set(''); d_v.set('0'); m_v.set('')
            mode_lbl.config(text="신규 실적 입력 중", fg='#2E7D32', bg='#E8F5E9')
            rtree.selection_remove(rtree.selection())

        def _on_sel(e):
            s = tree.selection()
            if not s: return
            v = tree.item(s[0])['values']
            sel_id[0] = v[0]
            sel_var.set(f"  {v[1]}  ({v[2]} / {v[3]} / {v[4]} / 계획 {v[5]})")
            _clear_form()
            _load_records()

        tree.bind('<<TreeviewSelect>>', _on_sel)

        def _on_record_select(e):
            s = rtree.selection()
            if not s: return
            v = rtree.item(s[0])['values']
            edit_pid[0] = v[0]
            dt_v.set(str(v[1] or ''))
            q_v.set(str(v[2] or ''))
            d_v.set(str(v[3] or '0'))
            m_v.set(str(v[5] or ''))
            mode_lbl.config(text=f"실적 [{v[1]}] 수정 중", fg='#E65100', bg='#FFF3E0')
        rtree.bind('<<TreeviewSelect>>', _on_record_select)

        make_label(f, "생산실적 입력", bold=True, size=11, color=C['primary'], bg='white').grid(row=0, column=0, columnspan=6, sticky='w')
        mode_lbl = tk.Label(f, text="신규 실적 입력 중",
                            font=('Malgun Gothic', 10, 'bold'),
                            bg='#E8F5E9', fg='#2E7D32', padx=10, pady=3)
        mode_lbl.grid(row=0, column=6, columnspan=2, sticky='e', padx=8)

        make_label(f, "선택:", size=9, color=C['secondary'], bg='white').grid(row=1, column=0, sticky='w', pady=4)
        tk.Label(f, textvariable=sel_var, font=('Malgun Gothic', 10, 'bold'),
                 fg=C['primary'], bg='white').grid(row=1, column=1, columnspan=5, sticky='w')

        def _lbl(r, c, t): make_label(f, t, size=9, color=C['secondary'], bg='white').grid(row=r, column=c, sticky='w', padx=6, pady=3)

        _lbl(2, 0, "작업일 *");    dt_v = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        dt_e = make_entry(f, dt_v, 14); dt_e.grid(row=2, column=1, padx=4)
        _lbl(2, 2, "양품 수량 *"); q_v  = tk.StringVar()
        q_e = make_entry(f, q_v, 10); q_e.grid(row=2, column=3, padx=4)
        _lbl(2, 4, "불량 수량");   d_v  = tk.StringVar(value='0')
        d_e = make_entry(f, d_v, 10); d_e.grid(row=2, column=5, padx=4)
        _lbl(3, 0, "비고");         m_v  = tk.StringVar()
        m_e = make_entry(f, m_v, 40); m_e.grid(row=3, column=1, columnspan=4, padx=4, sticky='w')

        def _read(widget, var, default=''):
            try:
                v = widget.get()
                if v: return v.strip()
            except Exception: pass
            return (var.get() or default).strip()

        def _refresh_totals():
            """선택된 WO의 누적 실적 재계산."""
            if not sel_id[0]: return
            tot = self.db.query("SELECT SUM(qty), SUM(defect_qty) FROM production_records WHERE wo_id=?", (sel_id[0],))[0]
            self.db.execute("UPDATE work_orders SET done_qty=?, defect_qty=? WHERE id=?",
                            (tot[0] or 0, tot[1] or 0, sel_id[0]))

        def _save_new():
            if not sel_id[0]:
                messagebox.showerror("오류", "작업을 선택하세요."); return
            if edit_pid[0] is not None:
                if not messagebox.askyesno("확인", "현재 수정 모드입니다.\n신규 등록으로 전환할까요?"):
                    return
                _clear_form(); return
            q_s = _read(q_e, q_v); d_s = _read(d_e, d_v, '0')
            dt_s = _read(dt_e, dt_v); m_s = _read(m_e, m_v)
            if not q_s:
                messagebox.showerror("입력 오류", f"양품 수량을 입력하세요. (현재: '{q_s}')"); return
            try: q = int(q_s); d = int(d_s or 0)
            except:
                messagebox.showerror("오류",
                    f"수량은 정수로 입력하세요.\n양품:'{q_s}' 불량:'{d_s}'"); return
            wo = self.db.query("SELECT equipment_id, worker_id FROM work_orders WHERE id=?", (sel_id[0],))[0]
            self.db.execute("""
                INSERT INTO production_records(wo_id,work_date,qty,defect_qty,worker_id,equipment_id,memo)
                VALUES(?,?,?,?,?,?,?)
            """, (sel_id[0], dt_s, q, d, wo[1] or self.user['id'], wo[0], m_s))
            _refresh_totals()
            messagebox.showinfo("완료", f"생산실적 등록! (양품 {q}, 불량 {d})")
            _clear_form(); _load(); _load_records()

        def _update():
            if edit_pid[0] is None:
                messagebox.showwarning("수정", "수정할 실적을 이력에서 선택하세요."); return
            q_s = _read(q_e, q_v); d_s = _read(d_e, d_v, '0')
            dt_s = _read(dt_e, dt_v); m_s = _read(m_e, m_v)
            if not q_s:
                messagebox.showerror("입력 오류", "양품 수량을 입력하세요."); return
            try: q = int(q_s); d = int(d_s or 0)
            except: messagebox.showerror("오류", "수량은 정수로 입력하세요."); return
            if not messagebox.askyesno("확인",
                f"실적을 수정하시겠습니까?\n작업일: {dt_s}\n양품: {q} / 불량: {d}"):
                return
            self.db.execute("""
                UPDATE production_records
                SET work_date=?, qty=?, defect_qty=?, memo=?
                WHERE id=?
            """, (dt_s, q, d, m_s, edit_pid[0]))
            _refresh_totals()
            messagebox.showinfo("완료", "실적 수정 완료!")
            _clear_form(); _load(); _load_records()

        def _delete():
            if edit_pid[0] is None:
                messagebox.showwarning("삭제", "삭제할 실적을 이력에서 선택하세요."); return
            if not messagebox.askyesno("삭제 확인",
                f"이 실적을 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다."):
                return
            self.db.execute("DELETE FROM production_records WHERE id=?", (edit_pid[0],))
            _refresh_totals()
            messagebox.showinfo("삭제 완료", "실적이 삭제되었습니다.")
            _clear_form(); _load(); _load_records()

        # 3개 액션 버튼 (우측 정렬)
        bf = tk.Frame(f, bg='white')
        bf.grid(row=4, column=0, columnspan=8, pady=(12, 4), sticky='e')
        color_btn(bf, "실적 등록", _save_new, theme='save').pack(side='right', padx=5)
        color_btn(bf, "실적 수정", _update, theme='update').pack(side='right', padx=5)
        color_btn(bf, "실적 삭제", _delete, theme='delete').pack(side='right', padx=5)
        # 초기화 — 화면 우측 하단
        _reset_bar = tk.Frame(p, bg=C['bg']); _reset_bar.pack(side='bottom', fill='x', padx=20, pady=8)
        color_btn(_reset_bar, "초기화", _clear_form, theme='neutral', size=9, padx=10, pady=5).pack(side='right')
        self.root.bind('<Escape>', lambda e: _clear_form())

        _load()

    # ========================================================
    # 품질검사
    # ========================================================
    def _pg_inspection(self):
        p = self.page_area
        page_header(p, "품질검사", "  자주검사 / 최종검사 — 등록 / 수정 / 삭제")

        # ── 1. 주문 선택 트리 ──
        make_label(p, " 1. 주문 선택", bold=True, size=11, bg=C['bg']).pack(anchor='w', padx=22, pady=(4, 2))
        wrap = tk.Frame(p, bg=C['bg']); wrap.pack(fill='x', padx=20, pady=4)
        cols = ('OID', '주문번호', '고객사', '생산품명', '수량', '진행률', '검사상태')
        ws   = (0, 135, 150, 100, 150, 70, 80, 90)
        tree = make_tree(wrap, cols, ws, height=8)
        tree.column('OID', width=0, stretch=False)

        def _load():
            rows = self.db.query("""
                SELECT o.id, o.order_no, c.name, i.name, o.quantity,
                       COALESCE((SELECT SUM(done_qty) FROM work_orders WHERE order_id=o.id),0)||'/'||o.quantity,
                       CASE WHEN EXISTS(SELECT 1 FROM inspections WHERE order_id=o.id AND result='합격')
                            THEN '합격'
                            WHEN EXISTS(SELECT 1 FROM inspections WHERE order_id=o.id AND result='불합격')
                            THEN '불합격'
                            ELSE '미검사' END
                FROM orders o
                LEFT JOIN customers c ON o.customer_id=c.id
                LEFT JOIN items i ON o.item_id=i.id
                WHERE o.status IN ('진행중','출하완료')
                ORDER BY o.due_date DESC
            """)
            def tag(i, r):
                if r[6] == '합격': return 'pass_tag'
                if r[6] == '불합격': return 'fail_tag'
                return 'even' if i%2 else ''
            fill_tree(tree, rows, tag)

        # ── 2. 검사 이력 트리 ──
        make_label(p, " 2. 검사 이력  (행 클릭 → 수정/삭제 모드)",
                   bold=True, size=11, bg=C['bg']).pack(anchor='w', padx=22, pady=(8, 2))
        iwrap = tk.Frame(p, bg=C['bg']); iwrap.pack(fill='x', padx=20, pady=4)
        icols = ('IID', '검사일', '구분', '결과', '샘플', '불량', '사유', '비고')
        iws   = (0, 110, 80, 70, 60, 60, 200, 200)
        itree = make_tree(iwrap, icols, iws, height=5)
        itree.column('IID', width=0, stretch=False)

        f = tk.Frame(p, bg='white', padx=22, pady=14); f.pack(fill='x', padx=20, pady=8)
        sel_var = tk.StringVar(value="주문을 선택하세요")
        sel_id  = [None]
        edit_iid = [None]

        def _load_inspections():
            for x in itree.get_children(): itree.delete(x)
            if not sel_id[0]: return
            rows = self.db.query("""
                SELECT id, inspect_date, inspect_type, result,
                       COALESCE(sample_qty,0), COALESCE(defect_qty,0),
                       COALESCE(defect_reason,''), COALESCE(memo,'')
                FROM inspections
                WHERE order_id=?
                ORDER BY inspect_date DESC, id DESC
            """, (sel_id[0],))
            def tag(i, r):
                if r[3] == '합격': return 'pass_tag'
                if r[3] == '불합격': return 'fail_tag'
                return 'even' if i%2 else ''
            fill_tree(itree, rows, tag)

        def _clear_form():
            edit_iid[0] = None
            type_v.set('최종검사')
            dt_v.set(datetime.now().strftime("%Y-%m-%d"))
            s_v.set(''); d_v.set('0'); r_v.set(''); m_v.set('')
            res_v.set('합격')
            mode_lbl.config(text="신규 검사 등록 중", fg='#2E7D32', bg='#E8F5E9')
            itree.selection_remove(itree.selection())

        def _on_sel(e):
            s = tree.selection()
            if not s: return
            v = tree.item(s[0])['values']
            sel_id[0] = v[0]
            sel_var.set(f"  {v[1]}  ({v[3]} / 수량 {v[4]})")
            _clear_form()
            _load_inspections()
        tree.bind('<<TreeviewSelect>>', _on_sel)

        def _on_insp_select(e):
            s = itree.selection()
            if not s: return
            v = itree.item(s[0])['values']
            edit_iid[0] = v[0]
            dt_v.set(str(v[1] or ''))
            type_v.set(str(v[2] or '최종검사'))
            res_v.set(str(v[3] or '합격'))
            s_v.set(str(v[4] or '0'))
            d_v.set(str(v[5] or '0'))
            r_v.set(str(v[6] or ''))
            m_v.set(str(v[7] or ''))
            mode_lbl.config(text=f"검사 [{v[1]}] 수정 중", fg='#E65100', bg='#FFF3E0')
        itree.bind('<<TreeviewSelect>>', _on_insp_select)

        make_label(f, "검사 결과 등록", bold=True, size=11, color=C['primary'], bg='white').grid(row=0, column=0, columnspan=6, sticky='w')
        mode_lbl = tk.Label(f, text="신규 검사 등록 중",
                            font=('Malgun Gothic', 10, 'bold'),
                            bg='#E8F5E9', fg='#2E7D32', padx=10, pady=3)
        mode_lbl.grid(row=0, column=6, columnspan=2, sticky='e', padx=8)

        make_label(f, "선택:", size=9, color=C['secondary'], bg='white').grid(row=1, column=0, sticky='w', pady=4)
        tk.Label(f, textvariable=sel_var, font=('Malgun Gothic', 10, 'bold'),
                 fg=C['primary'], bg='white').grid(row=1, column=1, columnspan=5, sticky='w')

        def _lbl(r, c, t): make_label(f, t, size=9, color=C['secondary'], bg='white').grid(row=r, column=c, sticky='w', padx=6, pady=3)

        _lbl(2, 0, "검사구분")
        type_v = tk.StringVar(value='최종검사')
        type_e = make_combo(f, type_v, ['자주검사', '최종검사'], width=12); type_e.grid(row=2, column=1, padx=4)
        _lbl(2, 2, "검사일 *");  dt_v = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        dt_e = make_entry(f, dt_v, 14); dt_e.grid(row=2, column=3, padx=4)
        _lbl(2, 4, "샘플수");    s_v = tk.StringVar()
        s_e = make_entry(f, s_v, 8); s_e.grid(row=2, column=5, padx=4)

        _lbl(3, 0, "결과 *")
        res_v = tk.StringVar(value='합격')
        rf = tk.Frame(f, bg='white'); rf.grid(row=3, column=1, sticky='w', pady=4)
        tk.Radiobutton(rf, text="합격", variable=res_v, value='합격',
                       bg='white', fg=C['success'], font=('Malgun Gothic', 11, 'bold')).pack(side='left')
        tk.Radiobutton(rf, text="불합격", variable=res_v, value='불합격',
                       bg='white', fg=C['danger'], font=('Malgun Gothic', 11, 'bold')).pack(side='left', padx=12)

        _lbl(3, 2, "불량수");      d_v = tk.StringVar(value='0')
        d_e = make_entry(f, d_v, 8); d_e.grid(row=3, column=3, padx=4)
        _lbl(4, 0, "불량사유");    r_v = tk.StringVar()
        r_e = make_entry(f, r_v, 40); r_e.grid(row=4, column=1, columnspan=4, padx=4, sticky='w', pady=4)
        _lbl(5, 0, "비고");        m_v = tk.StringVar()
        m_e = make_entry(f, m_v, 40); m_e.grid(row=5, column=1, columnspan=4, padx=4, sticky='w')

        def _read(widget, var, default=''):
            try:
                v = widget.get()
                if v: return v.strip()
            except Exception: pass
            return (var.get() or default).strip()

        def _gather():
            return {
                'type': _read(type_e, type_v, '최종검사'),
                'date': _read(dt_e, dt_v),
                'samp': _read(s_e, s_v, '0'),
                'def':  _read(d_e, d_v, '0'),
                'rsn':  _read(r_e, r_v),
                'memo': _read(m_e, m_v),
                'res':  res_v.get() or '합격',
            }

        def _save_new():
            if not sel_id[0]:
                messagebox.showerror("오류", "주문을 선택하세요."); return
            if edit_iid[0] is not None:
                if not messagebox.askyesno("확인", "현재 수정 모드입니다.\n신규 등록으로 전환할까요?"):
                    return
                _clear_form(); return
            g = _gather()
            try:
                samp = int(g['samp'] or 0); dqty = int(g['def'] or 0)
            except:
                messagebox.showerror("오류", f"샘플수/불량수는 정수로 입력하세요."); return
            self.db.execute("""
                INSERT INTO inspections(order_id,inspect_type,inspect_date,sample_qty,result,defect_qty,defect_reason,inspector_id,memo)
                VALUES(?,?,?,?,?,?,?,?,?)
            """, (sel_id[0], g['type'], g['date'], samp, g['res'], dqty, g['rsn'], self.user['id'], g['memo']))
            messagebox.showinfo("완료", f"검사 결과 [{g['res']}] 등록!")
            _clear_form(); _load(); _load_inspections()

        def _update():
            if edit_iid[0] is None:
                messagebox.showwarning("수정", "수정할 검사를 이력에서 선택하세요."); return
            g = _gather()
            try:
                samp = int(g['samp'] or 0); dqty = int(g['def'] or 0)
            except:
                messagebox.showerror("오류", "샘플수/불량수는 정수로 입력하세요."); return
            if not messagebox.askyesno("확인",
                f"검사를 수정하시겠습니까?\n검사일: {g['date']} / 결과: {g['res']}"):
                return
            self.db.execute("""
                UPDATE inspections
                SET inspect_type=?, inspect_date=?, sample_qty=?, result=?,
                    defect_qty=?, defect_reason=?, memo=?
                WHERE id=?
            """, (g['type'], g['date'], samp, g['res'], dqty, g['rsn'], g['memo'], edit_iid[0]))
            messagebox.showinfo("완료", "검사 수정 완료!")
            _clear_form(); _load(); _load_inspections()

        def _delete():
            if edit_iid[0] is None:
                messagebox.showwarning("삭제", "삭제할 검사를 이력에서 선택하세요."); return
            if not messagebox.askyesno("삭제 확인",
                "이 검사 결과를 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다."):
                return
            self.db.execute("DELETE FROM inspections WHERE id=?", (edit_iid[0],))
            messagebox.showinfo("삭제 완료", "검사가 삭제되었습니다.")
            _clear_form(); _load(); _load_inspections()

        # 3개 액션 버튼
        bf = tk.Frame(f, bg='white')
        bf.grid(row=6, column=0, columnspan=8, pady=(12, 4), sticky='e')
        color_btn(bf, "검사 결과 저장", _save_new, theme='save').pack(side='right', padx=5)
        color_btn(bf, "검사 수정", _update, theme='update').pack(side='right', padx=5)
        color_btn(bf, "검사 삭제", _delete, theme='delete').pack(side='right', padx=5)
        _reset_bar = tk.Frame(p, bg=C['bg']); _reset_bar.pack(side='bottom', fill='x', padx=20, pady=8)
        color_btn(_reset_bar, "초기화", _clear_form, theme='neutral', size=9, padx=10, pady=5).pack(side='right')
        self.root.bind('<Escape>', lambda e: _clear_form())

        _load()

    # ========================================================
    # 출하 관리
    # ========================================================
    def _pg_shipment(self):
        p = self.page_area
        page_header(p, "출하 관리", "  완성품 납품 — 등록 / 수정 / 삭제")

        # ── 1. 출하 가능 주문 (검사 합격) ──
        make_label(p, " 1. 출하 가능 주문  (검사 합격)",
                   bold=True, size=11, bg=C['bg']).pack(anchor='w', padx=22, pady=(4, 2))
        wrap = tk.Frame(p, bg=C['bg']); wrap.pack(fill='x', padx=20, pady=4)
        cols = ('OID', '주문번호', '고객사', '생산품명', '수량', '납기일', '검사결과')
        ws   = (0, 135, 150, 100, 150, 70, 100, 90)
        tree = make_tree(wrap, cols, ws, height=6)
        tree.column('OID', width=0, stretch=False)

        def _load_orders():
            rows = self.db.query("""
                SELECT o.id, o.order_no, c.name, i.name, o.quantity, o.due_date,
                       COALESCE((SELECT result FROM inspections WHERE order_id=o.id ORDER BY id DESC LIMIT 1),'-')
                FROM orders o
                LEFT JOIN customers c ON o.customer_id=c.id
                LEFT JOIN items i ON o.item_id=i.id
                WHERE o.status IN ('진행중','출하완료')
                  AND EXISTS(SELECT 1 FROM inspections WHERE order_id=o.id AND result='합격')
                ORDER BY o.due_date DESC
            """)
            fill_tree(tree, rows, lambda i, r: 'pass_tag')

        # ── 2. 전체 출하 이력 ──
        make_label(p, " 2. 출하 이력  (행 클릭 → 수정/삭제 모드)",
                   bold=True, size=11, bg=C['bg']).pack(anchor='w', padx=22, pady=(8, 2))
        hwrap = tk.Frame(p, bg=C['bg']); hwrap.pack(fill='x', padx=20, pady=4)
        hcols = ('SID', '출하번호', '주문번호', '고객사', '생산품명', '출하수량', '출하일', '담당', '비고')
        hws   = (0, 130, 135, 130, 100, 80, 100, 90, 200)
        htree = make_tree(hwrap, hcols, hws, height=6)
        htree.column('SID', width=0, stretch=False)

        def _load_ships():
            rows = self.db.query("""
                SELECT sh.id, sh.ship_no, o.order_no, c.name, i.name,
                       sh.quantity, sh.ship_date,
                       COALESCE(u.name,'-'), COALESCE(sh.memo,'')
                FROM shipments sh
                LEFT JOIN orders o ON sh.order_id=o.id
                LEFT JOIN customers c ON o.customer_id=c.id
                LEFT JOIN items i ON o.item_id=i.id
                LEFT JOIN users u ON sh.shipped_by=u.id
                ORDER BY sh.ship_date DESC, sh.id DESC LIMIT 50
            """)
            fill_tree(htree, rows, lambda i, r: 'even' if i%2 else '')

        f = tk.Frame(p, bg='white', padx=22, pady=14); f.pack(fill='x', padx=20, pady=8)
        sel_var = tk.StringVar(value="출하할 주문을 선택하세요")
        sel_oid = [None]   # 신규 등록 대상 주문 id
        edit_sid = [None]  # 수정/삭제 대상 출하 id
        edit_oid = [None]  # 수정 시 원래 주문 id

        def _clear_form():
            edit_sid[0] = None
            edit_oid[0] = None
            dt_v.set(datetime.now().strftime("%Y-%m-%d"))
            q_v.set(''); m_v.set('')
            mode_lbl.config(text="신규 출하 등록 중", fg='#2E7D32', bg='#E8F5E9')
            htree.selection_remove(htree.selection())

        def _on_order_sel(e):
            s = tree.selection()
            if not s: return
            v = tree.item(s[0])['values']
            sel_oid[0] = v[0]
            sel_var.set(f"  {v[1]}  ({v[2]} / {v[3]} / {v[4]}개)")
            _clear_form()
            q_v.set(str(v[4] or ''))  # 출하수량 기본값 = 주문수량
        tree.bind('<<TreeviewSelect>>', _on_order_sel)

        def _on_ship_sel(e):
            s = htree.selection()
            if not s: return
            v = htree.item(s[0])['values']
            edit_sid[0] = v[0]
            # 출하번호로 원래 주문 id 조회
            row = self.db.query("SELECT order_id FROM shipments WHERE id=?", (v[0],))
            edit_oid[0] = row[0][0] if row else None
            sel_var.set(f"  [{v[1]}]  {v[2]} / {v[3]} / {v[4]}")
            dt_v.set(str(v[6] or ''))
            q_v.set(str(v[5] or ''))
            m_v.set(str(v[8] or ''))
            mode_lbl.config(text=f"출하 [{v[1]}] 수정 중", fg='#E65100', bg='#FFF3E0')
        htree.bind('<<TreeviewSelect>>', _on_ship_sel)

        make_label(f, "출하 등록", bold=True, size=11, color=C['primary'], bg='white').grid(row=0, column=0, columnspan=6, sticky='w')
        mode_lbl = tk.Label(f, text="신규 출하 등록 중",
                            font=('Malgun Gothic', 10, 'bold'),
                            bg='#E8F5E9', fg='#2E7D32', padx=10, pady=3)
        mode_lbl.grid(row=0, column=6, columnspan=2, sticky='e', padx=8)

        make_label(f, "선택:", size=9, color=C['secondary'], bg='white').grid(row=1, column=0, sticky='w', pady=4)
        tk.Label(f, textvariable=sel_var, font=('Malgun Gothic', 10, 'bold'),
                 fg=C['primary'], bg='white').grid(row=1, column=1, columnspan=5, sticky='w')

        def _lbl(r, c, t): make_label(f, t, size=9, color=C['secondary'], bg='white').grid(row=r, column=c, sticky='w', padx=6, pady=3)

        _lbl(2, 0, "출하일 *"); dt_v = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        dt_e = make_entry(f, dt_v, 14); dt_e.grid(row=2, column=1, padx=4)
        _lbl(2, 2, "출하수량 *"); q_v = tk.StringVar()
        q_e = make_entry(f, q_v, 10); q_e.grid(row=2, column=3, padx=4)
        _lbl(3, 0, "비고"); m_v = tk.StringVar()
        m_e = make_entry(f, m_v, 40); m_e.grid(row=3, column=1, columnspan=4, padx=4, sticky='w', pady=4)

        def _read(widget, var, default=''):
            try:
                v = widget.get()
                if v: return v.strip()
            except Exception: pass
            return (var.get() or default).strip()

        def _gather():
            return _read(dt_e, dt_v), _read(q_e, q_v), _read(m_e, m_v)

        def _save_new():
            if edit_sid[0] is not None:
                if not messagebox.askyesno("확인", "현재 수정 모드입니다.\n신규 등록으로 전환할까요?"):
                    return
                _clear_form(); return
            if not sel_oid[0]:
                messagebox.showerror("오류", "주문을 선택하세요."); return
            dt_s, q_s, m_s = _gather()
            if not q_s:
                messagebox.showerror("입력 오류", "출하수량을 입력하세요."); return
            try: q = int(q_s)
            except: messagebox.showerror("오류", f"수량을 정수로 입력하세요. (현재: '{q_s}')"); return
            ship_no = self.db.next_ship_no()
            self.db.execute("""
                INSERT INTO shipments(order_id,ship_no,ship_date,quantity,shipped_by,memo)
                VALUES(?,?,?,?,?,?)
            """, (sel_oid[0], ship_no, dt_s, q, self.user['id'], m_s))
            self.db.execute("UPDATE orders SET status='출하완료' WHERE id=?", (sel_oid[0],))
            messagebox.showinfo("완료", f"출하번호 [{ship_no}] 등록 완료!")
            sel_oid[0] = None; sel_var.set("출하할 주문을 선택하세요")
            _clear_form(); _load_orders(); _load_ships()

        def _update():
            if edit_sid[0] is None:
                messagebox.showwarning("수정", "수정할 출하를 이력에서 선택하세요."); return
            dt_s, q_s, m_s = _gather()
            if not q_s:
                messagebox.showerror("입력 오류", "출하수량을 입력하세요."); return
            try: q = int(q_s)
            except: messagebox.showerror("오류", "수량을 정수로 입력하세요."); return
            if not messagebox.askyesno("확인",
                f"출하 정보를 수정하시겠습니까?\n출하일: {dt_s} / 수량: {q}"):
                return
            self.db.execute("""
                UPDATE shipments SET ship_date=?, quantity=?, memo=?
                WHERE id=?
            """, (dt_s, q, m_s, edit_sid[0]))
            messagebox.showinfo("완료", "출하 수정 완료!")
            _clear_form(); _load_orders(); _load_ships()

        def _delete():
            if edit_sid[0] is None:
                messagebox.showwarning("삭제", "삭제할 출하를 이력에서 선택하세요."); return
            if not messagebox.askyesno("삭제 확인",
                "이 출하 기록을 삭제하시겠습니까?\n주문 상태는 '진행중'으로 되돌아갑니다.\n이 작업은 되돌릴 수 없습니다."):
                return
            oid = edit_oid[0]
            self.db.execute("DELETE FROM shipments WHERE id=?", (edit_sid[0],))
            # 해당 주문에 다른 출하가 없으면 진행중으로 되돌리기
            if oid:
                cnt = self.db.query("SELECT COUNT(*) FROM shipments WHERE order_id=?", (oid,))[0][0]
                if cnt == 0:
                    self.db.execute("UPDATE orders SET status='진행중' WHERE id=?", (oid,))
            messagebox.showinfo("삭제 완료", "출하가 삭제되었습니다.")
            _clear_form(); _load_orders(); _load_ships()

        # 3개 액션 버튼
        bf = tk.Frame(f, bg='white')
        bf.grid(row=4, column=0, columnspan=8, pady=(12, 4), sticky='e')
        color_btn(bf, "출하 등록", _save_new, theme='save').pack(side='right', padx=5)
        color_btn(bf, "출하 수정", _update, theme='update').pack(side='right', padx=5)
        color_btn(bf, "출하 삭제", _delete, theme='delete').pack(side='right', padx=5)
        _reset_bar = tk.Frame(p, bg=C['bg']); _reset_bar.pack(side='bottom', fill='x', padx=20, pady=8)
        color_btn(_reset_bar, "초기화", _clear_form, theme='neutral', size=9, padx=10, pady=5).pack(side='right')
        self.root.bind('<Escape>', lambda e: _clear_form())

        _load_orders(); _load_ships()

    # ========================================================
    # 보고서/출력
    # ========================================================
    def _pg_report(self):
        p = self.page_area
        page_header(p, "보고서 / 출력", "  생산 / 품질 / 출하 보고서")

        btn_row = tk.Frame(p, bg=C['bg']); btn_row.pack(fill='x', padx=24, pady=18)

        def _card(title, cmd):
            card = tk.Frame(btn_row, bg='white', padx=16, pady=20)
            card.pack(side='left', expand=True, fill='both', padx=8)
            make_label(card, title, bold=True, size=12, color=C['primary'], bg='white').pack(pady=4)
            color_btn(card, "미리보기 / 인쇄", cmd, theme='export',
                      size=11, padx=18, pady=8).pack(pady=8)

        _card("작업지시서",   self._rpt_workorder)
        _card("생산일보",     self._rpt_daily)
        _card("거래명세서",   self._rpt_invoice)
        _card("설비별 가동", self._rpt_equipment)
        _card("📊 그래프 분석", self._rpt_graph)
        _card("💾 DB → PDF 백업", self._rpt_db_backup)

        tk.Frame(p, bg=C['border'], height=1).pack(fill='x', padx=20, pady=4)
        make_label(p, " 미리보기", bold=True, size=10, bg=C['bg']).pack(anchor='w', padx=22)

        pf = tk.Frame(p, bg='white'); pf.pack(fill='both', expand=True, padx=20, pady=8)
        self._preview = tk.Text(pf, font=('Courier New', 9), bg='#FAFAFA',
                                 relief='flat', wrap='none', state='disabled')
        sy = ttk.Scrollbar(pf, command=self._preview.yview)
        sx = ttk.Scrollbar(pf, orient='horizontal', command=self._preview.xview)
        self._preview.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side='right', fill='y'); sx.pack(side='bottom', fill='x')
        self._preview.pack(fill='both', expand=True)

    def _set_preview(self, text):
        self._preview.config(state='normal')
        self._preview.delete('1.0', 'end')
        self._preview.insert('1.0', text)
        self._preview.config(state='disabled')

    def _rpt_workorder(self):
        rows = self.db.query("""
            SELECT wo.wo_no, o.order_no, c.name, i.name, i.material,
                   wo.process, COALESCE(e.code,'-'), COALESCE(u.name,'-'),
                   wo.plan_qty, wo.done_qty, wo.status, o.due_date
            FROM work_orders wo
            LEFT JOIN orders o ON wo.order_id=o.id
            LEFT JOIN customers c ON o.customer_id=c.id
            LEFT JOIN items i ON o.item_id=i.id
            LEFT JOIN equipments e ON wo.equipment_id=e.id
            LEFT JOIN users u ON wo.worker_id=u.id
            ORDER BY wo.created_at DESC LIMIT 30
        """)
        W = 110
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        parts = [
            "=" * W,
            f"{COMPANY:^{W}}",
            f"{'작 업 지 시 서':^{W}}",
            f"{'출력일시: ' + now:>{W}}",
            "=" * W,
        ]
        for r in rows:
            parts += [
                f"  작업번호: {r[0]:<20}  주문번호: {r[1]:<20}  납기일: {r[11]}",
                f"  고  객: {r[2]:<20}",
                f"  생산품명: {r[3]:<28}  재질: {r[4]}",
                f"  공  정: {r[5]:<10}  설비: {r[6]:<14}  담당: {r[7]}",
                f"  계획수량: {r[8]:>6}    실적: {r[9]:>6}    상태: {r[10]}",
                "-" * W,
            ]
        content = '\n'.join(parts) if rows else "작업지시 없음"
        self._set_preview(content)
        open_print_preview(content, "작업지시서")

    def _rpt_daily(self):
        today = datetime.now().strftime("%Y-%m-%d")
        rows = self.db.query("""
            SELECT pr.work_date, wo.wo_no, i.name, wo.process,
                   COALESCE(e.code,'-'), pr.qty, pr.defect_qty, COALESCE(u.name,'-')
            FROM production_records pr
            LEFT JOIN work_orders wo ON pr.wo_id=wo.id
            LEFT JOIN orders o ON wo.order_id=o.id
            LEFT JOIN items i ON o.item_id=i.id
            LEFT JOIN equipments e ON pr.equipment_id=e.id
            LEFT JOIN users u ON pr.worker_id=u.id
            ORDER BY pr.work_date DESC, pr.created_at DESC
            LIMIT 50
        """)
        W = 110
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        lines = [
            "=" * W,
            f"{COMPANY:^{W}}",
            f"{'생 산 일 보':^{W}}",
            f"{'출력일시: ' + now:>{W}}",
            "=" * W,
            f"{'작업일':<12} {'작업번호':<16} {'생산품명':<22} {'공정':<8} {'설비':<10} {'양품':>6} {'불량':>5} {'담당':<8}",
            "-" * W,
        ]
        total_q = total_d = 0
        for r in rows:
            lines.append(f"{r[0]:<12} {r[1]:<16} {r[2]:<22} {r[3]:<8} {r[4]:<10} {r[5]:>6} {r[6]:>5} {r[7]:<8}")
            total_q += r[5]; total_d += r[6]
        lines += ["-" * W, f"  합계  양품 {total_q}    불량 {total_d}    불량률 {(total_d*100/(total_q+total_d)) if (total_q+total_d) else 0:.2f}%", "=" * W]
        content = '\n'.join(lines)
        self._set_preview(content)
        open_print_preview(content, "생산일보")

    def _rpt_invoice(self):
        try:
            rows = self.db.query("""
                SELECT COALESCE(sh.ship_no,''), COALESCE(sh.ship_date,''),
                       COALESCE(c.name,'-'), COALESCE(i.name,'-'),
                       COALESCE(sh.quantity,0), COALESCE(o.order_no,'-'),
                       COALESCE(sh.memo,'')
                FROM shipments sh
                LEFT JOIN orders o ON sh.order_id=o.id
                LEFT JOIN customers c ON o.customer_id=c.id
                LEFT JOIN items i ON o.item_id=i.id
                ORDER BY sh.ship_date DESC, sh.id DESC LIMIT 20
            """)
        except Exception as e:
            messagebox.showerror("쿼리 오류", str(e)); return

        if not rows:
            content = "출하 기록이 없습니다.\n\n출하 관리 메뉴에서 먼저 출하를 등록하세요."
            self._set_preview(content)
            open_print_preview(content, "거래명세서")
            return

        W = 56
        parts = []
        for r in rows:
            parts.append('\n'.join([
                f"+{'=' * W}+",
                f"|{COMPANY:^{W}}|",
                f"|{'거 래 명 세 서':^{W}}|",
                f"+{'-' * W}+",
                f"| 출하번호: {str(r[0]):<{W-12}}|",
                f"| 출하일자: {str(r[1]):<{W-12}}|",
                f"| 거 래 처: {str(r[2]):<{W-12}}|",
                f"| 주문번호: {str(r[5]):<{W-12}}|",
                f"+{'-' * W}+",
                f"| 생산품명: {str(r[3]):<{W-12}}|",
                f"| 수    량: {str(r[4]):<{W-12}}|",
                f"| 비    고: {str(r[6])[:W-12]:<{W-12}}|",
                f"+{'-' * W}+",
                f"| 인수자 확인: ___________________ (인){' ':<{W-40}}|",
                f"+{'=' * W}+",
            ]))
        content = '\n\n'.join(parts)
        self._set_preview(content)
        open_print_preview(content, "거래명세서")

    def _rpt_equipment(self):
        rows = self.db.query("""
            SELECT e.code, e.name, e.process, e.spec, e.status,
                   (SELECT COUNT(*) FROM work_orders WHERE equipment_id=e.id),
                   (SELECT COUNT(*) FROM work_orders WHERE equipment_id=e.id AND status='완료'),
                   (SELECT SUM(qty) FROM production_records WHERE equipment_id=e.id),
                   (SELECT SUM(defect_qty) FROM production_records WHERE equipment_id=e.id)
            FROM equipments e
            WHERE e.active=1
        """)
        W = 110
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        lines = [
            "=" * W,
            f"{COMPANY:^{W}}",
            f"{'설 비 별  가 동  현 황':^{W}}",
            f"{'출력일시: ' + now:>{W}}",
            "=" * W,
            f"{'코드':<10} {'설비명':<24} {'공정':<10} {'생산품목':<28} {'상태':<8} {'배정':>5} {'완료':>5} {'생산량':>8} {'불량':>5}",
            "-" * W,
        ]
        for r in rows:
            lines.append(f"{r[0]:<10} {r[1]:<24} {r[2]:<10} {r[3]:<28} {r[4]:<8} {r[5] or 0:>5} {r[6] or 0:>5} {r[7] or 0:>8} {r[8] or 0:>5}")
        lines.append("=" * W)
        content = '\n'.join(lines)
        self._set_preview(content)
        open_print_preview(content, "설비별 가동현황")

    # ========================================================
    # 그래프 분석
    # ========================================================
    def _rpt_db_backup(self):
        """전체 DB를 PDF로 백업 (EXE/스크립트와 같은 폴더에 저장)."""
        import subprocess, tempfile, html as _html, platform
        from datetime import datetime as _dt

        # 저장 위치: EXE 옆 (frozen) 또는 production.db 옆
        if getattr(sys, 'frozen', False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))

        ts = _dt.now().strftime('%Y%m%d_%H%M')
        pdf_path = os.path.join(base, f"DB_백업_{ts}.pdf")

        # 한글 라벨 매핑
        LABELS = {
            'orders': '📋 수주', 'work_orders': '🔧 작업지시',
            'production_records': '⚙️ 생산실적', 'inspections': '🔍 품질검사',
            'shipments': '🚚 출하', 'customers': '🏢 고객사',
            'items': '📦 품목', 'equipments': '🏭 설비', 'users': '👤 사용자',
        }
        ORDER = ['orders', 'work_orders', 'production_records', 'inspections',
                 'shipments', 'customers', 'items', 'equipments', 'users']

        def _safe(v):
            if v is None: return ''
            s = str(v)
            if len(s) > 80: s = s[:78] + '..'
            return _html.escape(s)

        try:
            total = sum(self.db.query(f"SELECT COUNT(*) FROM {t}")[0][0] for t in ORDER)
        except Exception as e:
            messagebox.showerror("DB 백업", f"DB 읽기 실패:\n{e}"); return

        now_str = _dt.now().strftime('%Y-%m-%d %H:%M:%S')
        CSS = """
@page { size: A4; margin: 14mm 12mm; }
body { font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif;
       font-size: 9.5pt; line-height: 1.4; color: #222; }
.cover { text-align:center; margin: 80px 0 60px; }
.cover h1 { color:#00695C; font-size: 28pt; margin: 0; }
.cover h2 { color:#37474F; font-size: 16pt; margin: 12px 0; font-weight: normal; }
.cover .meta { color:#78909C; font-size: 10pt; margin-top: 20px; line-height: 1.8; }
.section { page-break-before: always; margin-bottom: 18px; }
.section:first-of-type { page-break-before: auto; }
h2.tname { color:#00695C; border-bottom: 3px solid #00695C; padding-bottom: 6px;
           font-size: 16pt; margin: 16px 0 6px; }
.cnt { color:#FF6F00; font-weight: bold; font-size: 11pt; margin-left: 8px; }
table { border-collapse: collapse; width: 100%; margin: 8px 0;
        font-size: 8pt; page-break-inside: avoid; }
th { background:#00695C; color:white; padding: 5px 6px; text-align:center; font-size: 8.5pt; }
td { border: 1px solid #CFD8DC; padding: 4px 6px; vertical-align: top; }
tbody tr:nth-child(even) { background:#F5F7F8; }
.no-data { color:#999; font-style: italic; padding: 8px; }
"""
        parts = [f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<title>SEO JIN PRECISION DB 백업</title>
<style>{CSS}</style></head><body>
<div class="cover">
  <h1>SEO JIN PRECISION CO.</h1>
  <h2>생산관리 시스템 — DB 전체 백업</h2>
  <div class="meta">
    파일: production.db<br>생성일시: {now_str}<br>
    총 {total:,} 행 / {len(ORDER)} 테이블
  </div>
</div>"""]

        c = self.db.conn.cursor()
        for tname in ORDER:
            label = LABELS[tname]
            cols = [r[1] for r in c.execute(f"PRAGMA table_info({tname})")]
            rows = self.db.query(f"SELECT * FROM {tname} ORDER BY rowid")
            parts.append(f'<div class="section"><h2 class="tname">{label}<span class="cnt">{len(rows):,} 행</span></h2>')
            if not rows:
                parts.append('<div class="no-data">데이터 없음</div></div>')
                continue
            parts.append('<table><thead><tr>')
            for col in cols: parts.append(f'<th>{_safe(col)}</th>')
            parts.append('</tr></thead><tbody>')
            for r in rows:
                parts.append('<tr>')
                for i, col in enumerate(cols):
                    parts.append(f'<td>{_safe(r[i])}</td>')
                parts.append('</tr>')
            parts.append('</tbody></table></div>')
        parts.append('</body></html>')

        # HTML → PDF (Chrome 헤드리스)
        tmp = tempfile.NamedTemporaryFile(suffix='.html', delete=False, mode='w', encoding='utf-8')
        tmp.write('\n'.join(parts)); tmp.close()

        chrome_paths = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
        chrome = next((p for p in chrome_paths if os.path.isfile(p)), None)

        if not chrome:
            # Chrome/Edge 없으면 HTML을 PDF 위치 옆에 저장하고 안내
            html_path = os.path.join(base, f"DB_백업_{ts}.html")
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(parts))
            messagebox.showinfo("DB 백업",
                f"Chrome/Edge가 없어 HTML로 저장했습니다.\n\n"
                f"{html_path}\n\n"
                f"브라우저로 열어 Ctrl+P로 PDF 저장 가능합니다.")
            try: webbrowser.open(f'file://{html_path}')
            except: pass
            return

        try:
            subprocess.run([chrome, "--headless", "--disable-gpu",
                            "--no-pdf-header-footer", "--no-margins",
                            f"--print-to-pdf={pdf_path}",
                            f"file://{tmp.name}"],
                           capture_output=True, timeout=60)
        except Exception as e:
            messagebox.showerror("DB 백업", f"PDF 생성 실패:\n{e}"); return

        if os.path.isfile(pdf_path):
            messagebox.showinfo("DB 백업 완료",
                f"✓ PDF로 저장되었습니다.\n\n"
                f"위치: {pdf_path}\n\n"
                f"총 {total:,} 행 / {len(ORDER)} 테이블")
            # 자동으로 열기
            try:
                if platform.system() == 'Windows':
                    os.startfile(pdf_path)
                elif platform.system() == 'Darwin':
                    subprocess.run(['open', pdf_path])
                else:
                    subprocess.run(['xdg-open', pdf_path])
            except: pass
        else:
            messagebox.showerror("DB 백업", "PDF 생성 실패. Chrome 실행 권한을 확인하세요.")

    def _rpt_graph(self):
        if not HAS_MPL:
            messagebox.showerror("matplotlib 미설치",
                "그래프 기능을 사용하려면 matplotlib을 설치해야 합니다.\n\n"
                "터미널에서 실행:\n  pip install matplotlib")
            return

        win = tk.Toplevel(self.root)
        win.title("그래프 분석")
        win.geometry("1200x780")
        win.configure(bg=C['bg'])

        # 상단 컨트롤
        top = tk.Frame(win, bg=C['header_bg'], height=56); top.pack(fill='x'); top.pack_propagate(False)
        tk.Label(top, text="📊  생산 / 품질 / 설비 그래프 분석",
                 font=('Malgun Gothic', 14, 'bold'),
                 fg='white', bg=C['header_bg']).pack(side='left', padx=20)

        kind_var = tk.StringVar(value='일별 생산량 추이')
        kinds = [
            '일별 생산량 추이',
            '공정별 생산 비중',
            '설비별 생산량 TOP 10',
            '일별 불량률 추이',
            '수주 상태 현황',
            '고객사별 수주 금액(수량)',
        ]
        tk.Label(top, text="유형:", font=('Malgun Gothic', 10),
                 fg='white', bg=C['header_bg']).pack(side='left', padx=(20, 6))
        cb = ttk.Combobox(top, textvariable=kind_var, values=kinds,
                          state='readonly', width=28, font=('Malgun Gothic', 10))
        cb.pack(side='left', padx=4)

        # 그래프 영역
        body = tk.Frame(win, bg='white'); body.pack(fill='both', expand=True, padx=10, pady=10)

        fig = Figure(figsize=(11, 6.5), dpi=100, facecolor='white')
        canvas = FigureCanvasTkAgg(fig, master=body)
        canvas.get_tk_widget().pack(fill='both', expand=True)

        toolbar_frame = tk.Frame(win, bg=C['bg']); toolbar_frame.pack(fill='x')
        NavigationToolbar2Tk(canvas, toolbar_frame).update()

        def _draw():
            fig.clear()
            # macOS Tk: textvariable가 미동기화될 수 있어 위젯에서 직접 읽기
            try:
                kind = (cb.get() or kind_var.get()).strip()
            except Exception:
                kind = kind_var.get()
            ax = fig.add_subplot(111)

            if kind == '일별 생산량 추이':
                rows = self.db.query("""
                    SELECT work_date, SUM(qty), SUM(defect_qty)
                    FROM production_records
                    WHERE work_date >= date('now','-30 day')
                    GROUP BY work_date ORDER BY work_date
                """)
                if not rows:
                    ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center',
                            fontsize=18, color='#999', transform=ax.transAxes)
                else:
                    dates = [r[0] for r in rows]
                    good = [r[1] or 0 for r in rows]
                    bad  = [r[2] or 0 for r in rows]
                    x = range(len(dates))
                    ax.bar(x, good, label='양품', color='#26A69A')
                    ax.bar(x, bad, bottom=good, label='불량', color='#EF5350')
                    ax.set_xticks(list(x))
                    ax.set_xticklabels(dates, rotation=45, ha='right', fontsize=8)
                    ax.set_ylabel('수량')
                    ax.set_title('최근 30일 일별 생산량', fontsize=14, fontweight='bold')
                    ax.legend(); ax.grid(axis='y', linestyle='--', alpha=0.5)

            elif kind == '공정별 생산 비중':
                rows = self.db.query("""
                    SELECT wo.process, SUM(pr.qty)
                    FROM production_records pr
                    LEFT JOIN work_orders wo ON pr.wo_id=wo.id
                    GROUP BY wo.process
                """)
                rows = [r for r in rows if r[0] and r[1]]
                if not rows:
                    ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center',
                            fontsize=18, color='#999', transform=ax.transAxes)
                else:
                    labels = [r[0] for r in rows]
                    sizes  = [r[1] for r in rows]
                    colors = ['#26A69A', '#FFA726', '#5C6BC0', '#EC407A'][:len(rows)]
                    ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                           startangle=90, textprops={'fontsize': 12})
                    ax.set_title('공정별 누적 생산 비중', fontsize=14, fontweight='bold')

            elif kind == '설비별 생산량 TOP 10':
                rows = self.db.query("""
                    SELECT e.code || ' ' || e.name, COALESCE(SUM(pr.qty),0)
                    FROM equipments e
                    LEFT JOIN production_records pr ON pr.equipment_id=e.id
                    GROUP BY e.id ORDER BY SUM(pr.qty) DESC NULLS LAST LIMIT 10
                """)
                rows = [r for r in rows if (r[1] or 0) > 0]
                if not rows:
                    ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center',
                            fontsize=18, color='#999', transform=ax.transAxes)
                else:
                    labels = [r[0] for r in rows][::-1]
                    vals   = [r[1] for r in rows][::-1]
                    bars = ax.barh(labels, vals, color='#00695C')
                    ax.set_xlabel('생산량')
                    ax.set_title('설비별 누적 생산량 TOP 10', fontsize=14, fontweight='bold')
                    ax.grid(axis='x', linestyle='--', alpha=0.5)
                    for b, v in zip(bars, vals):
                        ax.text(v, b.get_y() + b.get_height()/2,
                                f' {v}', va='center', fontsize=9)

            elif kind == '일별 불량률 추이':
                rows = self.db.query("""
                    SELECT work_date,
                           COALESCE(SUM(qty),0), COALESCE(SUM(defect_qty),0)
                    FROM production_records
                    WHERE work_date >= date('now','-30 day')
                    GROUP BY work_date ORDER BY work_date
                """)
                if not rows:
                    ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center',
                            fontsize=18, color='#999', transform=ax.transAxes)
                else:
                    dates = [r[0] for r in rows]
                    rates = [(r[2]*100/(r[1]+r[2])) if (r[1]+r[2]) else 0 for r in rows]
                    ax.plot(dates, rates, marker='o', linewidth=2.2,
                            color='#EF5350', markerfacecolor='#C62828')
                    ax.set_ylabel('불량률 (%)')
                    ax.set_title('최근 30일 일별 불량률', fontsize=14, fontweight='bold')
                    ax.tick_params(axis='x', rotation=45, labelsize=8)
                    ax.grid(linestyle='--', alpha=0.5)
                    ax.axhline(y=3, color='#FFA726', linestyle='--', label='관리 한계 3%')
                    ax.legend()

            elif kind == '수주 상태 현황':
                rows = self.db.query("""
                    SELECT status, COUNT(*) FROM orders GROUP BY status
                """)
                if not rows:
                    ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center',
                            fontsize=18, color='#999', transform=ax.transAxes)
                else:
                    labels = [r[0] for r in rows]
                    vals   = [r[1] for r in rows]
                    color_map = {'접수': '#42A5F5', '계획완료': '#66BB6A',
                                 '진행중': '#FFA726', '출하완료': '#26A69A',
                                 '취소': '#EF5350'}
                    colors = [color_map.get(l, '#90A4AE') for l in labels]
                    bars = ax.bar(labels, vals, color=colors)
                    ax.set_ylabel('건수')
                    ax.set_title('수주 상태별 건수', fontsize=14, fontweight='bold')
                    ax.grid(axis='y', linestyle='--', alpha=0.5)
                    for b, v in zip(bars, vals):
                        ax.text(b.get_x() + b.get_width()/2, v,
                                f'{v}', ha='center', va='bottom', fontsize=11, fontweight='bold')

            elif kind == '고객사별 수주 금액(수량)':
                rows = self.db.query("""
                    SELECT c.name, COALESCE(SUM(o.quantity),0)
                    FROM orders o LEFT JOIN customers c ON o.customer_id=c.id
                    GROUP BY c.id ORDER BY SUM(o.quantity) DESC LIMIT 10
                """)
                rows = [r for r in rows if r[0] and r[1]]
                if not rows:
                    ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center',
                            fontsize=18, color='#999', transform=ax.transAxes)
                else:
                    labels = [r[0] for r in rows]
                    vals   = [r[1] for r in rows]
                    bars = ax.bar(labels, vals, color='#5C6BC0')
                    ax.set_ylabel('총 수주량')
                    ax.set_title('고객사별 누적 수주량 TOP 10', fontsize=14, fontweight='bold')
                    ax.tick_params(axis='x', rotation=30, labelsize=9)
                    ax.grid(axis='y', linestyle='--', alpha=0.5)
                    for b, v in zip(bars, vals):
                        ax.text(b.get_x() + b.get_width()/2, v,
                                f'{v}', ha='center', va='bottom', fontsize=9)

            fig.tight_layout()
            canvas.draw()

        # 콤보 변경 → 즉시 다시 그리기 (이중 안전장치)
        cb.bind('<<ComboboxSelected>>', lambda e: _draw())
        kind_var.trace_add('write', lambda *a: _draw())
        tk.Button(top, text="🔄 새로고침", font=('Malgun Gothic', 10, 'bold'),
                  bg=C['accent'], fg='white', relief='flat',
                  cursor='hand2', padx=12, command=_draw).pack(side='left', padx=10)
        tk.Button(top, text="💾 PNG 저장", font=('Malgun Gothic', 10),
                  bg='#455A64', fg='white', relief='flat',
                  cursor='hand2', padx=12,
                  command=lambda: _save_png()).pack(side='left', padx=4)

        def _save_png():
            from tkinter import filedialog
            path = filedialog.asksaveasfilename(
                defaultextension='.png',
                filetypes=[('PNG 이미지', '*.png')],
                initialfile=f"graph_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
            if path:
                fig.savefig(path, dpi=150, bbox_inches='tight')
                messagebox.showinfo("저장 완료", f"저장됨:\n{path}")

        _draw()

    # ========================================================
    # 통계 (일별 / 월별 / 연간) — CSV 저장 가능
    # ========================================================
    def _pg_statistics(self):
        p = self.page_area
        page_header(p, "통계 (일별 / 월별 / 연간)", "  생산 / 품질 / 출하 통계 집계 및 저장")

        # ── 기간 탭 (일별 / 월별 / 연간) ──
        tab_bar = tk.Frame(p, bg=C['bg'])
        tab_bar.pack(fill='x', padx=20, pady=(8, 0))

        period_var = tk.StringVar(value='일별')
        tab_btns = {}

        def _switch_tab(name):
            period_var.set(name)
            now = datetime.now()
            if name == '일별':
                from_var.set((now - timedelta(days=30)).strftime('%Y-%m-%d'))
                to_var.set(now.strftime('%Y-%m-%d'))
            elif name == '월별':
                from_var.set((now - timedelta(days=365)).replace(day=1).strftime('%Y-%m-%d'))
                to_var.set(now.strftime('%Y-%m-%d'))
            else:  # 연간
                from_var.set(f"{now.year - 4}-01-01")
                to_var.set(f"{now.year}-12-31")
            for n, b in tab_btns.items():
                if n == name:
                    b.config(bg=C['primary'], fg='white', relief='flat')
                else:
                    b.config(bg='#ECEFF1', fg=C['primary'], relief='flat')
            _query()

        for name, color, emoji in [('일별', '#42A5F5', '📅'),
                                     ('월별', '#66BB6A', '📆'),
                                     ('연간', '#AB47BC', '🗓')]:
            b = tk.Button(tab_bar, text=f"  {emoji}  {name} 통계  ",
                          font=('Malgun Gothic', 14, 'bold'),
                          bg='#ECEFF1', fg=C['primary'], relief='flat',
                          padx=22, pady=12, cursor='hand2',
                          activebackground=color, activeforeground='white',
                          command=lambda n=name: _switch_tab(n))
            b.pack(side='left', padx=4)
            tab_btns[name] = b

        # 컨트롤 바 (기간 입력 + 구분 + 버튼)
        ctrl = tk.Frame(p, bg='white', padx=14, pady=12)
        ctrl.pack(fill='x', padx=20, pady=(6, 4))

        tk.Label(ctrl, text="시작:", font=('Malgun Gothic', 10),
                 bg='white').pack(side='left', padx=(4, 4))
        from_var = tk.StringVar(value=(datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d'))
        tk.Entry(ctrl, textvariable=from_var, width=12,
                 font=('Malgun Gothic', 11)).pack(side='left', padx=2)

        tk.Label(ctrl, text="종료:", font=('Malgun Gothic', 10),
                 bg='white').pack(side='left', padx=(10, 4))
        to_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(ctrl, textvariable=to_var, width=12,
                 font=('Malgun Gothic', 11)).pack(side='left', padx=2)

        kind_var = tk.StringVar(value='생산')
        tk.Label(ctrl, text="구분:", font=('Malgun Gothic', 10),
                 bg='white').pack(side='left', padx=(16, 4))
        ttk.Combobox(ctrl, textvariable=kind_var,
                     values=['생산', '품질', '출하', '설비별', '고객사별'],
                     state='readonly', width=10,
                     font=('Malgun Gothic', 11)).pack(side='left', padx=2)

        # 결과 요약 라벨
        sum_frame = tk.Frame(p, bg=C['bg'])
        sum_frame.pack(fill='x', padx=20, pady=(6, 4))
        sum_lbls = {}
        for k, color in [('총 행', '#455A64'), ('합계', '#00695C'),
                         ('불량합계', '#EF5350'), ('불량률', '#FF6F00')]:
            box = tk.Frame(sum_frame, bg='white', padx=14, pady=10)
            box.pack(side='left', expand=True, fill='x', padx=4)
            tk.Label(box, text=k, font=('Malgun Gothic', 10),
                     fg='#90A4AE', bg='white').pack()
            l = tk.Label(box, text='-', font=('Malgun Gothic', 16, 'bold'),
                         fg=color, bg='white')
            l.pack()
            sum_lbls[k] = l

        # 결과 테이블
        tf = tk.Frame(p, bg='white'); tf.pack(fill='both', expand=True, padx=20, pady=(4, 12))
        cols_set = {
            '생산':   ('기간', '작업건수', '양품', '불량', '불량률(%)'),
            '품질':   ('기간', '검사건수', '합격', '불합격', '불합격률(%)'),
            '출하':   ('기간', '출하건수', '출하수량', '거래처수', '-'),
            '설비별': ('기간', '설비코드', '설비명', '생산량', '불량'),
            '고객사별': ('기간', '고객사', '주문건수', '주문수량', '-'),
        }
        widths = (140, 120, 120, 120, 140)

        cols = cols_set['생산']
        tree = ttk.Treeview(tf, columns=cols, show='headings', height=14)
        for c, w in zip(cols, widths):
            tree.heading(c, text=c); tree.column(c, width=w, anchor='center')
        sy = ttk.Scrollbar(tf, orient='vertical', command=tree.yview)
        sx = ttk.Scrollbar(tf, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side='right', fill='y'); sx.pack(side='bottom', fill='x')
        tree.pack(fill='both', expand=True)

        current_data = {'rows': [], 'cols': cols, 'meta': {}}

        def _period_format(p):
            if period_var.get() == '일별':   return f"strftime('%Y-%m-%d', {p})"
            if period_var.get() == '월별':   return f"strftime('%Y-%m', {p})"
            return f"strftime('%Y', {p})"

        def _query():
            cols = cols_set[kind_var.get()]
            tree.config(columns=cols)
            for c, w in zip(cols, widths):
                tree.heading(c, text=c); tree.column(c, width=w, anchor='center')
            for r in tree.get_children(): tree.delete(r)

            f = from_var.get().strip(); t = to_var.get().strip()
            kind = kind_var.get()
            rows = []
            try:
                if kind == '생산':
                    pf = _period_format('pr.work_date')
                    rows = self.db.query(f"""
                        SELECT {pf} AS period,
                               COUNT(DISTINCT pr.wo_id),
                               COALESCE(SUM(pr.qty),0),
                               COALESCE(SUM(pr.defect_qty),0)
                        FROM production_records pr
                        WHERE pr.work_date BETWEEN ? AND ?
                        GROUP BY period ORDER BY period
                    """, (f, t))
                    data = []
                    tot_g = tot_d = 0
                    for r in rows:
                        g, d = r[2], r[3]; tot_g += g; tot_d += d
                        rate = (d * 100 / (g + d)) if (g + d) else 0
                        data.append((r[0], r[1], g, d, f"{rate:.2f}"))

                elif kind == '품질':
                    pf = _period_format('insp_date')
                    rows = self.db.query(f"""
                        SELECT {pf} AS period,
                               COUNT(*),
                               COALESCE(SUM(pass_qty),0),
                               COALESCE(SUM(fail_qty),0)
                        FROM inspections
                        WHERE insp_date BETWEEN ? AND ?
                        GROUP BY period ORDER BY period
                    """, (f, t))
                    data = []
                    tot_g = tot_d = 0
                    for r in rows:
                        ps, fl = r[2], r[3]; tot_g += ps; tot_d += fl
                        rate = (fl * 100 / (ps + fl)) if (ps + fl) else 0
                        data.append((r[0], r[1], ps, fl, f"{rate:.2f}"))

                elif kind == '출하':
                    pf = _period_format('sh.ship_date')
                    rows = self.db.query(f"""
                        SELECT {pf} AS period,
                               COUNT(*),
                               COALESCE(SUM(sh.quantity),0),
                               COUNT(DISTINCT o.customer_id)
                        FROM shipments sh
                        LEFT JOIN orders o ON sh.order_id=o.id
                        WHERE sh.ship_date BETWEEN ? AND ?
                        GROUP BY period ORDER BY period
                    """, (f, t))
                    data = [(r[0], r[1], r[2], r[3], '-') for r in rows]
                    tot_g = sum(r[2] for r in rows); tot_d = 0

                elif kind == '설비별':
                    pf = _period_format('pr.work_date')
                    rows = self.db.query(f"""
                        SELECT {pf} AS period, e.code, e.name,
                               COALESCE(SUM(pr.qty),0),
                               COALESCE(SUM(pr.defect_qty),0)
                        FROM production_records pr
                        LEFT JOIN equipments e ON pr.equipment_id=e.id
                        WHERE pr.work_date BETWEEN ? AND ?
                        GROUP BY period, e.id
                        ORDER BY period, SUM(pr.qty) DESC
                    """, (f, t))
                    data = [(r[0], r[1] or '-', r[2] or '-', r[3], r[4]) for r in rows]
                    tot_g = sum(r[3] for r in rows); tot_d = sum(r[4] for r in rows)

                else:  # 고객사별
                    pf = _period_format('o.order_date')
                    rows = self.db.query(f"""
                        SELECT {pf} AS period, c.name,
                               COUNT(*), COALESCE(SUM(o.quantity),0)
                        FROM orders o LEFT JOIN customers c ON o.customer_id=c.id
                        WHERE o.order_date BETWEEN ? AND ?
                        GROUP BY period, c.id
                        ORDER BY period, SUM(o.quantity) DESC
                    """, (f, t))
                    data = [(r[0], r[1] or '-', r[2], r[3], '-') for r in rows]
                    tot_g = sum(r[3] for r in rows); tot_d = 0
            except Exception as e:
                messagebox.showerror("조회 오류", str(e)); return

            for r in data:
                tree.insert('', 'end', values=r)

            current_data['rows'] = data
            current_data['cols'] = cols
            current_data['meta'] = {
                'period': period_var.get(), 'kind': kind,
                'from': f, 'to': t,
            }
            sum_lbls['총 행'].config(text=f"{len(data):,}")
            sum_lbls['합계'].config(text=f"{tot_g:,}")
            sum_lbls['불량합계'].config(text=f"{tot_d:,}")
            rate = (tot_d * 100 / (tot_g + tot_d)) if (tot_g + tot_d) else 0
            sum_lbls['불량률'].config(text=f"{rate:.2f} %")

        def _save_csv():
            if not current_data['rows']:
                messagebox.showwarning("저장", "조회된 데이터가 없습니다."); return
            from tkinter import filedialog
            import csv
            meta = current_data['meta']
            default_name = f"통계_{meta['kind']}_{meta['period']}_{meta['from']}_{meta['to']}.csv"
            path = filedialog.asksaveasfilename(
                defaultextension='.csv',
                filetypes=[('CSV (Excel)', '*.csv'), ('모든 파일', '*.*')],
                initialfile=default_name)
            if not path: return
            try:
                with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                    w = csv.writer(f)
                    w.writerow([f"# SEO JIN PRECISION CO. - 통계 보고서"])
                    w.writerow([f"# 구분: {meta['kind']}  /  기간: {meta['period']}"
                                f"  /  {meta['from']} ~ {meta['to']}"
                                f"  /  생성: {datetime.now().strftime('%Y-%m-%d %H:%M')}"])
                    w.writerow([])
                    w.writerow(current_data['cols'])
                    for r in current_data['rows']:
                        w.writerow(r)
                    w.writerow([])
                    w.writerow(['총 행', sum_lbls['총 행'].cget('text')])
                    w.writerow(['합계',  sum_lbls['합계'].cget('text')])
                    w.writerow(['불량합계', sum_lbls['불량합계'].cget('text')])
                    w.writerow(['불량률', sum_lbls['불량률'].cget('text')])
                messagebox.showinfo("저장 완료", f"CSV 파일로 저장되었습니다.\n\n{path}")
            except Exception as e:
                messagebox.showerror("저장 오류", str(e))

        def _save_print():
            if not current_data['rows']:
                messagebox.showwarning("출력", "조회된 데이터가 없습니다."); return
            meta = current_data['meta']
            W = 100
            lines = [
                "=" * W,
                f"{COMPANY:^{W}}",
                f"{'통 계 보 고 서 (' + meta['kind'] + ' / ' + meta['period'] + ')':^{W}}",
                f"기간: {meta['from']} ~ {meta['to']}    출력: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                "=" * W,
                "  ".join(f"{c:<14}" for c in current_data['cols']),
                "-" * W,
            ]
            for r in current_data['rows']:
                lines.append("  ".join(f"{str(v):<14}" for v in r))
            lines += [
                "-" * W,
                f"  총 행: {sum_lbls['총 행'].cget('text')}    "
                f"합계: {sum_lbls['합계'].cget('text')}    "
                f"불량: {sum_lbls['불량합계'].cget('text')}    "
                f"불량률: {sum_lbls['불량률'].cget('text')}",
                "=" * W,
            ]
            open_print_preview('\n'.join(lines), f"통계_{meta['kind']}_{meta['period']}")

        def _save_graph():
            if not HAS_MPL:
                messagebox.showerror("그래프 저장",
                    "matplotlib이 필요합니다.\npip install matplotlib"); return
            if not current_data['rows']:
                messagebox.showwarning("그래프", "조회된 데이터가 없습니다."); return
            from tkinter import filedialog
            meta = current_data['meta']
            data = current_data['rows']
            # 단순 막대그래프: 첫 컬럼 = x, 합계 위치 컬럼 = y
            kind = meta['kind']
            if kind in ('생산', '품질'):
                xs = [r[0] for r in data]; good = [r[2] for r in data]; bad = [r[3] for r in data]
            elif kind == '출하':
                xs = [r[0] for r in data]; good = [r[2] for r in data]; bad = [0]*len(data)
            elif kind == '설비별':
                xs = [f"{r[0]} {r[1]}" for r in data]; good = [r[3] for r in data]; bad = [r[4] for r in data]
            else:  # 고객사별
                xs = [f"{r[0]} {r[1]}" for r in data]; good = [r[3] for r in data]; bad = [0]*len(data)

            fig = Figure(figsize=(11, 5.5), dpi=120)
            ax = fig.add_subplot(111)
            x = range(len(xs))
            ax.bar(x, good, label='양품/실적', color='#26A69A')
            ax.bar(x, bad, bottom=good, label='불량', color='#EF5350')
            ax.set_xticks(list(x))
            ax.set_xticklabels(xs, rotation=45, ha='right', fontsize=8)
            ax.set_title(f"{kind} 통계 ({meta['period']}) {meta['from']}~{meta['to']}",
                         fontsize=13, fontweight='bold')
            ax.grid(axis='y', linestyle='--', alpha=0.5)
            if any(bad): ax.legend()
            fig.tight_layout()

            default_name = f"통계_{kind}_{meta['period']}_{meta['from']}_{meta['to']}.png"
            path = filedialog.asksaveasfilename(
                defaultextension='.png',
                filetypes=[('PNG 이미지', '*.png')],
                initialfile=default_name)
            if path:
                fig.savefig(path, dpi=150, bbox_inches='tight')
                messagebox.showinfo("저장 완료", f"그래프 저장됨:\n{path}")

        # 버튼들
        color_btn(ctrl, "조회", _query, theme='action', size=12).pack(side='left', padx=(20, 4))
        color_btn(ctrl, "CSV 저장", _save_csv, theme='export', size=12).pack(side='left', padx=4)
        color_btn(ctrl, "인쇄", _save_print, theme='export', size=12).pack(side='left', padx=4)
        color_btn(ctrl, "그래프 저장", _save_graph, theme='export', size=12).pack(side='left', padx=4)

        # 초기 활성 탭 표시 (일별)
        tab_btns['일별'].config(bg=C['primary'], fg='white')
        _query()

    # ========================================================
    # 일자별 이력 (전 메뉴 입력 데이터를 등록일로 조회)
    # ========================================================
    def _pg_history(self):
        p = self.page_area
        page_header(p, "일자별 이력",
                    "  프로그램에 입력한 모든 데이터를 등록 날짜별로 조회 / 저장")

        # ── 컨트롤 바 (날짜 선택 + 종류 + 조회) ──
        ctrl = tk.Frame(p, bg='white', padx=14, pady=12)
        ctrl.pack(fill='x', padx=20, pady=(8, 4))

        now = datetime.now()
        tk.Label(ctrl, text="시작:", font=('Malgun Gothic', 10),
                 bg='white').pack(side='left', padx=(4, 4))
        from_var = tk.StringVar(value=(now - timedelta(days=7)).strftime('%Y-%m-%d'))
        from_e = tk.Entry(ctrl, textvariable=from_var, width=12,
                          font=('Malgun Gothic', 11), justify='center')
        from_e.pack(side='left', padx=2)

        tk.Label(ctrl, text="종료:", font=('Malgun Gothic', 10),
                 bg='white').pack(side='left', padx=(10, 4))
        to_var = tk.StringVar(value=now.strftime('%Y-%m-%d'))
        to_e = tk.Entry(ctrl, textvariable=to_var, width=12,
                        font=('Malgun Gothic', 11), justify='center')
        to_e.pack(side='left', padx=2)

        tk.Label(ctrl, text="종류:", font=('Malgun Gothic', 10),
                 bg='white').pack(side='left', padx=(16, 4))
        kind_var = tk.StringVar(value='전체')
        kinds = ['전체', '수주', '작업지시', '생산실적', '품질검사', '출하',
                 '품목', '고객사', '설비']
        kind_combo = ttk.Combobox(ctrl, textvariable=kind_var, values=kinds,
                                  state='readonly', width=12,
                                  font=('Malgun Gothic', 11))
        kind_combo.pack(side='left', padx=2)

        # ── 요약 카드 ──
        sum_frame = tk.Frame(p, bg=C['bg'])
        sum_frame.pack(fill='x', padx=20, pady=(6, 4))
        sum_lbls = {}
        for k, color in [('총 건수', '#455A64'), ('수주', '#26A69A'),
                         ('작업지시', '#FFA726'), ('생산실적', '#EF5350'),
                         ('검사', '#AB47BC'), ('출하', '#FF7043')]:
            box = tk.Frame(sum_frame, bg='white', padx=10, pady=8)
            box.pack(side='left', expand=True, fill='x', padx=3)
            tk.Label(box, text=k, font=('Malgun Gothic', 9),
                     fg='#90A4AE', bg='white').pack()
            l = tk.Label(box, text='-', font=('Malgun Gothic', 14, 'bold'),
                         fg=color, bg='white')
            l.pack()
            sum_lbls[k] = l

        # ── 결과 트리 ──
        tf = tk.Frame(p, bg='white'); tf.pack(fill='both', expand=True, padx=20, pady=(4, 12))
        cols = ('등록일시', '종류', '식별번호', '내용', '담당')
        widths = (160, 100, 180, 350, 120)
        tree = ttk.Treeview(tf, columns=cols, show='headings', height=18)
        for c, w in zip(cols, widths):
            tree.heading(c, text=c); tree.column(c, width=w, anchor='center')
        sy = ttk.Scrollbar(tf, orient='vertical', command=tree.yview)
        sx = ttk.Scrollbar(tf, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side='right', fill='y'); sx.pack(side='bottom', fill='x')
        tree.pack(fill='both', expand=True)

        # 종류별 색상 태그
        tree.tag_configure('수주',     background='#E0F2F1')
        tree.tag_configure('작업지시', background='#FFF3E0')
        tree.tag_configure('생산실적', background='#FFEBEE')
        tree.tag_configure('품질검사', background='#F3E5F5')
        tree.tag_configure('출하',     background='#FBE9E7')
        tree.tag_configure('품목',     background='#E8EAF6')
        tree.tag_configure('고객사',   background='#E0F7FA')
        tree.tag_configure('설비',     background='#EFEBE9')

        def _query():
            for r in tree.get_children(): tree.delete(r)
            f = (from_e.get() or from_var.get()).strip()
            t = (to_e.get() or to_var.get()).strip()
            if not f or not t:
                messagebox.showerror("오류", "시작일/종료일을 입력하세요."); return
            kind = (kind_combo.get() or kind_var.get()).strip()

            rows = []  # (datetime_str, 종류, 식별번호, 내용, 담당)

            # 수주
            if kind in ('전체', '수주'):
                for r in self.db.query("""
                    SELECT o.created_at, o.order_no, c.name, i.name, o.quantity, o.status,
                           COALESCE(u.name, '-')
                    FROM orders o
                    LEFT JOIN customers c ON o.customer_id=c.id
                    LEFT JOIN items i ON o.item_id=i.id
                    LEFT JOIN users u ON o.created_by=u.id
                    WHERE date(o.created_at) BETWEEN ? AND ?
                    ORDER BY o.created_at DESC
                """, (f, t)):
                    rows.append((r[0], '수주', r[1],
                                 f"{r[2] or '-'} / {r[3] or '-'} / {r[4]}개 [{r[5]}]",
                                 r[6]))

            # 작업지시
            if kind in ('전체', '작업지시'):
                for r in self.db.query("""
                    SELECT wo.created_at, wo.wo_no, o.order_no, i.name,
                           wo.process, wo.plan_qty, wo.status, COALESCE(u.name,'-')
                    FROM work_orders wo
                    LEFT JOIN orders o ON wo.order_id=o.id
                    LEFT JOIN items i ON o.item_id=i.id
                    LEFT JOIN users u ON wo.worker_id=u.id
                    WHERE date(wo.created_at) BETWEEN ? AND ?
                    ORDER BY wo.created_at DESC
                """, (f, t)):
                    rows.append((r[0], '작업지시', r[1],
                                 f"[{r[2] or '-'}] {r[3] or '-'} / {r[4]} / 계획 {r[5]} [{r[6]}]",
                                 r[7]))

            # 생산실적
            if kind in ('전체', '생산실적'):
                for r in self.db.query("""
                    SELECT pr.created_at, wo.wo_no, i.name,
                           pr.qty, pr.defect_qty, COALESCE(u.name,'-'), pr.work_date
                    FROM production_records pr
                    LEFT JOIN work_orders wo ON pr.wo_id=wo.id
                    LEFT JOIN orders o ON wo.order_id=o.id
                    LEFT JOIN items i ON o.item_id=i.id
                    LEFT JOIN users u ON pr.worker_id=u.id
                    WHERE date(pr.created_at) BETWEEN ? AND ?
                    ORDER BY pr.created_at DESC
                """, (f, t)):
                    rows.append((r[0], '생산실적', r[1] or '-',
                                 f"{r[2] or '-'} / 양품 {r[3]} / 불량 {r[4]} (작업일 {r[6]})",
                                 r[5]))

            # 품질검사
            if kind in ('전체', '품질검사'):
                for r in self.db.query("""
                    SELECT i.created_at, o.order_no, it.name, i.inspect_type,
                           i.result, i.sample_qty, i.defect_qty, COALESCE(u.name,'-')
                    FROM inspections i
                    LEFT JOIN orders o ON i.order_id=o.id
                    LEFT JOIN items it ON o.item_id=it.id
                    LEFT JOIN users u ON i.inspector_id=u.id
                    WHERE date(i.created_at) BETWEEN ? AND ?
                    ORDER BY i.created_at DESC
                """, (f, t)):
                    rows.append((r[0], '품질검사', r[1] or '-',
                                 f"{r[2] or '-'} / {r[3]} / {r[4]} (샘플 {r[5]}, 불량 {r[6]})",
                                 r[7]))

            # 출하
            if kind in ('전체', '출하'):
                for r in self.db.query("""
                    SELECT sh.created_at, sh.ship_no, o.order_no, c.name, i.name,
                           sh.quantity, COALESCE(u.name,'-')
                    FROM shipments sh
                    LEFT JOIN orders o ON sh.order_id=o.id
                    LEFT JOIN customers c ON o.customer_id=c.id
                    LEFT JOIN items i ON o.item_id=i.id
                    LEFT JOIN users u ON sh.shipped_by=u.id
                    WHERE date(sh.created_at) BETWEEN ? AND ?
                    ORDER BY sh.created_at DESC
                """, (f, t)):
                    rows.append((r[0], '출하', r[1],
                                 f"[{r[2] or '-'}] {r[3] or '-'} / {r[4] or '-'} / {r[5]}개",
                                 r[6]))

            # 품목 (created_at 컬럼이 없을 수 있음 - active=1만)
            if kind in ('전체', '품목'):
                try:
                    for r in self.db.query("""
                        SELECT it.created_at, it.part_no, it.name, it.spec,
                               it.material, COALESCE(c.name,'-')
                        FROM items it LEFT JOIN customers c ON it.customer_id=c.id
                        WHERE date(it.created_at) BETWEEN ? AND ?
                        ORDER BY it.created_at DESC
                    """, (f, t)):
                        rows.append((r[0] or '-', '품목', r[1] or '-',
                                     f"{r[2] or '-'} / {r[3] or '-'} / {r[4] or '-'}",
                                     r[5]))
                except Exception: pass

            # 고객사
            if kind in ('전체', '고객사'):
                try:
                    for r in self.db.query("""
                        SELECT created_at, code, name, contact, phone
                        FROM customers
                        WHERE date(created_at) BETWEEN ? AND ?
                        ORDER BY created_at DESC
                    """, (f, t)):
                        rows.append((r[0] or '-', '고객사', r[1] or '-',
                                     f"{r[2] or '-'} ({r[3] or '-'} / {r[4] or '-'})",
                                     '-'))
                except Exception: pass

            # 설비
            if kind in ('전체', '설비'):
                try:
                    for r in self.db.query("""
                        SELECT created_at, code, name, process, spec, status
                        FROM equipments
                        WHERE date(created_at) BETWEEN ? AND ?
                        ORDER BY created_at DESC
                    """, (f, t)):
                        rows.append((r[0] or '-', '설비', r[1] or '-',
                                     f"{r[2] or '-'} / {r[3] or '-'} / {r[4] or '-'} [{r[5] or '-'}]",
                                     '-'))
                except Exception: pass

            # 등록일시 내림차순 정렬
            rows.sort(key=lambda x: str(x[0]), reverse=True)

            # 종류별 카운트
            counts = {'수주':0, '작업지시':0, '생산실적':0, '품질검사':0, '출하':0,
                      '품목':0, '고객사':0, '설비':0}
            for r in rows:
                counts[r[1]] = counts.get(r[1], 0) + 1
                tree.insert('', 'end', values=r, tags=(r[1],))

            sum_lbls['총 건수'].config(text=f"{len(rows):,}")
            sum_lbls['수주'].config(text=f"{counts['수주']:,}")
            sum_lbls['작업지시'].config(text=f"{counts['작업지시']:,}")
            sum_lbls['생산실적'].config(text=f"{counts['생산실적']:,}")
            sum_lbls['검사'].config(text=f"{counts['품질검사']:,}")
            sum_lbls['출하'].config(text=f"{counts['출하']:,}")

        # 콤보 변경 시 자동 조회
        kind_combo.bind('<<ComboboxSelected>>', lambda e: _query())
        # 초기 조회
        _query()

        # 액션 버튼
        bf = tk.Frame(ctrl, bg='white')
        bf.pack(side='left', padx=(20, 0))
        color_btn(bf, "조회", _query, theme='action', size=11, padx=14, pady=6).pack(side='left', padx=4)

    # ========================================================
    # 품목 관리 (메뉴에서 제거됨, 호환용으로 보존)
    # ========================================================
    def _pg_items(self):
        p = self.page_area
        page_header(p, "품목 관리", "  기계부품 생산품명 마스터")

        customers = self.db.query("SELECT id, name FROM customers WHERE active=1")
        cus_names = [r[1] for r in customers]

        f = tk.Frame(p, bg='white', padx=20, pady=12); f.pack(fill='x', padx=20, pady=12)
        vs = {k: tk.StringVar() for k in ('part_no','name','spec','material','weight','customer')}
        entries = {}  # IME 미커밋 대비 직접 .get() 백업

        def _lbl(r, c, t): make_label(f, t, size=9, color=C['secondary'], bg='white').grid(row=r, column=c, sticky='w', padx=6, pady=3)

        def _ent(key, w):
            e = make_entry(f, vs[key], w)
            entries[key] = e
            return e

        _lbl(0, 0, "생산품명 *"); _ent('name', 26).grid(row=0, column=1, columnspan=3, padx=4)
        _lbl(0, 4, "재질");     _ent('material', 12).grid(row=0, column=5, padx=4)
        _lbl(1, 0, "규격");     _ent('spec', 26).grid(row=1, column=1, columnspan=2, padx=4, pady=4, sticky='w')
        _lbl(1, 3, "단중(kg)"); _ent('weight', 10).grid(row=1, column=4, padx=4)
        _lbl(2, 0, "고객사")
        cus_combo = make_combo(f, vs['customer'], cus_names, width=22)
        cus_combo.grid(row=2, column=1, padx=4, columnspan=2, sticky='w', pady=4)
        entries['customer'] = cus_combo

        def _val(key):
            try:
                v = entries[key].get()
            except Exception:
                v = ''
            if not v:
                v = vs[key].get()
            return (v or '').strip()

        wrap = tk.Frame(p, bg=C['bg']); wrap.pack(fill='both', expand=True, padx=20, pady=6)
        cols = ('생산품명','규격','재질','단중','고객사')
        tree = make_tree(wrap, cols, [110, 200, 220, 100, 70, 160], height=16)

        def _load():
            rows = self.db.query("""
                SELECT i.name, i.spec, i.material, i.unit_weight, COALESCE(c.name,'-')
                FROM items i LEFT JOIN customers c ON i.customer_id=c.id
                WHERE i.active=1 ORDER BY i.part_no
            """)
            fill_tree(tree, rows, lambda i, r: 'even' if i%2 else '')

        def _save():
            try:
                p.focus_set(); p.update_idletasks()
            except: pass
            name     = _val('name')
            spec     = _val('spec')
            material = _val('material')
            weight_s = _val('weight')
            customer = _val('customer')

            if not name:
                messagebox.showerror("입력 오류",
                    f"생산품명은 반드시 입력해야 합니다.\n\n"
                    f"현재 입력 값:\n"
                    f"  · 생산품명 = '{name}'\n\n"
                    f"한/영 키 또는 입력 확정(Enter/Tab) 여부를 확인하세요.")
                return
            try: w = float(weight_s or 0)
            except: w = 0
            cid = None
            if customer and customer in cus_names:
                cid = customers[cus_names.index(customer)][0]
            try:
                self.db.execute("""
                    INSERT OR REPLACE INTO items(part_no,name,spec,material,unit_weight,customer_id)
                    VALUES(?,?,?,?,?,?)
                """, (name, name, spec, material, w, cid))
            except Exception as e:
                messagebox.showerror("저장 실패", f"DB 오류: {e}"); return
            messagebox.showinfo("저장 완료", f"품목 [{name}] 가 저장되었습니다.")
            for k in vs: vs[k].set('')
            _load()

        def _clear_form():
            for k in vs: vs[k].set('')

        # 통일된 컬러 액션 버튼 (저장 + 초기화)
        bf = tk.Frame(f, bg='white')
        bf.grid(row=2, column=5, padx=10, pady=4, sticky='e')
        color_btn(bf, "품목 저장", _save, theme='save').pack(side='right', padx=4)
        _reset_bar = tk.Frame(p, bg=C['bg']); _reset_bar.pack(side='bottom', fill='x', padx=20, pady=8)
        color_btn(_reset_bar, "초기화", _clear_form, theme='neutral', size=9, padx=10, pady=5).pack(side='right')
        self.root.bind('<Escape>', lambda e: _clear_form())

        _load()

    # ========================================================
    # 고객사 관리
    # ========================================================
    def _pg_customers(self):
        p = self.page_area
        page_header(p, "고객사 관리", "  거래처 등록 / 수정 / 삭제")

        f = tk.Frame(p, bg='white', padx=20, pady=12); f.pack(fill='x', padx=20, pady=12)
        vs = {k: tk.StringVar() for k in ('code','name','contact','phone','email')}
        entries = {}  # 직접 .get() 호출용 백업
        labels = [("코드 *","code",12), ("회사명 *","name",24), ("담당자","contact",14),
                  ("연락처","phone",16), ("이메일","email",22)]
        for i, (lbl, key, w) in enumerate(labels):
            make_label(f, lbl, size=9, color=C['secondary'], bg='white').grid(row=i//3, column=(i%3)*2, sticky='w', padx=6, pady=3)
            e = make_entry(f, vs[key], w)
            e.grid(row=i//3, column=(i%3)*2+1, padx=4)
            entries[key] = e

        def _val(key):
            """StringVar/Entry 양쪽에서 값 읽기 (IME 미커밋 대비)."""
            try:
                v = entries[key].get()
            except Exception:
                v = ''
            if not v:
                v = vs[key].get()
            return (v or '').strip()

        # 편집 상태 추적용 (UI 라벨 없음)
        edit_id = [None]

        class _NoLbl:
            def config(self, **kw): pass
        mode_lbl = _NoLbl()

        wrap = tk.Frame(p, bg=C['bg']); wrap.pack(fill='both', expand=True, padx=20, pady=6)
        cols = ('코드','회사명','담당자','연락처','이메일')
        tree = make_tree(wrap, cols, [100, 220, 130, 150, 200], height=16)

        def _load():
            rows = self.db.query("SELECT code,name,contact,phone,email FROM customers WHERE active=1 ORDER BY code")
            fill_tree(tree, rows, lambda i, r: 'even' if i%2 else '')

        def _clear_form():
            edit_id[0] = None
            for k in vs: vs[k].set('')
            mode_lbl.config(text="새 고객사 입력 중", fg='#2E7D32', bg='#E8F5E9')
            tree.selection_remove(tree.selection())

        def _save_new():
            # 포커스가 Entry에 있으면 IME 입력 강제 커밋
            try:
                p.focus_set(); p.update_idletasks()
            except: pass

            # 수정 모드에서 등록 버튼 누르면 폼만 비움 (취소)
            if edit_id[0] is not None:
                if not messagebox.askyesno("확인", "지금은 기존 고객사를 '수정' 하는 화면입니다.\n신규 입력으로 전환할까요?\n\n(예 = 폼 비우기, 아니오 = 그대로)"):
                    return
                _clear_form(); return

            code    = _val('code')
            name    = _val('name')
            contact = _val('contact')
            phone   = _val('phone')
            email   = _val('email')

            if not code or not name:
                messagebox.showerror("입력 오류",
                    f"코드와 회사명은 반드시 입력해야 합니다.\n\n"
                    f"현재 입력 값:\n"
                    f"  · 코드 = '{code}'\n"
                    f"  · 회사명 = '{name}'\n\n"
                    f"한/영 키가 잠겨있는지, 입력이 확정(Enter)되었는지 확인하세요.")
                return

            # 중복 검사 — 활성/삭제된 고객 모두 확인
            dup = self.db.query("SELECT id, active FROM customers WHERE code=?", (code,))
            if dup:
                row_id, active = dup[0][0], dup[0][1]
                if active == 1:
                    messagebox.showerror("등록 실패",
                        f"코드 [{code}] 는 이미 사용 중입니다.\n다른 코드를 입력하세요.")
                    return
                else:
                    if not messagebox.askyesno("복원",
                        f"코드 [{code}] 는 이전에 삭제된 고객입니다.\n복원하고 정보를 새로 갱신할까요?"):
                        return
                    self.db.execute("""
                        UPDATE customers SET name=?, contact=?, phone=?, email=?, active=1
                        WHERE id=?
                    """, (name, contact, phone, email, row_id))
                    messagebox.showinfo("복원 완료", f"고객사 [{name}] 가 복원/등록되었습니다.")
                    _clear_form(); _load(); return

            # 신규 등록
            try:
                self.db.execute("""
                    INSERT INTO customers(code,name,contact,phone,email)
                    VALUES(?,?,?,?,?)
                """, (code, name, contact, phone, email))
            except Exception as e:
                messagebox.showerror("등록 실패", f"DB 오류: {e}"); return
            messagebox.showinfo("등록 완료", f"고객사 [{name}] 가 등록되었습니다.")
            _clear_form(); _load()

        def _update():
            try:
                p.focus_set(); p.update_idletasks()
            except: pass
            if edit_id[0] is None:
                messagebox.showwarning("수정", "수정할 고객사를 목록에서 선택하세요."); return
            code = _val('code'); name = _val('name')
            if not code or not name:
                messagebox.showerror("오류", "코드/회사명은 필수입니다."); return
            if not messagebox.askyesno("확인", f"고객사 [{name}] 의 내용을 수정하시겠습니까?"):
                return
            self.db.execute("""
                UPDATE customers SET code=?, name=?, contact=?, phone=?, email=?
                WHERE id=?
            """, (code, name, _val('contact'), _val('phone'), _val('email'), edit_id[0]))
            messagebox.showinfo("완료", "고객사 수정 완료!")
            _clear_form(); _load()

        def _delete():
            if edit_id[0] is None:
                messagebox.showwarning("삭제", "삭제할 고객사를 목록에서 선택하세요."); return
            cnt = self.db.query("SELECT COUNT(*) FROM orders WHERE customer_id=?", (edit_id[0],))[0][0]
            warn = ""
            if cnt:
                warn = f"\n\n⚠ 이 고객사의 수주 {cnt}건이 존재합니다.\n   삭제하면 수주 목록에서 거래처가 표시되지 않습니다.\n   (수주 데이터 자체는 보존)"
            if not messagebox.askyesno("삭제 확인",
                f"고객사 [{vs['name'].get()}] 를 삭제하시겠습니까?{warn}"):
                return
            # 소프트 삭제 (active=0) — 관련 수주 데이터 보존
            self.db.execute("UPDATE customers SET active=0 WHERE id=?", (edit_id[0],))
            messagebox.showinfo("삭제 완료", f"고객사 [{vs['name'].get()}] 가 삭제되었습니다.")
            _clear_form(); _load()

        def _on_select(e):
            s = tree.selection()
            if not s: return
            v = tree.item(s[0])['values']
            row = self.db.query("SELECT id,code,name,contact,phone,email FROM customers WHERE code=?", (v[0],))
            if not row: return
            r = row[0]
            edit_id[0] = r[0]
            vs['code'].set(r[1] or '');     vs['name'].set(r[2] or '')
            vs['contact'].set(r[3] or '');  vs['phone'].set(r[4] or '')
            vs['email'].set(r[5] or '')
            mode_lbl.config(text=f"기존 고객 [{r[2]}] 수정 중",
                            fg='#E65100', bg='#FFF3E0')
        tree.bind('<<TreeviewSelect>>', _on_select)

        # 4개 버튼 — Frame+Label 방식으로 macOS에서도 색상 표시
        btn_row = tk.Frame(f, bg='white')
        btn_row.grid(row=2, column=0, columnspan=8, sticky='ew', pady=(12, 0))

        # 통일된 컬러 액션 버튼
        color_btn(btn_row, "고객사 등록", _save_new, theme='save').pack(side='right', padx=5)
        color_btn(btn_row, "고객사 수정", _update, theme='update').pack(side='right', padx=5)
        color_btn(btn_row, "고객사 삭제", _delete, theme='delete').pack(side='right', padx=5)
        _reset_bar = tk.Frame(p, bg=C['bg']); _reset_bar.pack(side='bottom', fill='x', padx=20, pady=8)
        color_btn(_reset_bar, "초기화", _clear_form, theme='neutral', size=9, padx=10, pady=5).pack(side='right')
        self.root.bind('<Escape>', lambda e: _clear_form())

        _load()

    # ========================================================
    # 설비 관리
    # ========================================================
    def _pg_equipments(self):
        p = self.page_area
        page_header(p, "설비 관리", "  CNC 설비 등록")

        f = tk.Frame(p, bg='white', padx=20, pady=12); f.pack(fill='x', padx=20, pady=12)
        vs = {k: tk.StringVar() for k in ('code','name','process','spec','status')}
        vs['process'].set('일반CNC'); vs['status'].set('가동')

        def _lbl(r,c,t): make_label(f,t,size=9,color=C['secondary'],bg='white').grid(row=r,column=c,sticky='w',padx=6,pady=3)

        _lbl(0,0,"코드 *");   make_entry(f, vs['code'], 12).grid(row=0,column=1,padx=4)
        _lbl(0,2,"설비명 *"); make_entry(f, vs['name'], 24).grid(row=0,column=3,padx=4)
        _lbl(0,4,"공정 *");   make_combo(f, vs['process'], PROCESSES, width=12).grid(row=0,column=5,padx=4)
        _lbl(1,0,"생산품목");  make_entry(f, vs['spec'], 30).grid(row=1,column=1,columnspan=3,padx=4,pady=4,sticky='w')
        _lbl(1,4,"상태");     make_combo(f, vs['status'], ['가동','정비','정지'], width=10).grid(row=1,column=5,padx=4)

        wrap = tk.Frame(p, bg=C['bg']); wrap.pack(fill='both', expand=True, padx=20, pady=6)
        cols = ('코드','설비명','공정','생산품목','상태')
        tree = make_tree(wrap, cols, [100, 220, 100, 280, 80], height=16)

        def _load():
            rows = self.db.query("SELECT code,name,process,spec,status FROM equipments WHERE active=1 ORDER BY code")
            fill_tree(tree, rows, lambda i, r: 'even' if i%2 else '')

        def _save():
            if not vs['code'].get() or not vs['name'].get():
                messagebox.showerror("오류", "코드/설비명은 필수입니다."); return
            self.db.execute("""
                INSERT OR REPLACE INTO equipments(code,name,process,spec,status)
                VALUES(?,?,?,?,?)
            """, (vs['code'].get(), vs['name'].get(), vs['process'].get(),
                  vs['spec'].get(), vs['status'].get()))
            for k in ('code','name','spec'): vs[k].set('')
            _load()

        def _clear_form():
            for k in vs: vs[k].set('')
            vs['process'].set('일반CNC'); vs['status'].set('가동')

        bf = tk.Frame(f, bg='white')
        bf.grid(row=1, column=6, padx=10, pady=4, sticky='e')
        color_btn(bf, "설비 저장", _save, theme='save').pack(side='right', padx=4)
        _reset_bar = tk.Frame(p, bg=C['bg']); _reset_bar.pack(side='bottom', fill='x', padx=20, pady=8)
        color_btn(_reset_bar, "초기화", _clear_form, theme='neutral', size=9, padx=10, pady=5).pack(side='right')
        self.root.bind('<Escape>', lambda e: _clear_form())
        _load()

    # ========================================================
    # 사용자 관리
    # ========================================================
    def _pg_users(self):
        p = self.page_area
        page_header(p, "사용자 관리", "  계정 / 권한")

        f = tk.Frame(p, bg='white', padx=22, pady=14); f.pack(fill='x', padx=20, pady=12)
        make_label(f, "사용자 등록", bold=True, size=11, color=C['primary'], bg='white').grid(row=0, column=0, columnspan=8, sticky='w', pady=(0,8))

        vs = {k: tk.StringVar() for k in ('uid','pw','name','team','role')}
        vs['team'].set('생산팀'); vs['role'].set('production')

        fields = [("아이디 *",'uid',14,None), ("비밀번호 *",'pw',14,'*'), ("이름 *",'name',14,None)]
        for i, (lbl, key, w, show) in enumerate(fields):
            make_label(f, lbl, size=9, color=C['secondary'], bg='white').grid(row=1, column=i*2, sticky='w', padx=6)
            tk.Entry(f, textvariable=vs[key], font=('Malgun Gothic', 10),
                     relief='flat', bd=3, bg='#F5F5F5', width=w,
                     show=show or '').grid(row=1, column=i*2+1, padx=4, pady=4)

        make_label(f, "팀", size=9, color=C['secondary'], bg='white').grid(row=1, column=6, sticky='w', padx=6)
        make_combo(f, vs['team'], ['관리','영업팀','생산팀','품질팀','출하팀'], width=12).grid(row=1, column=7, padx=4)
        make_label(f, "권한", size=9, color=C['secondary'], bg='white').grid(row=2, column=0, sticky='w', padx=6, pady=4)
        make_combo(f, vs['role'], ['admin','sales','production','quality','shipping'], width=14).grid(row=2, column=1, padx=4)

        cols = ('아이디','이름','팀','권한','등록일')
        wrap = tk.Frame(p, bg=C['bg']); wrap.pack(fill='both', expand=True, padx=20, pady=6)
        tree = make_tree(wrap, cols, [140, 140, 140, 120, 180], height=16)

        def _load():
            rows = self.db.query("SELECT username,name,team,role,created_at FROM users ORDER BY created_at")
            fill_tree(tree, rows, lambda i, r: 'even' if i%2 else '')

        def _save():
            if not vs['uid'].get() or not vs['pw'].get() or not vs['name'].get():
                messagebox.showerror("오류", "필수 항목을 입력하세요."); return
            pw = hashlib.sha256(vs['pw'].get().encode()).hexdigest()
            self.db.execute("""
                INSERT OR REPLACE INTO users(username,password,name,team,role)
                VALUES(?,?,?,?,?)
            """, (vs['uid'].get(), pw, vs['name'].get(), vs['team'].get(), vs['role'].get()))
            messagebox.showinfo("완료", f"[{vs['uid'].get()}] 등록 완료!")
            for k in vs: vs[k].set('')
            vs['team'].set('생산팀'); vs['role'].set('production')
            _load()

        color_btn(f, "사용자 등록", _save, theme='save').grid(row=2, column=7, pady=6, sticky='e')
        _load()


# ============================================================
if __name__ == "__main__":
    ProductionApp()

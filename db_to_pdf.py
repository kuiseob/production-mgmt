# -*- coding: utf-8 -*-
"""production.db → PDF 변환 스크립트.

사용법:
    python db_to_pdf.py
    또는 더블클릭 (Windows)
"""
import os, sys, sqlite3, html, subprocess, tempfile, platform
from datetime import datetime

# ──────────────────────────────────────────────
HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "production.db")

if not os.path.isfile(DB_PATH):
    print(f"❌ DB 파일을 찾을 수 없습니다: {DB_PATH}")
    input("Enter를 눌러 종료...")
    sys.exit(1)

# ──────────────────────────────────────────────
LABELS = {
    'customers':  ('🏢 고객사',    ['코드','회사명','담당자','연락처','이메일','활성']),
    'items':      ('📦 품목',      ['품번','생산품명','규격','재질','단중','고객사ID','활성']),
    'equipments': ('🏭 설비',      ['코드','설비명','공정','생산품목','상태','활성']),
    'orders':     ('📋 수주',      ['주문번호','고객사ID','품목ID','수량','수주일','납기일','상태','비고','생성자','생성일시']),
    'work_orders':('🔧 작업지시',  ['작업번호','주문ID','공정','계획수량','실적수량','불량수량','설비ID','담당ID','시작','종료','상태','순번','생성일시']),
    'production_records':('⚙️ 생산실적', ['ID','작업ID','작업일','양품','불량','담당ID','설비ID','비고','생성일시']),
    'inspections':('🔍 품질검사',  ['ID','주문ID','검사구분','검사일','샘플','결과','불량','불량사유','검사자ID','비고','생성일시']),
    'shipments':  ('🚚 출하',      ['ID','주문ID','출하번호','출하일','수량','담당ID','비고','생성일시']),
    'users':      ('👤 사용자',    ['ID','계정','비밀번호','이름','소속','역할','생성일시']),
}
ORDER = ['orders','work_orders','production_records','inspections',
         'shipments','customers','items','equipments','users']

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
th { background:#00695C; color:white; padding: 5px 6px; text-align:center;
     font-size: 8.5pt; }
td { border: 1px solid #CFD8DC; padding: 4px 6px; vertical-align: top; }
tbody tr:nth-child(even) { background:#F5F7F8; }
.no-data { color:#999; font-style: italic; padding: 8px; }
.footer { text-align:center; color:#90A4AE; font-size: 8pt; margin-top: 30px; }
"""

def safe(v):
    if v is None: return ''
    s = str(v)
    if len(s) > 80: s = s[:78] + '..'
    return html.escape(s)

# ──────────────────────────────────────────────
print(f"📂 DB 읽는 중: {DB_PATH}")
conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row
c = conn.cursor()
total_rows = sum(c.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] for t in ORDER)
print(f"   총 {total_rows} 행")

now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
parts = [f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>SEO JIN PRECISION 데이터 백업</title>
<style>{CSS}</style></head>
<body>
<div class="cover">
  <h1>SEO JIN PRECISION CO.</h1>
  <h2>생산관리 시스템 — 데이터베이스 전체 백업</h2>
  <div class="meta">
    파일: production.db<br>
    생성일시: {now}<br>
    총 데이터: {total_rows:,} 행<br>
    테이블 수: {len(ORDER)} 개
  </div>
</div>
"""]

for tname in ORDER:
    label, _ = LABELS[tname]
    rows = c.execute(f"SELECT * FROM {tname} ORDER BY rowid").fetchall()
    cols = [r[1] for r in c.execute(f"PRAGMA table_info({tname})")]
    parts.append(f'<div class="section"><h2 class="tname">{label}<span class="cnt">{len(rows):,} 행</span></h2>')
    if not rows:
        parts.append('<div class="no-data">데이터 없음</div></div>')
        continue
    parts.append('<table><thead><tr>')
    for col in cols: parts.append(f'<th>{safe(col)}</th>')
    parts.append('</tr></thead><tbody>')
    for r in rows:
        parts.append('<tr>')
        for col in cols: parts.append(f'<td>{safe(r[col])}</td>')
        parts.append('</tr>')
    parts.append('</tbody></table></div>')

parts.append(f'<div class="footer">© SEO JIN PRECISION CO. — Generated {now}</div></body></html>')

# 임시 HTML 저장
tmp_html = tempfile.NamedTemporaryFile(suffix='.html', delete=False, mode='w', encoding='utf-8')
tmp_html.write('\n'.join(parts))
tmp_html.close()
print(f"📄 HTML 생성: {tmp_html.name}")

# PDF 출력
pdf_name = f"DB_백업_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
pdf_path = os.path.join(HERE, pdf_name)

# Chrome / Edge / Brave 등 자동 탐색
chrome_paths = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]
chrome = next((p for p in chrome_paths if os.path.isfile(p)), None)

if chrome:
    print(f"🌐 브라우저: {chrome}")
    cmd = [chrome, "--headless", "--disable-gpu", "--no-pdf-header-footer",
           "--no-margins", f"--print-to-pdf={pdf_path}",
           f"file:///{tmp_html.name.replace(os.sep, '/')}"]
    subprocess.run(cmd, capture_output=True)
    if os.path.isfile(pdf_path):
        print(f"✅ PDF 저장 완료: {pdf_path}")
        # 자동으로 열기
        if platform.system() == 'Windows':
            os.startfile(pdf_path)
        elif platform.system() == 'Darwin':
            subprocess.run(['open', pdf_path])
        else:
            subprocess.run(['xdg-open', pdf_path])
    else:
        print(f"❌ PDF 생성 실패. HTML만 사용 가능: {tmp_html.name}")
else:
    print(f"⚠ Chrome/Edge를 찾을 수 없습니다.")
    print(f"   HTML로만 저장됨: {tmp_html.name}")
    if platform.system() == 'Windows':
        os.startfile(tmp_html.name)

print("\n완료. Enter를 눌러 종료...")
try: input()
except: pass

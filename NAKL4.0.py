# -*- coding: utf-8 -*-
# ============================================================
#  NAKL4.4  —  print_labels.py
#  Термоэтикетки 100 × 35 мм.
#
#  Изменения от 4.3:
#    • Убрана кнопка "Показать" (автозагрузка при выборе).
#    • Принтер по умолчанию: HPRT LPG4.
# ============================================================

import io
import os
import sys
import time
import tempfile
import argparse
from pathlib import Path
from datetime import datetime

import mysql.connector
import qrcode
from PIL import Image, ImageFile

from reportlab.lib.units import mm
from reportlab.lib.colors import black, grey
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader


ImageFile.LOAD_TRUNCATED_IMAGES = True

VERSION = 'NAKL4.4'
DEFAULT_PRINTER = 'HPRT LPG4'


# ================= КОНФИГ =================
DB_HOST = '192.168.0.10'
DB_USER = 'workshop'
DB_PASSWORD = 'w0rK5h0p'
DB_CHARSET = 'utf8mb4'
SRC_DB = 'CCALC'
DST_DB = 'workshop'

# ---- База Access ----
MDB_PATH = r'E:\ACCESS\BACKUP\15.10.2026\ACCESS\2010.mdb'
MDB_TABLE = 'лист1'
MDB_CACHE_TTL = 15 * 60  # Время жизни кэша (15 минут в секундах)

# ---- Глобальный кэш Access ----
_mdb_cache = None
_mdb_cache_time = 0

if sys.platform == 'linux':
    IMG_DIR = Path(r'/mnt/smb/ПВХ/JPG')
elif sys.platform == 'win32':
    IMG_DIR = Path(r'\\Synas\work\ПВХ\JPG')
else:
    IMG_DIR = None

# ---- Размер этикетки ----
LABEL_W = 100 * mm
LABEL_H = 35 * mm
PAGE_SIZE = (LABEL_W, LABEL_H)

# ---- QR + DETAL ----
QR_SIZE = 22 * mm
QR_X    = 1.5 * mm
QR_Y    = 3.5 * mm
DETAL_Y = 2.2 * mm

# ---- Начала текста ----
LEFT_X = 1.5 * mm
TXT_X  = QR_X + QR_SIZE + 2 * mm

# ---- Картинка ----
IMG_RIGHT_EDGE = 115 * mm
IMG_MAX_H      = 45 * mm
IMG_V_CENTER   = True

# ---- Фурнитура ----
FRN_MAX_W      = 45 * mm
FRN_MAX_LINES  = 3

# ---- Примечание ----
PRIM_MAX_LINES = 2

# ---- Вертикальные базовые линии строк ----
Y_L1  = 29.0 * mm
Y_L2  = 25.5 * mm
Y_L3  = 20.0 * mm
Y_L4  = 16.5 * mm
Y_L4B = 13.5 * mm
Y_L4C = 10.5 * mm
Y_L7  = 8.0 * mm
Y_L8  = 5.0 * mm
Y_L8B = 2.5 * mm

# ---- Шрифты ----
FONT_NAME = 'Arial'
FONT_BOLD = 'Arial-Bold'
ARIAL      = r'C:\Windows\Fonts\arial.ttf'
ARIAL_BOLD = r'C:\Windows\Fonts\arialbd.ttf'

SZ_ZAKNUM  = 24
SZ_ROLE    = 18
SZ_MATNAME = 12
SZ_COLOR   = 12
SZ_LENGTH  = 14
SZ_ANGLES  = 14
SZ_FRN     = 8
SZ_SEAL    = 8
SZ_PRIM    = 8
SZ_QR_VAL  = 7
SZ_NOFILE  = 7

ORIENT_ORDER = {'В': 0, 'Н': 1, 'Л': 2, 'П': 3}


def register_fonts():
    pdfmetrics.registerFont(TTFont(FONT_NAME, ARIAL))
    pdfmetrics.registerFont(TTFont(FONT_BOLD, ARIAL_BOLD))


def connect():
    return mysql.connector.connect(
        host=DB_HOST, user=DB_USER, password=DB_PASSWORD,
        charset=DB_CHARSET,
    )


# ================= ACCESS ЧЕРЕЗ OLEDB (С КЭШЕМ) =================
def fetch_orders_from_mdb(force_refresh=False):
    global _mdb_cache, _mdb_cache_time
    now = time.time()

    if not force_refresh and _mdb_cache is not None and (now - _mdb_cache_time < MDB_CACHE_TTL):
        return _mdb_cache

    print(f"[{datetime.now().strftime('%H:%M:%S')}] Обновление кэша Access...")

    if not os.path.exists(MDB_PATH):
        raise RuntimeError(f'Файл не найден: {MDB_PATH}')

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise RuntimeError('Не установлен pywin32. Выполните: pip install pywin32')

    pythoncom.CoInitialize()
    try:
        providers = [
            'Microsoft.ACE.OLEDB.12.0',
            'Microsoft.Jet.OLEDB.4.0',
        ]
        conn = None
        last_err = None

        for provider in providers:
            try:
                conn_str = (
                    f'Provider={provider};'
                    f'Data Source={MDB_PATH};'
                    f'Persist Security Info=False;'
                )
                conn = win32com.client.Dispatch("ADODB.Connection")
                conn.Open(conn_str)
                break
            except Exception as e:
                last_err = e
                conn = None

        if not conn:
            raise RuntimeError(
                f'Не удалось открыть Access через OLEDB.\n'
                f'Проверьте, установлен ли Microsoft Office или Access Runtime.\n\n'
                f'Последняя ошибка: {last_err}')

        sql = f"""
            SELECT DISTINCT
                   [Заказчи]         AS ZAKNUM,
                   [Дата заявки]     AS DateOrder,
                   [Дата отгрузки]   AS DateShip
            FROM   [{MDB_TABLE}]
            WHERE  [Заказчи] IS NOT NULL
              AND  [Заказчи] <> ''
            ORDER BY [Дата заявки] DESC
        """
        rs = conn.Execute(sql)[0]

        rows = []
        while not rs.EOF:
            rows.append((
                rs.Fields("ZAKNUM").Value,
                rs.Fields("DateOrder").Value,
                rs.Fields("DateShip").Value,
            ))
            rs.MoveNext()

        rs.Close()
        conn.Close()

    finally:
        pythoncom.CoUninitialize()

    def _fmt(v):
        if v is None:
            return ''
        if hasattr(v, 'strftime'):
            return v.strftime('%d.%m.%Y')
        s = str(v).strip()
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try:
                return datetime.strptime(s[:19], fmt).strftime('%d.%m.%Y')
            except Exception:
                pass
        return s

    result = []
    for r in rows:
        zak = str(r[0]).strip()
        if not zak:
            continue
        result.append((zak, _fmt(r[1]), _fmt(r[2])))

    _mdb_cache = result
    _mdb_cache_time = time.time()
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Кэш Access обновлён: {len(result)} заказов.")
    return result


# ================= MYSQL =================
def fetch_parts(zaknum):
    conn = connect()
    cur = conn.cursor(dictionary=True)

    cur.execute(f"""
        SELECT
            ZAKNUM, KONNAME, NUMIZDELIE, DETAL, PRIMPROIZV,
            BCI, PROFILTYPE, KONID, GRI, RAMAID,
            MATNAME, MATART, ELEMID, OBJID, MATCOLOR,
            LENGTH, ZAKID, ANGLE_L, ANGLE_R, STAR,
            ORIENT, ORIENT2, IMPOST, RUCHKA, ZAMOK, DUGA
        FROM `{DST_DB}`.`CCALC`
        WHERE ZAKNUM = %s
    """, (zaknum,))
    parts = cur.fetchall()
    if not parts:
        cur.close(); conn.close(); return []

    cur.execute(f"""
        SELECT ZAKNUM, KONID, NAME, TYPENAME, NAPRNAME
        FROM `{DST_DB}`.`CCALC_FRN`
        WHERE ZAKNUM = %s
        ORDER BY ZAKNUM, KONID, NAMEID
    """, (zaknum,))
    frn = {}
    for r in cur.fetchall():
        frn.setdefault((r['ZAKNUM'], r['KONID']), r)

    # Уплотнение
    cur.execute(f"""
        SELECT ZAKNUM, KONID, MIN(MATCOLOR) AS SEAL_COLOR
        FROM `{DST_DB}`.`CCALC_TBL`
        WHERE ZAKNUM = %s AND MATNAME LIKE %s
        GROUP BY ZAKNUM, KONID
    """, (zaknum, '%Уплотнение%'))
    seals = {(r['ZAKNUM'], r['KONID']): r['SEAL_COLOR'] for r in cur.fetchall()}

    # Штапик
    cur.execute(f"""
        SELECT ZAKNUM, KONID, MIN(MATNAME) AS STAP_NAME
        FROM `{DST_DB}`.`CCALC_TBL`
        WHERE ZAKNUM = %s AND MATNAME LIKE %s
        GROUP BY ZAKNUM, KONID
    """, (zaknum, '%Штапик%'))
    staps = {(r['ZAKNUM'], r['KONID']): r['STAP_NAME'] for r in cur.fetchall()}

    for p in parts:
        key = (p['ZAKNUM'], p['KONID'])
        f = frn.get(key)
        p['FRN_NAME']     = f['NAME']     if f else None
        p['FRN_TYPENAME'] = f['TYPENAME'] if f else None
        p['FRN_NAPRNAME'] = f['NAPRNAME'] if f else None
        p['SEAL_COLOR']   = seals.get(key)
        p['STAP_NAME']    = staps.get(key)

    cur.close(); conn.close()
    return parts


# ================= ФИЛЬТР + СОРТИРОВКА =================
def _int_or_inf(v):
    try:
        return int(v) if v is not None else 10**9
    except (TypeError, ValueError):
        return 10**9


def sort_for_print(parts):
    ALLOWED_GRI = {'r', 's', '3', 'п', 'ш', 'w', 't', 'т'}

    def gri_of(p):
        return str(p.get('GRI') or '').strip().lower()

    def is_allowed(p):
        gri = gri_of(p)
        if gri in ALLOWED_GRI:
            return True
        if 'штульп' in str(p.get('MATNAME') or '').lower():
            return True
        return False

    filtered = [p for p in parts if is_allowed(p)]

    def key_rs(p):
        return (_int_or_inf(p.get('KONID')),
                _int_or_inf(p.get('OBJID')),
                ORIENT_ORDER.get((p.get('ORIENT') or '').strip(), 99),
                str(p.get('DETAL') or ''))

    rams    = sorted([p for p in filtered if gri_of(p) == 'r'], key=key_rs)
    stvorki = sorted([p for p in filtered if gri_of(p) == 's'], key=key_rs)
    imposts = [p for p in filtered if gri_of(p) == '3']
    porogi  = [p for p in filtered if gri_of(p) == 'п']
    shtulps = [p for p in filtered if gri_of(p) not in ('r', 's', '3', 'п')]

    groups = [
        ('РАМЫ',    rams),
        ('СТВОРКИ', stvorki),
        ('ИМПОСТА', imposts),
        ('ШТУЛЬПЫ', shtulps),
        ('ПОРОГИ',  porogi),
    ]
    return [(n, l) for n, l in groups if l]


# ================= КЕШИ =================
_image_cache = {}
_qr_cache    = {}


def get_imagereader(zaknum, konid):
    if IMG_DIR is None:
        return None
    key = (str(zaknum), str(konid))
    if key in _image_cache:
        return _image_cache[key]

    filepath = IMG_DIR / f"{zaknum} ({konid}).jpg"
    if not filepath.exists():
        _image_cache[key] = None
        return None

    try:
        img = Image.open(str(filepath))
        img.load()
        w, h = img.size
        if w > 80 and h > 80:
            img = img.crop((40, 40, w - 40, h - 40))
        if img.mode not in ('RGB', 'L'):
            img = img.convert('RGB')
        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=88)
        buf.seek(0)
        reader = ImageReader(buf)
    except Exception:
        reader = None
    _image_cache[key] = reader
    return reader


def get_qr_reader(data):
    data = str(data or '')
    if data in _qr_cache:
        return _qr_cache[data]
    if not data:
        _qr_cache[data] = None
        return None
    try:
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=10,
            border=1,
        )
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white').convert('RGB')
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        buf.seek(0)
        reader = ImageReader(buf)
    except Exception:
        reader = None
    _qr_cache[data] = reader
    return reader


# ================= ФОРМАТТЕРЫ =================
def clean_color(v):
    if not v:
        return ''
    s = str(v).strip()
    if s.startswith('#'):
        s = s.lstrip('#').strip()
    return s


def build_role(rec):
    kon = rec.get('KONID') or ''
    obj = rec.get('OBJID') or ''
    gri = str(rec.get('GRI') or '').strip().lower()
    letter = (rec.get('ORIENT2') or '') if gri == '3' else (rec.get('ORIENT') or '')
    letter = str(letter).strip()
    suffix = 'Рчк' if rec.get('RUCHKA') == -1 else ''
    base = f"K{kon}.{obj}"
    if letter:
        base += f" {letter}"
    return base + suffix


def build_orient(rec):
    t = (rec.get('FRN_TYPENAME') or '').strip()
    n = (rec.get('FRN_NAPRNAME') or '').strip()
    if t and n:
        return f"{t} – {n}"
    return t or n


def fit_text(c, text, font, size, max_w):
    text = str(text or '')
    if not text or max_w <= 0:
        return ''
    while text and c.stringWidth(text, font, size) > max_w:
        text = text[:-1]
    return text


def wrap_text(c, text, font, size, max_w, max_lines=3):
    text = str(text or '').strip()
    if not text:
        return []
    words = text.split()
    lines, cur = [], ''
    for w in words:
        trial = (cur + ' ' + w).strip()
        if c.stringWidth(trial, font, size) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
                if len(lines) >= max_lines:
                    cur = w
                    break
            cur = w
    if cur and len(lines) < max_lines:
        lines.append(cur)
    elif cur and len(lines) == max_lines:
        tail = lines[-1]
        while tail and c.stringWidth(tail + '…', font, size) > max_w:
            tail = tail[:-1]
        lines[-1] = (tail + '…') if tail else '…'
    return lines[:max_lines]


# ================= РИСОВАНИЕ =================
def draw_label(c, rec):
    """Одна этикетка 100 × 35 мм."""

    # ---- 1. Картинка ----
    reader = get_imagereader(rec.get('ZAKNUM'), rec.get('KONID'))
    if reader is not None:
        try:
            iw0, ih0 = reader.getSize()
            h = IMG_MAX_H
            w = iw0 * (h / ih0)
            img_x = IMG_RIGHT_EDGE - w
            img_y = (LABEL_H - h) / 2 if IMG_V_CENTER else 0
            c.drawImage(reader, img_x, img_y, width=w, height=h, mask='auto')
        except Exception:
            c.setStrokeColor(grey)
            c.rect(IMG_RIGHT_EDGE - 42 * mm, 0, 42 * mm, LABEL_H)
    else:
        c.setStrokeColor(grey)
        c.rect(IMG_RIGHT_EDGE - 42 * mm, 0, 42 * mm, LABEL_H)
        c.setFont(FONT_BOLD, SZ_NOFILE)
        c.setFillColor(grey)
        c.drawCentredString(IMG_RIGHT_EDGE - 21 * mm, LABEL_H / 2, 'нет файла')
        c.setFillColor(black)

    # ---- 2. QR + DETAL ----
    bc_val = str(rec.get('DETAL') or '')
    qr_reader = get_qr_reader(bc_val)
    if qr_reader is not None:
        try:
            c.drawImage(qr_reader, QR_X, QR_Y,
                        width=QR_SIZE, height=QR_SIZE,
                        preserveAspectRatio=True, mask='auto')
        except Exception:
            c.setStrokeColor(grey)
            c.rect(QR_X, QR_Y, QR_SIZE, QR_SIZE)
    else:
        c.setStrokeColor(grey)
        c.rect(QR_X, QR_Y, QR_SIZE, QR_SIZE)

    if bc_val:
        c.setFont(FONT_BOLD, SZ_QR_VAL)
        c.setFillColor(black)
        c.drawString(QR_X, DETAL_Y, bc_val)

    LW_EDGE = LABEL_W - 1.5 * mm

    # === Строка 1 ===
    zak = str(rec.get('ZAKNUM') or '')
    c.setFont(FONT_BOLD, SZ_ZAKNUM)
    c.drawString(LEFT_X, Y_L1, zak)
    zak_w = c.stringWidth(zak, FONT_BOLD, SZ_ZAKNUM)

    role = build_role(rec)
    c.setFont(FONT_BOLD, SZ_ROLE)
    gap_role = c.stringWidth(' ', FONT_BOLD, SZ_ROLE)
    c.drawString(LEFT_X + zak_w + gap_role, Y_L1, role)

    # === Строка 2 ===
    matname = (rec.get('MATNAME') or '').strip()
    color_txt = clean_color(rec.get('MATCOLOR'))
    c.setFont(FONT_BOLD, SZ_MATNAME)
    c.drawString(LEFT_X, Y_L2, matname)
    mat_w = c.stringWidth(matname, FONT_BOLD, SZ_MATNAME)

    if color_txt:
        c.setFont(FONT_BOLD, SZ_COLOR)
        gap_color = c.stringWidth(' ', FONT_BOLD, SZ_MATNAME)
        c.drawString(LEFT_X + mat_w + gap_color, Y_L2, color_txt)

    # === Строка 3 ===
    l_str = f"L:{int(rec.get('LENGTH') or 0)}"
    al = int(rec.get('ANGLE_L') or 0)
    ar = int(rec.get('ANGLE_R') or 0)
    angles = f"{al}° {ar}°"

    c.setFont(FONT_BOLD, SZ_LENGTH)
    c.drawString(TXT_X, Y_L3, l_str)
    l_w = c.stringWidth(l_str, FONT_BOLD, SZ_LENGTH)
    gap3 = c.stringWidth('  ', FONT_BOLD, SZ_LENGTH)

    c.setFont(FONT_BOLD, SZ_ANGLES)
    c.drawString(TXT_X + l_w + gap3, Y_L3, angles)

    # === Строки 4-6: фурнитура + ПО-Лв ===
    orient = build_orient(rec)
    frn_name = (rec.get('FRN_NAME') or '').strip()

    c.setFont(FONT_BOLD, SZ_FRN)
    lines = wrap_text(c, frn_name, FONT_BOLD, SZ_FRN,
                      FRN_MAX_W, max_lines=FRN_MAX_LINES)
    y_positions = [Y_L4, Y_L4B, Y_L4C]
    last_x = TXT_X
    last_y = Y_L4

    for i, line in enumerate(lines):
        if i < len(y_positions):
            c.drawString(TXT_X, y_positions[i], line)
            last_x = TXT_X + c.stringWidth(line, FONT_BOLD, SZ_FRN)
            last_y = y_positions[i]

    if orient:
        c.setFont(FONT_BOLD, SZ_FRN + 2)
        gap_orient = c.stringWidth('  ', FONT_BOLD, SZ_FRN + 2)
        if last_x + gap_orient + c.stringWidth(orient, FONT_BOLD, SZ_FRN + 2) < TXT_X + FRN_MAX_W:
            c.drawString(last_x + gap_orient, last_y, orient)
        else:
            if len(lines) < FRN_MAX_LINES:
                c.drawString(TXT_X, y_positions[len(lines)], orient)
            else:
                c.drawString(TXT_X, y_positions[-1] - 3 * mm, orient)

    # === Строка 7: Упл. + ШТП ===
    seal = clean_color(rec.get('SEAL_COLOR'))
    stap = (rec.get('STAP_NAME') or '').strip()
    line7_parts = []
    if seal:
        line7_parts.append(f"Упл. {seal}")
    if stap:
        line7_parts.append(f"ШТП: {stap}")

    if line7_parts:
        c.setFont(FONT_BOLD, SZ_SEAL)
        c.drawString(TXT_X, Y_L7, "  ".join(line7_parts))

    # === Строка 8: примечание ===
    prim = (rec.get('PRIMPROIZV') or '').strip()
    if prim:
        c.setFont(FONT_BOLD, SZ_PRIM)
        prim_w = LW_EDGE - TXT_X
        lines = wrap_text(c, prim, FONT_BOLD, SZ_PRIM,
                          prim_w, max_lines=PRIM_MAX_LINES)
        if lines:
            c.drawString(TXT_X, Y_L8, lines[0])
        if len(lines) > 1:
            c.drawString(TXT_X, Y_L8B, lines[1])


# ================= PDF =================
def build_pdf_to(target, zaknum, parts):
    register_fonts()
    c = Canvas(target, pagesize=PAGE_SIZE)
    c.setTitle(f'{VERSION} — Наклейки {zaknum}')

    for group_name, items in sort_for_print(parts):
        for rec in items:
            try:
                draw_label(c, rec)
            except Exception as e:
                print(f"⚠ Ошибка отрисовки {rec.get('DETAL')}: {e}")
            finally:
                c.showPage()
    c.save()


def build_pdf_file(zaknum, out_path, parts):
    build_pdf_to(str(out_path), zaknum, parts)


def build_pdf_bytes(zaknum, parts):
    buf = io.BytesIO()
    build_pdf_to(buf, zaknum, parts)
    buf.seek(0)
    return buf


# ================= ПРИНТЕРЫ =================
def list_printers():
    try:
        import win32print
        printers = [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        if printers:
            return sorted(set(printers))
    except Exception:
        pass

    try:
        import subprocess
        out = subprocess.check_output(
            ['wmic', 'printer', 'get', 'name'],
            stderr=subprocess.DEVNULL, text=True, encoding='cp866')
        lines = [l.strip() for l in out.splitlines()[1:] if l.strip()]
        if lines:
            return sorted(set(lines))
    except Exception:
        pass
    return []


def print_to_printer(pdf_path, printer_name=None):
    pdf_path = str(pdf_path)
    if printer_name:
        try:
            import win32api
            win32api.ShellExecute(0, 'printto', pdf_path,
                                  f'"{printer_name}"', '.', 0)
            return True
        except Exception:
            pass
    try:
        os.startfile(pdf_path, 'print')
        return True
    except Exception:
        return False


# ================= CLI =================
def run_cli(zaknum, out_path):
    print(f'📖 {VERSION}  заказ {zaknum}')
    if IMG_DIR:
        print(f'   Картинки : {IMG_DIR}')
    print(f'   Этикетка : {LABEL_W/mm:.0f} × {LABEL_H/mm:.0f} мм')
    recs = fetch_parts(zaknum)
    if not recs:
        print('❌ Деталей не найдено'); sys.exit(1)
    total = sum(len(l) for _, l in sort_for_print(recs))
    print(f'   Деталей  : {len(recs)}   страниц: {total}')
    print(f'📄 Пишу {out_path}')
    build_pdf_file(zaknum, out_path, recs)
    print(f'✅ Готово: {out_path}')


# ================= GUI =================
def run_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    from PIL import ImageTk, Image as PILImage
    import pypdfium2 as pdfium

    class OrderPickerWindow:
        def __init__(self, parent, on_select):
            self.on_select = on_select
            self.win = tk.Toplevel(parent)
            self.win.title('Заказы из Access (лист1) — OLEDB')
            self.win.geometry('750x550')
            self.win.transient(parent)
            self.win.grab_set()

            top = ttk.Frame(self.win, padding=6)
            top.pack(side=tk.TOP, fill=tk.X)
            
            ttk.Label(top, text=f'База: {MDB_PATH}',
                      font=('Arial', 9), foreground='gray').pack(side=tk.LEFT)

            self.search_var = tk.StringVar()
            self.search_var.trace('w', lambda *a: self.filter_rows())
            ttk.Label(top, text='Поиск:').pack(side=tk.LEFT, padx=(20, 4))
            ttk.Entry(top, textvariable=self.search_var,
                      width=20, font=('Arial', 10)).pack(side=tk.LEFT)

            ttk.Button(top, text='🔄 Обновить из БД',
                       command=self.refresh_cache).pack(side=tk.RIGHT, padx=6)

            self.last_update_lbl = ttk.Label(top, text='', font=('Arial', 9), foreground='blue')
            self.last_update_lbl.pack(side=tk.RIGHT, padx=6)

            cols = ('zaknum', 'date_order', 'date_ship')
            self.tree = ttk.Treeview(self.win, columns=cols, show='headings')
            self.tree.heading('zaknum', text='ZAKNUM (Заказчик)')
            self.tree.heading('date_order', text='Дата заявки')
            self.tree.heading('date_ship', text='Дата отгрузки')
            self.tree.column('zaknum', width=260, anchor='w')
            self.tree.column('date_order', width=130, anchor='center')
            self.tree.column('date_ship', width=130, anchor='center')
            self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=6, pady=6)

            sb = ttk.Scrollbar(self.tree, orient='vertical', command=self.tree.yview)
            sb.pack(side=tk.RIGHT, fill=tk.Y)
            self.tree.configure(yscrollcommand=sb.set)

            self.tree.bind('<Double-1>', lambda e: self.pick())
            self.tree.bind('<Return>', lambda e: self.pick())

            bar = ttk.Frame(self.win, padding=6)
            bar.pack(side=tk.BOTTOM, fill=tk.X)
            ttk.Button(bar, text='Выбрать', command=self.pick).pack(side=tk.LEFT)
            ttk.Button(bar, text='Закрыть', command=self.win.destroy).pack(side=tk.LEFT, padx=6)

            self.rows_all = []
            self.load()

        def load(self, force=False):
            self.tree.delete(*self.tree.get_children())
            self.rows_all = []
            try:
                self.rows_all = fetch_orders_from_mdb(force_refresh=force)
                t = datetime.fromtimestamp(_mdb_cache_time).strftime('%H:%M:%S')
                self.last_update_lbl.config(text=f'Кэш от {t}')
            except Exception as e:
                messagebox.showerror('Ошибка чтения Access', str(e), parent=self.win)
                return

            for zak, d1, d2 in self.rows_all:
                self.tree.insert('', 'end', values=(zak, d1, d2))

        def refresh_cache(self):
            self.load(force=True)

        def filter_rows(self):
            q = self.search_var.get().strip().lower()
            self.tree.delete(*self.tree.get_children())
            for zak, d1, d2 in self.rows_all:
                if q and q not in zak.lower():
                    continue
                self.tree.insert('', 'end', values=(zak, d1, d2))

        def pick(self):
            sel = self.tree.selection()
            if not sel:
                return
            vals = self.tree.item(sel[0], 'values')
            zak = vals[0]
            self.on_select(zak)
            self.win.destroy()

    class App:
        def __init__(self, root):
            self.root = root
            root.title(f'{VERSION} — Термоэтикетки CCalc')
            root.geometry('1100x850')
            root.minsize(800, 600)

            top = ttk.Frame(root, padding=8)
            top.pack(side=tk.TOP, fill=tk.X)

            ttk.Label(top, text='Номер заказа:',
                      font=('Arial', 11)).pack(side=tk.LEFT)

            self.zak_entry = ttk.Entry(top, width=24, font=('Arial', 12))
            self.zak_entry.pack(side=tk.LEFT, padx=6)
            self.zak_entry.bind('<Return>', lambda e: self.load())
            self.zak_entry.focus_set()

            ttk.Button(top, text='📂 Заказы из Access',
                       command=self.open_order_picker).pack(side=tk.LEFT, padx=3)
            ttk.Button(top, text='📋 Из буфера',
                       command=self.paste_from_clipboard).pack(side=tk.LEFT, padx=3)
            # Кнопка "Показать" убрана по требованию
            ttk.Button(top, text='Сохранить PDF',
                       command=self.save_pdf).pack(side=tk.LEFT, padx=3)
            ttk.Button(top, text='Очистить',
                       command=self.clear).pack(side=tk.LEFT, padx=3)

            printers = list_printers()
            self.printer_var = tk.StringVar()
            ttk.Label(top, text='Принтер:',
                      font=('Arial', 10)).pack(side=tk.LEFT, padx=(20, 4))
            self.printer_cb = ttk.Combobox(
                top, textvariable=self.printer_var,
                values=printers, width=28, state='readonly')
            
            # Установка принтера по умолчанию HPRT LPG4
            if printers:
                idx = 0
                for i, p in enumerate(printers):
                    if DEFAULT_PRINTER.lower() in p.lower():
                        idx = i
                        break
                self.printer_cb.current(idx)
                
            self.printer_cb.pack(side=tk.LEFT)
            ttk.Button(top, text='🖨 Печать (обр. порядок)',
                       command=self.print_pdf).pack(side=tk.LEFT, padx=6)

            self.status = ttk.Label(root, text=f'{VERSION}  —  введите номер заказа',
                                    foreground='blue', font=('Arial', 10))
            self.status.pack(side=tk.TOP, fill=tk.X, padx=10)

            root.bind('<Control-v>', self._on_ctrl_v)
            root.bind('<Control-V>', self._on_ctrl_v)

            nav = ttk.Frame(root, padding=6)
            nav.pack(side=tk.BOTTOM, fill=tk.X)

            ttk.Button(nav, text='◀', width=3,
                       command=self.prev_page).pack(side=tk.LEFT, padx=2)

            self.page_var = tk.StringVar(value='1')
            self.page_entry = ttk.Entry(nav, textvariable=self.page_var,
                                        width=6, justify='center',
                                        font=('Arial', 11))
            self.page_entry.pack(side=tk.LEFT, padx=2)
            self.page_entry.bind('<Return>', lambda e: self.goto_page())

            ttk.Label(nav, text='из', font=('Arial', 10)).pack(side=tk.LEFT)

            self.total_label = ttk.Label(nav, text='0',
                                          font=('Arial', 11, 'bold'))
            self.total_label.pack(side=tk.LEFT, padx=2)

            ttk.Button(nav, text='▶', width=3,
                       command=self.next_page).pack(side=tk.LEFT, padx=2)

            ttk.Button(nav, text='Перейти',
                       command=self.goto_page).pack(side=tk.LEFT, padx=6)

            ttk.Button(nav, text='🖨 Печать этой страницы',
                       command=lambda: self.print_single_page(self.current_page)
                       ).pack(side=tk.RIGHT, padx=6)

            body = ttk.Frame(root)
            body.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=6)

            self.image_label = tk.Label(body, bg='#d8d8d8',
                                         borderwidth=1, relief='solid')
            self.image_label.pack(fill=tk.BOTH, expand=True)

            self.pdf_doc = None
            self.pdf_bytes = None
            self.pdf_path = None
            self.last_zaknum = ''
            self.current_page = 0
            self.total_pages = 0
            self.preview_img = None

            self.image_label.bind('<Configure>', self._on_resize)

        def open_order_picker(self):
            OrderPickerWindow(self.root, self.on_order_selected)

        def on_order_selected(self, zaknum):
            self.zak_entry.delete(0, tk.END)
            self.zak_entry.insert(0, zaknum)
            self.zak_entry.focus_set()
            self.set_status(f'Выбран заказ: {zaknum}', 'green')
            self.load()  # Автоматическая загрузка при выборе

        def paste_from_clipboard(self):
            try:
                text = self.root.clipboard_get()
            except tk.TclError:
                text = ''
            if text:
                self.zak_entry.delete(0, tk.END)
                self.zak_entry.insert(0, text.strip())
                self.zak_entry.focus_set()
                self.load()  # Автоматическая загрузка при вставке

        def _on_ctrl_v(self, event=None):
            try:
                text = self.root.clipboard_get()
            except tk.TclError:
                return 'break'
            if text:
                self.zak_entry.delete(0, tk.END)
                self.zak_entry.insert(0, text.strip())
                self.zak_entry.focus_set()
                self.load()  # Автоматическая загрузка при Ctrl+V
            return 'break'

        def set_status(self, text, color='blue'):
            self.status.configure(text=text, foreground=color)
            self.root.update_idletasks()

        def load(self):
            zaknum = self.zak_entry.get().strip()
            if not zaknum:
                messagebox.showwarning('Внимание', 'Введите номер заказа')
                return
            self.last_zaknum = zaknum
            self.set_status('Чтение данных...')

            try:
                recs = fetch_parts(zaknum)
            except Exception as e:
                messagebox.showerror('Ошибка БД', str(e))
                self.set_status('Ошибка БД', 'red')
                return

            if not recs:
                self.set_status('Деталей не найдено', 'red')
                messagebox.showinfo('Нет данных',
                                    f'По заказу {zaknum} деталей не найдено')
                return

            self.set_status(f'Найдено {len(recs)} деталей, генерация PDF...')

            try:
                buf = build_pdf_bytes(zaknum, recs)
            except Exception as e:
                messagebox.showerror('Ошибка генерации', str(e))
                self.set_status('Ошибка генерации', 'red')
                return

            tmp = Path(tempfile.gettempdir()) / f'labels_{zaknum}.pdf'
            tmp.write_bytes(buf.getvalue())
            self.pdf_path = tmp

            try:
                self.open_preview(buf.getvalue())
            except Exception as e:
                messagebox.showerror('Ошибка предпросмотра', str(e))
                return

            groups = sort_for_print(recs)
            summary = ', '.join(f'{n}: {len(l)}' for n, l in groups)
            total = sum(len(l) for _, l in groups)
            filtered_out = len(recs) - total
            self.set_status(
                f'Готово: {summary} — этикеток {total} (отфильтровано {filtered_out}) — {tmp}', 'green')

        def open_preview(self, pdf_bytes):
            if self.pdf_doc:
                try:
                    self.pdf_doc.close()
                except Exception:
                    pass
            self.pdf_bytes = pdf_bytes
            self.pdf_doc = pdfium.PdfDocument(self.pdf_bytes)
            self.total_pages = len(self.pdf_doc)
            self.total_label.config(text=str(self.total_pages))
            self.show_page(0)

        def show_page(self, index):
            if not self.pdf_doc:
                return
            if index < 0:
                index = 0
            if index >= self.total_pages:
                index = self.total_pages - 1
            self.current_page = index
            self.page_var.set(str(index + 1))

            try:
                page = self.pdf_doc[index]
                scale = 3.0
                bitmap = page.render(scale=scale)
                pil = bitmap.to_pil()

                w = self.image_label.winfo_width()
                h = self.image_label.winfo_height()
                if w > 100 and h > 100:
                    pil.thumbnail((w - 4, h - 4), PILImage.LANCZOS)

                self.preview_img = ImageTk.PhotoImage(pil)
                self.image_label.config(image=self.preview_img)
            except Exception as e:
                self.set_status(f'Ошибка отрисовки страницы {index + 1}: {e}', 'red')

        def _on_resize(self, event):
            if self.pdf_doc and self.total_pages > 0:
                self.root.after(200, lambda: self.show_page(self.current_page))

        def prev_page(self):
            if self.current_page > 0:
                self.show_page(self.current_page - 1)

        def next_page(self):
            if self.current_page < self.total_pages - 1:
                self.show_page(self.current_page + 1)

        def goto_page(self):
            try:
                p = int(self.page_var.get()) - 1
                if 0 <= p < self.total_pages:
                    self.show_page(p)
                else:
                    self.page_var.set(str(self.current_page + 1))
            except ValueError:
                self.page_var.set(str(self.current_page + 1))

        def _current_printer(self):
            p = self.printer_var.get().strip()
            return p or None

        def print_single_page(self, page_index):
            if not self.pdf_path or not self.pdf_path.exists():
                messagebox.showinfo('Нет данных', 'Сначала выберите заказ')
                return
            try:
                src = pdfium.PdfDocument(str(self.pdf_path))
                if page_index < 0 or page_index >= len(src):
                    src.close(); return
                new_pdf = pdfium.PdfDocument.new()
                new_pdf.import_pages(src, [page_index])
                zaknum = self.last_zaknum or 'zakaz'
                dst = (Path(tempfile.gettempdir())
                       / f'labels_{zaknum}_p{page_index + 1}.pdf')
                new_pdf.save(str(dst))
                new_pdf.close(); src.close()

                ok = print_to_printer(dst, self._current_printer())
                self.set_status(
                    f'{"Отправлено" if ok else "Открыто"}: стр. {page_index + 1}',
                    'green')
            except Exception as e:
                messagebox.showerror('Ошибка печати', str(e))

        def print_pdf(self):
            if not self.pdf_path or not self.pdf_path.exists():
                messagebox.showinfo('Нет данных', 'Сначала выберите заказ')
                return
            printer = self._current_printer()

            try:
                src = pdfium.PdfDocument(str(self.pdf_path))
                n = len(src)
                if n == 0:
                    src.close()
                    return

                new_pdf = pdfium.PdfDocument.new()
                indices = list(range(n - 1, -1, -1))
                new_pdf.import_pages(src, indices)

                zaknum = self.last_zaknum or 'zakaz'
                dst = (Path(tempfile.gettempdir())
                       / f'labels_{zaknum}_reversed.pdf')
                new_pdf.save(str(dst))
                new_pdf.close()
                src.close()

                ok = print_to_printer(dst, printer)
                self.set_status(
                    f'{"Отправлено" if ok else "Открыто"} (в обратном порядке)'
                    + (f' ({printer})' if printer else ''), 'green')
            except Exception as e:
                messagebox.showerror('Ошибка печати', f'Не удалось развернуть PDF: {e}')

        def save_pdf(self):
            if not self.pdf_path or not self.pdf_path.exists():
                messagebox.showinfo('Нет данных', 'Сначала выберите заказ')
                return
            zaknum = self.last_zaknum or 'zakaz'
            dst = filedialog.asksaveasfilename(
                defaultextension='.pdf',
                initialfile=f'labels_{zaknum}.pdf',
                filetypes=[('PDF', '*.pdf')])
            if not dst:
                return
            Path(dst).write_bytes(self.pdf_path.read_bytes())
            self.set_status(f'Сохранено: {dst}', 'green')

        def clear(self):
            if self.pdf_doc:
                try:
                    self.pdf_doc.close()
                except Exception:
                    pass
            self.pdf_doc = None
            self.pdf_bytes = None
            self.pdf_path = None
            self.last_zaknum = ''
            self.current_page = 0
            self.total_pages = 0
            self.page_var.set('1')
            self.total_label.config(text='0')
            self.image_label.config(image='')
            self.preview_img = None
            self.zak_entry.delete(0, tk.END)
            self.set_status(f'{VERSION}  —  введите номер заказа')

    root = tk.Tk()
    App(root)
    root.mainloop()


# ================= MAIN =================
def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('--zaknum', default=None)
    ap.add_argument('--out', default=None)
    ap.add_argument('--img', default=None)
    args = ap.parse_args()

    global IMG_DIR
    if args.img:
        IMG_DIR = Path(args.img)

    if args.zaknum:
        out = args.out or f'labels_{args.zaknum}.pdf'
        run_cli(args.zaknum, out)
    else:
        run_gui()


if __name__ == '__main__':
    main()
# -*- coding: utf-8 -*-
# ============================================================
#  NAKL3.6  —  print_labels.py
#  Термоэтикетки 100 × 35 мм, 203 dpi.
#
#  Постраничный просмотр: 1 страница на экране.
#  Навигация: ◀ / ▶ / ввод номера страницы.
#  Нет ограничения по количеству страниц в PDF.
# ============================================================

import io
import os
import sys
import tempfile
import argparse
from pathlib import Path

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

VERSION = 'NAKL3.6'


# ================= КОНФИГ =================
DB_HOST = '192.168.0.10'
DB_USER = 'workshop'
DB_PASSWORD = 'w0rK5h0p'
DB_CHARSET = 'utf8mb4'
SRC_DB = 'CCALC'
DST_DB = 'workshop'

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
DETAL_Y = 1.2 * mm

# ---- Начала текста ----
LEFT_X = 1.5 * mm
TXT_X  = QR_X + QR_SIZE + 2 * mm

# ---- Картинка ----
IMG_RIGHT_EDGE = 105 * mm
IMG_MAX_H      = 45 * mm
IMG_V_CENTER   = True

# ---- Фурнитура ----
FRN_MAX_W      = 45 * mm
FRN_MAX_LINES  = 3

# ---- Примечание ----
PRIM_MAX_LINES = 2

# ---- Вертикальные базовые линии строк ----
Y_L1  = 31.0 * mm
Y_L2  = 25.5 * mm
Y_L3  = 20.0 * mm
Y_L4  = 15.5 * mm
Y_L4B = 13.0 * mm
Y_L4C = 10.5 * mm
Y_L7  = 8.0 * mm
Y_L8  = 5.0 * mm
Y_L8B = 2.5 * mm

# ---- Шрифты ----
FONT_NAME = 'Arial'
FONT_BOLD = 'Arial-Bold'
ARIAL      = r'C:\Windows\Fonts\arial.ttf'
ARIAL_BOLD = r'C:\Windows\Fonts\arialbd.ttf'

SZ_ROLE    = 12
SZ_MATNAME = 10
SZ_COLOR   = 8
SZ_ZAKNUM  = 18
SZ_LENGTH  = 14
SZ_ORIENT  = 13
SZ_ANGLES  = 13
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


# ================= ДАННЫЕ =================
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

    cur.execute(f"""
        SELECT ZAKNUM, KONID, MIN(MATCOLOR) AS SEAL_COLOR
        FROM `{DST_DB}`.`CCALC_TBL`
        WHERE ZAKNUM = %s AND MATNAME LIKE %s
        GROUP BY ZAKNUM, KONID
    """, (zaknum, '%Уплотнение%'))
    seals = {(r['ZAKNUM'], r['KONID']): r['SEAL_COLOR'] for r in cur.fetchall()}

    for p in parts:
        key = (p['ZAKNUM'], p['KONID'])
        f = frn.get(key)
        p['FRN_NAME']     = f['NAME']     if f else None
        p['FRN_TYPENAME'] = f['TYPENAME'] if f else None
        p['FRN_NAPRNAME'] = f['NAPRNAME'] if f else None
        p['SEAL_COLOR']   = seals.get(key)

    cur.close(); conn.close()
    return parts


# ================= СОРТИРОВКА =================
def _int_or_inf(v):
    try:
        return int(v) if v is not None else 10**9
    except (TypeError, ValueError):
        return 10**9


def sort_for_print(parts):
    def gri_of(p):
        return str(p.get('GRI') or '').strip().lower()

    def key_rs(p):
        return (_int_or_inf(p.get('KONID')),
                _int_or_inf(p.get('OBJID')),
                ORIENT_ORDER.get((p.get('ORIENT') or '').strip(), 99),
                str(p.get('DETAL') or ''))

    rams    = sorted([p for p in parts if gri_of(p) == 'r'], key=key_rs)
    stvorki = sorted([p for p in parts if gri_of(p) == 's'], key=key_rs)
    imposts = [p for p in parts if gri_of(p) == '3']
    others  = [p for p in parts if gri_of(p) not in ('r', 's', '3')]

    groups = [('РАМЫ', rams), ('СТВОРКИ', stvorki), ('ИМПОСТА', imposts + others)]
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
            if IMG_V_CENTER:
                img_y = (LABEL_H - h) / 2
            else:
                img_y = 0
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

    # ================= ТЕКСТ =================
    LW_EDGE = LABEL_W - 1.5 * mm

    # === Строка 1 ===
    role = build_role(rec)
    c.setFont(FONT_BOLD, SZ_ROLE)
    c.drawString(LEFT_X, Y_L1, role)
    role_w = c.stringWidth(role, FONT_BOLD, SZ_ROLE)

    c.setFont(FONT_BOLD, SZ_MATNAME)
    gap_role = c.stringWidth('  ', FONT_BOLD, SZ_MATNAME)
    cur_x = LEFT_X + role_w + gap_role

    matname = (rec.get('MATNAME') or '').strip()
    color_txt = clean_color(rec.get('MATCOLOR'))
    gap_color = c.stringWidth(' ', FONT_BOLD, SZ_MATNAME)

    if color_txt:
        c.setFont(FONT_BOLD, SZ_COLOR)
        color_w = c.stringWidth(color_txt, FONT_BOLD, SZ_COLOR)
    else:
        color_w = 0

    if color_txt:
        mat_w = LW_EDGE - cur_x - gap_color - color_w
    else:
        mat_w = LW_EDGE - cur_x

    if matname:
        c.setFont(FONT_BOLD, SZ_MATNAME)
        mat_draw = fit_text(c, matname, FONT_BOLD, SZ_MATNAME, mat_w)
        c.drawString(cur_x, Y_L1, mat_draw)
        cur_x += c.stringWidth(mat_draw, FONT_BOLD, SZ_MATNAME)

    if color_txt:
        cur_x += gap_color
        c.setFont(FONT_BOLD, SZ_COLOR)
        c.drawString(cur_x, Y_L1, color_txt)

    # === Строка 2 ===
    zak = str(rec.get('ZAKNUM') or '')
    c.setFont(FONT_BOLD, SZ_ZAKNUM)
    c.drawString(LEFT_X, Y_L2, zak)
    zak_w = c.stringWidth(zak, FONT_BOLD, SZ_ZAKNUM)
    gap2 = c.stringWidth('  ', FONT_BOLD, SZ_ZAKNUM)

    c.setFont(FONT_BOLD, SZ_LENGTH)
    c.drawString(LEFT_X + zak_w + gap2, Y_L2,
                 f"L:{int(rec.get('LENGTH') or 0)}")

    # === Строка 3 ===
    orient = build_orient(rec) or ''
    c.setFont(FONT_BOLD, SZ_ORIENT)
    c.drawString(TXT_X, Y_L3, orient)
    orient_w = c.stringWidth(orient, FONT_BOLD, SZ_ORIENT)
    gap3 = c.stringWidth('   ', FONT_BOLD, SZ_ORIENT)

    al = int(rec.get('ANGLE_L') or 0)
    ar = int(rec.get('ANGLE_R') or 0)
    c.setFont(FONT_BOLD, SZ_ANGLES)
    c.drawString(TXT_X + orient_w + gap3, Y_L3, f"{al}° {ar}°")

    # === Строки 4-6: фурнитура ===
    name = (rec.get('FRN_NAME') or '').strip()
    if name:
        c.setFont(FONT_BOLD, SZ_FRN)
        lines = wrap_text(c, name, FONT_BOLD, SZ_FRN,
                          FRN_MAX_W, max_lines=FRN_MAX_LINES)
        y_positions = [Y_L4, Y_L4B, Y_L4C]
        for i, line in enumerate(lines):
            if i < len(y_positions):
                c.drawString(TXT_X, y_positions[i], line)

    # === Строка 7: уплотнение ===
    seal = clean_color(rec.get('SEAL_COLOR'))
    if seal:
        c.setFont(FONT_BOLD, SZ_SEAL)
        c.drawString(TXT_X, Y_L7, f"Упл. {seal}")

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
    print(f'   Картинка : высота {IMG_MAX_H/mm:.0f} мм')
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

    class App:
        def __init__(self, root):
            self.root = root
            root.title(f'{VERSION} — Термоэтикетки CCalc')
            root.geometry('1100x850')
            root.minsize(800, 600)

            # --- Верхняя панель ---
            top = ttk.Frame(root, padding=8)
            top.pack(side=tk.TOP, fill=tk.X)

            ttk.Label(top, text='Номер заказа:',
                      font=('Arial', 11)).pack(side=tk.LEFT)

            self.zak_entry = ttk.Entry(top, width=24, font=('Arial', 12))
            self.zak_entry.pack(side=tk.LEFT, padx=6)
            self.zak_entry.bind('<Return>', lambda e: self.load())
            self.zak_entry.focus_set()

            ttk.Button(top, text='📋 Из буфера',
                       command=self.paste_from_clipboard).pack(side=tk.LEFT, padx=3)
            ttk.Button(top, text='Показать',
                       command=self.load).pack(side=tk.LEFT, padx=3)
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
                values=printers, width=30, state='readonly')
            if printers:
                self.printer_cb.current(0)
            self.printer_cb.pack(side=tk.LEFT)

            ttk.Button(top, text='🖨 Печать',
                       command=self.print_pdf).pack(side=tk.LEFT, padx=6)

            # --- Статус ---
            self.status = ttk.Label(root, text=f'{VERSION}  —  введите номер заказа',
                                    foreground='blue', font=('Arial', 10))
            self.status.pack(side=tk.TOP, fill=tk.X, padx=10)

            root.bind('<Control-v>', self._on_ctrl_v)
            root.bind('<Control-V>', self._on_ctrl_v)

            # --- ПАНЕЛЬ НАВИГАЦИИ (снизу) ---
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

            # --- ПАНЕЛЬ ПРЕДПРОСМОТРА (по центру) ---
            body = ttk.Frame(root)
            body.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=6)

            self.image_label = tk.Label(body, bg='#d8d8d8',
                                         borderwidth=1, relief='solid')
            self.image_label.pack(fill=tk.BOTH, expand=True)

            # --- Переменные состояния ---
            self.pdf_doc = None
            self.pdf_bytes = None
            self.pdf_path = None
            self.last_zaknum = ''
            self.current_page = 0
            self.total_pages = 0
            self.preview_img = None

            # отслеживание изменения размера окна для перерисовки текущей страницы
            self.image_label.bind('<Configure>', self._on_resize)

        # ---------- буфер обмена ----------
        def paste_from_clipboard(self):
            try:
                text = self.root.clipboard_get()
            except tk.TclError:
                text = ''
            if text:
                self.zak_entry.delete(0, tk.END)
                self.zak_entry.insert(0, text.strip())
            self.zak_entry.focus_set()

        def _on_ctrl_v(self, event=None):
            try:
                text = self.root.clipboard_get()
            except tk.TclError:
                return 'break'
            if text:
                self.zak_entry.delete(0, tk.END)
                self.zak_entry.insert(0, text.strip())
                self.zak_entry.focus_set()
            return 'break'

        def set_status(self, text, color='blue'):
            self.status.configure(text=text, foreground=color)
            self.root.update_idletasks()

        # ---------- загрузка ----------
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

            # ПЕРЕХОДИМ В РЕЖИМ ПРОСМОТРА ОДНОЙ СТРАНИЦЫ
            try:
                self.open_preview(buf.getvalue())
            except Exception as e:
                messagebox.showerror('Ошибка предпросмотра', str(e))
                return

            groups = sort_for_print(recs)
            summary = ', '.join(f'{n}: {len(l)}' for n, l in groups)
            total = sum(len(l) for _, l in groups)
            self.set_status(
                f'Готово: {summary} — этикеток {total} — {tmp}', 'green')

        # ---------- предпросмотр ОДНОЙ страницы ----------
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
                # рендерим с высокой детализацией
                scale = 3.0
                bitmap = page.render(scale=scale)
                pil = bitmap.to_pil()

                # вписываем в текущий размер области предпросмотра
                w = self.image_label.winfo_width()
                h = self.image_label.winfo_height()
                if w > 100 and h > 100:
                    pil.thumbnail((w - 4, h - 4), PILImage.LANCZOS)

                self.preview_img = ImageTk.PhotoImage(pil)
                self.image_label.config(image=self.preview_img)
            except Exception as e:
                self.set_status(f'Ошибка отрисовки страницы {index + 1}: {e}', 'red')

        def _on_resize(self, event):
            # при изменении размера окна — перерисовать текущую страницу
            if self.pdf_doc and self.total_pages > 0:
                # небольшая задержка, чтобы окно успело перерисоваться
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

        # ---------- печать ----------
        def _current_printer(self):
            p = self.printer_var.get().strip()
            return p or None

        def print_single_page(self, page_index):
            if not self.pdf_path or not self.pdf_path.exists():
                messagebox.showinfo('Нет данных', 'Сначала нажмите «Показать»')
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
                messagebox.showinfo('Нет данных', 'Сначала нажмите «Показать»')
                return
            printer = self._current_printer()
            ok = print_to_printer(self.pdf_path, printer)
            self.set_status(
                f'{"Отправлено" if ok else "Открыто"}'
                + (f' ({printer})' if printer else ''), 'green')

        # ---------- сохранение / очистка ----------
        def save_pdf(self):
            if not self.pdf_path or not self.pdf_path.exists():
                messagebox.showinfo('Нет данных', 'Сначала нажмите «Показать»')
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
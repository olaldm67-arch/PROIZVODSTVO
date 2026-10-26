# -*- coding: utf-8 -*-
# ============================================================
#  NAKL2.1  —  print_labels.py
#  Наклейки для деталей заказа.
#
#  Этикетка 205 × 35 мм. A4 вертикально.
#  Верхнее поле 12 мм (непечатаемая зона), нижнее 5 мм → 8 шт. на лист.
#
#  Непропорциональные части этикетки:
#     [1: 70 мм текст] [2: 55 мм текст] [3: 50 мм картинка] [4: 30 мм QR]
#
#  Картинка: 40 мм по высоте жёстко, ширина по пропорциям,
#            может выходить за границы этикетки.
#
#  Сортировка РАМ / СТВОРОК: KONID → OBJID → ORIENT (В, Н, Л, П)
#  Порядок групп: РАМЫ → СТВОРКИ → ИМПОСТА (каждая с нового листа).
#
#  Новое в NAKL2.1:
#     • вставка номера заказа из буфера обмена (кнопка "Из буфера");
#     • печать отдельной страницы предпросмотра
#       (кнопка "Печать стр. N" под каждой страницей).
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

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import black, grey
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader


ImageFile.LOAD_TRUNCATED_IMAGES = True

VERSION = 'NAKL2.1'


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

# ---- Размеры этикетки ----
LABEL_W = 205 * mm
LABEL_H = 35 * mm

# Поля страницы A4
MARGIN_TOP    = 12 * mm
MARGIN_BOTTOM = 5 * mm
MARGIN_X      = (A4[0] - LABEL_W) / 2

LABELS_PER_PAGE = int((A4[1] - MARGIN_TOP - MARGIN_BOTTOM) // LABEL_H)

IMG_H = 40 * mm

# ---- Ширина частей этикетки ----
PART1_W = 70 * mm
PART2_W = 55 * mm
PART3_W = 50 * mm
PART4_W = 30 * mm
PART1_X = 0
PART2_X = PART1_W
PART3_X = PART1_W + PART2_W
PART4_X = PART1_W + PART2_W + PART3_W

# ---- Шрифты ----
FONT_NAME = 'Arial'
FONT_BOLD = 'Arial-Bold'
ARIAL      = r'C:\Windows\Fonts\arial.ttf'
ARIAL_BOLD = r'C:\Windows\Fonts\arialbd.ttf'

SZ_ROLE    = 17
SZ_ZAKNUM  = 32
SZ_NAME    = 10
SZ_ORIENT  = 10
SZ_MATNAME = 10
SZ_COLOR   = 9
SZ_LENGTH  = 14
SZ_ANGLES  = 10
SZ_SEAL    = 9
SZ_PRIM    = 9
SZ_QR_VAL  = 8
SZ_NOFILE  = 8

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
        kon = _int_or_inf(p.get('KONID'))
        obj = _int_or_inf(p.get('OBJID'))
        o = (p.get('ORIENT') or '').strip()
        return (kon, obj, ORIENT_ORDER.get(o, 99), str(p.get('DETAL') or ''))

    rams    = sorted([p for p in parts if gri_of(p) == 'r'], key=key_rs)
    stvorki = sorted([p for p in parts if gri_of(p) == 's'], key=key_rs)
    imposts = [p for p in parts if gri_of(p) == '3']
    others  = [p for p in parts if gri_of(p) not in ('r', 's', '3')]

    groups = [
        ('РАМЫ',    rams),
        ('СТВОРКИ', stvorki),
        ('ИМПОСТА', imposts + others),
    ]
    return [(name, lst) for name, lst in groups if lst]


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
        img.save(buf, format='JPEG', quality=85)
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
        return f"{t} - {n}"
    return t or n


def split_two_lines(text, max_len=65):
    text = (text or '').strip()
    if len(text) <= max_len:
        return [text]
    cut = text.rfind(' ', 0, max_len)
    if cut <= 0:
        cut = max_len
    return [text[:cut].rstrip(), text[cut:].lstrip()]


# ================= РИСОВАНИЕ ЭТИКЕТКИ =================
def draw_label(c, x, y, rec):
    c.setStrokeColor(black)
    c.setLineWidth(0.4)
    c.rect(x, y, LABEL_W, LABEL_H)

    pad = 1.5 * mm
    top = y + LABEL_H

    p1x = x + PART1_X + pad
    p2x = x + PART2_X + pad
    p4x = x + PART4_X + pad

    cx2 = x + PART2_X + PART2_W / 2
    cx3 = x + PART3_X + PART3_W / 2

    # ---------- Часть 1: 70 мм ----------
    c.setFont(FONT_BOLD, SZ_ROLE)
    c.drawString(p1x, top - 5 * mm, build_role(rec))

    c.setFont(FONT_BOLD, SZ_ZAKNUM)
    c.drawString(p1x, top - 15 * mm, str(rec.get('ZAKNUM') or ''))

    name = (rec.get('FRN_NAME') or '').strip()
    y_name = top - 20 * mm
    if name:
        c.setFont(FONT_BOLD, SZ_NAME)
        line1 = name[:40]
        line2 = name[40:80]
        c.drawString(p1x, y_name, line1)
        if line2:
            y_name -= 3.5 * mm
            c.drawString(p1x, y_name, line2)

    orient = build_orient(rec)
    if orient:
        c.setFont(FONT_BOLD, SZ_ORIENT)
        c.drawString(p1x, y_name - 4 * mm, orient[:44])

    # ---------- Часть 2: 55 мм ----------
    c.setFont(FONT_BOLD, SZ_MATNAME)
    c.drawString(p2x, top - 4.5 * mm, (rec.get('MATNAME') or '')[:28])

    color = clean_color(rec.get('MATCOLOR'))
    if color:
        c.setFont(FONT_BOLD, SZ_COLOR)
        c.drawString(p2x, top - 9 * mm, f"Цвет: {color}")

    c.setFont(FONT_BOLD, SZ_LENGTH)
    c.drawCentredString(cx2, top - 17 * mm,
                        f"L:{int(rec.get('LENGTH') or 0)}")

    al = int(rec.get('ANGLE_L') or 0)
    ar = int(rec.get('ANGLE_R') or 0)
    c.setFont(FONT_BOLD, SZ_ANGLES)
    c.drawCentredString(cx2, top - 22 * mm, f"{al}° {ar}°")

    seal = clean_color(rec.get('SEAL_COLOR'))
    if seal:
        c.setFont(FONT_BOLD, SZ_SEAL)
        c.drawString(p2x, top - 27 * mm, f"Упл. {seal}")

    # ---------- PRIMPROIZV — 2 строки внизу частей 1+2 ----------
    prim = (rec.get('PRIMPROIZV') or '').strip()
    if prim:
        c.setFont(FONT_BOLD, SZ_PRIM)
        lines = split_two_lines(prim, max_len=65)
        if len(lines) == 1:
            c.drawString(x + PART1_X + pad, y + 3 * mm, lines[0])
        else:
            c.drawString(x + PART1_X + pad, y + 5 * mm, lines[0])
            c.drawString(x + PART1_X + pad, y + 1.5 * mm, lines[1])

    # ---------- Часть 3: 50 мм, картинка ----------
    reader = get_imagereader(rec.get('ZAKNUM'), rec.get('KONID'))
    if reader is not None:
        try:
            iw0, ih0 = reader.getSize()
            img_h = IMG_H
            img_w = img_h * (iw0 / ih0)
            img_x = cx3 - img_w / 2
            img_y = y + (LABEL_H - img_h) / 2
            c.drawImage(reader, img_x, img_y,
                        width=img_w, height=img_h,
                        mask='auto')
        except Exception:
            c.setStrokeColor(grey)
            c.rect(cx3 - 20 * mm, y + 0.5 * mm, 40 * mm, LABEL_H - 1 * mm)
    else:
        c.setStrokeColor(grey)
        c.rect(cx3 - 20 * mm, y + 0.5 * mm, 40 * mm, LABEL_H - 1 * mm)
        c.setFont(FONT_BOLD, SZ_NOFILE)
        c.setFillColor(grey)
        c.drawCentredString(cx3, y + LABEL_H / 2, 'нет файла')
        c.setFillColor(black)

    # ---------- Часть 4: 30 мм, QR ----------
    bc_val = str(rec.get('DETAL') or '')
    qr_reader = get_qr_reader(bc_val)

    qr_size = LABEL_H - 8 * mm
    qr_x = p4x
    qr_y = y + (LABEL_H - qr_size) / 2

    if qr_reader is not None:
        try:
            c.drawImage(qr_reader, qr_x, qr_y,
                        width=qr_size, height=qr_size,
                        preserveAspectRatio=True, mask='auto')
        except Exception:
            c.setStrokeColor(grey)
            c.rect(qr_x, qr_y, qr_size, qr_size)
    else:
        c.setStrokeColor(grey)
        c.rect(qr_x, qr_y, qr_size, qr_size)

    if bc_val:
        c.setFont(FONT_BOLD, SZ_QR_VAL)
        c.setFillColor(black)
        c.drawString(qr_x, y + 1.5 * mm, bc_val)


# ================= PDF =================
def build_pdf_to(target, zaknum, parts):
    register_fonts()
    c = Canvas(target, pagesize=A4)
    c.setTitle(f'{VERSION} — Наклейки {zaknum}')

    for group_name, items in sort_for_print(parts):
        idx = 0
        while idx < len(items):
            for r in range(LABELS_PER_PAGE):
                if idx >= len(items):
                    break
                x = MARGIN_X
                y = A4[1] - MARGIN_TOP - (r + 1) * LABEL_H
                draw_label(c, x, y, items[idx])
                idx += 1
            c.showPage()
    c.save()


def build_pdf_file(zaknum, out_path, parts):
    build_pdf_to(str(out_path), zaknum, parts)


def build_pdf_bytes(zaknum, parts):
    buf = io.BytesIO()
    build_pdf_to(buf, zaknum, parts)
    buf.seek(0)
    return buf


# ================= CLI =================
def run_cli(zaknum, out_path):
    print(f'📖 {VERSION}  заказ {zaknum}')
    if IMG_DIR:
        print(f'   Картинки : {IMG_DIR}')
    print(f'   Этикетка : {LABEL_W/mm:.0f} × {LABEL_H/mm:.0f} мм')
    print(f'   Части    : {PART1_W/mm:.0f} / {PART2_W/mm:.0f} / '
          f'{PART3_W/mm:.0f} / {PART4_W/mm:.0f} мм')
    print(f'   Поля     : верх {MARGIN_TOP/mm:.0f} мм / низ {MARGIN_BOTTOM/mm:.0f} мм')
    print(f'   На листе : {LABELS_PER_PAGE} шт.')
    recs = fetch_parts(zaknum)
    if not recs:
        print('❌ Деталей не найдено'); sys.exit(1)
    pages = sum((len(l) + LABELS_PER_PAGE - 1) // LABELS_PER_PAGE
                for _, l in sort_for_print(recs))
    print(f'   Деталей  : {len(recs)}')
    print(f'   Листов   : {pages}')
    print(f'📄 Пишу {out_path}')
    build_pdf_file(zaknum, out_path, recs)
    print(f'✅ Готово: {out_path}')


# ================= GUI =================
def run_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    from PIL import ImageTk
    import pypdfium2 as pdfium

    class App:
        def __init__(self, root):
            self.root = root
            root.title(f'{VERSION} — Наклейки CCalc')
            root.geometry('1150x820')
            root.minsize(800, 600)

            top = ttk.Frame(root, padding=8)
            top.pack(side=tk.TOP, fill=tk.X)

            ttk.Label(top, text='Номер заказа:',
                      font=('Arial', 11)).pack(side=tk.LEFT)

            self.zak_entry = ttk.Entry(top, width=28, font=('Arial', 12))
            self.zak_entry.pack(side=tk.LEFT, padx=6)
            self.zak_entry.bind('<Return>', lambda e: self.load())
            self.zak_entry.focus_set()

            ttk.Button(top, text='📋 Из буфера',
                       command=self.paste_from_clipboard).pack(side=tk.LEFT, padx=3)
            ttk.Button(top, text='Показать',
                       command=self.load).pack(side=tk.LEFT, padx=3)
            ttk.Button(top, text='Сохранить PDF',
                       command=self.save_pdf).pack(side=tk.LEFT, padx=3)
            ttk.Button(top, text='Печать всего',
                       command=self.print_pdf).pack(side=tk.LEFT, padx=3)
            ttk.Button(top, text='Очистить',
                       command=self.clear).pack(side=tk.LEFT, padx=3)

            self.status = ttk.Label(top, text=f'{VERSION}  —  введите номер заказа',
                                    foreground='blue', font=('Arial', 10))
            self.status.pack(side=tk.RIGHT)

            # Ctrl+V работает и в Entry, но сделаем глобально:
            root.bind('<Control-v>', self._on_ctrl_v)
            root.bind('<Control-V>', self._on_ctrl_v)

            body = ttk.Frame(root)
            body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            self.canvas = tk.Canvas(body, bg='#d8d8d8', highlightthickness=0)
            vbar = ttk.Scrollbar(body, orient='vertical', command=self.canvas.yview)
            self.canvas.configure(yscrollcommand=vbar.set)
            vbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            self.inner = ttk.Frame(self.canvas)
            self.canvas.create_window((0, 0), window=self.inner, anchor='nw')
            self.inner.bind(
                '<Configure>',
                lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))

            self.canvas.bind_all(
                '<MouseWheel>',
                lambda e: self.canvas.yview_scroll(int(-e.delta / 120), 'units'))

            self.pdf_path = None
            self.preview_images = []
            self.last_zaknum = ''
            self.pages_count = 0

        def set_status(self, text, color='blue'):
            self.status.configure(text=text, foreground=color)
            self.root.update_idletasks()

        # ---- буфер обмена ----
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
            # если фокус в поле — вставляем туда, иначе просто игнорируем
            try:
                text = self.root.clipboard_get()
            except tk.TclError:
                return 'break'
            if text:
                if self.root.focus_get() == self.zak_entry:
                    self.zak_entry.delete(0, tk.END)
                    self.zak_entry.insert(0, text.strip())
                else:
                    self.zak_entry.delete(0, tk.END)
                    self.zak_entry.insert(0, text.strip())
                    self.zak_entry.focus_set()
            return 'break'

        # ---- загрузка ----
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

            self.render_preview(buf)

            groups = sort_for_print(recs)
            summary = ', '.join(f'{n}: {len(l)}' for n, l in groups)
            total_pages = sum((len(l) + LABELS_PER_PAGE - 1) // LABELS_PER_PAGE
                              for _, l in groups)
            self.pages_count = total_pages
            self.set_status(
                f'Готово: {summary} — {total_pages} стр. — {tmp}', 'green')

        # ---- предпросмотр ----
        def render_preview(self, pdf_bytes):
            for w in self.inner.winfo_children():
                w.destroy()
            self.preview_images = []

            pdf = pdfium.PdfDocument(pdf_bytes.getvalue())
            total = len(pdf)
            scale = 1.6
            for i in range(total):
                page = pdf[i]
                bitmap = page.render(scale=scale)
                pil = bitmap.to_pil()
                tk_img = ImageTk.PhotoImage(pil)
                self.preview_images.append(tk_img)

                # фрейм страницы с рамкой
                fr = ttk.Frame(self.inner, borderwidth=1, relief='solid')
                fr.pack(padx=12, pady=6)

                ttk.Label(fr, image=tk_img).pack()

                bar = ttk.Frame(fr)
                bar.pack(fill=tk.X, pady=2, padx=4)
                ttk.Label(bar, text=f'стр. {i + 1} из {total}',
                          font=('Arial', 10)).pack(side=tk.LEFT)
                ttk.Button(bar, text=f'🖨 Печать стр. {i + 1}',
                           command=lambda idx=i: self.print_single_page(idx)
                           ).pack(side=tk.RIGHT)
            pdf.close()

        # ---- печать ----
        def print_single_page(self, page_index):
            if not self.pdf_path or not self.pdf_path.exists():
                messagebox.showinfo('Нет данных', 'Сначала нажмите «Показать»')
                return
            try:
                src = pdfium.PdfDocument(str(self.pdf_path))
                n = len(src)
                if page_index < 0 or page_index >= n:
                    messagebox.showerror(
                        'Ошибка',
                        f'Страница {page_index + 1} не существует (всего {n})')
                    src.close()
                    return

                new_pdf = pdfium.PdfDocument.new()
                new_pdf.import_pages(src, [page_index])

                zaknum = self.last_zaknum or 'zakaz'
                dst = (Path(tempfile.gettempdir())
                       / f'labels_{zaknum}_p{page_index + 1}.pdf')
                new_pdf.save(str(dst))
                new_pdf.close()
                src.close()

                os.startfile(str(dst), 'print')
                self.set_status(
                    f'Печать стр. {page_index + 1} — {dst}', 'green')
            except Exception as e:
                messagebox.showerror('Ошибка печати', str(e))

        def print_pdf(self):
            if not self.pdf_path or not self.pdf_path.exists():
                messagebox.showinfo('Нет данных', 'Сначала нажмите «Показать»')
                return
            try:
                os.startfile(str(self.pdf_path), 'print')
                self.set_status('Отправлено на печать (весь документ)', 'green')
            except Exception as e:
                messagebox.showerror('Ошибка печати', str(e))

        # ---- сохранение ----
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

        # ---- очистка ----
        def clear(self):
            for w in self.inner.winfo_children():
                w.destroy()
            self.preview_images = []
            self.pdf_path = None
            self.last_zaknum = ''
            self.pages_count = 0
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
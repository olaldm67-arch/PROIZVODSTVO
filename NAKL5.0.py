# -*- coding: utf-8 -*-
# ============================================================
#  NAKL4.5  —  print_labels.py
#  Термоэтикетки 100 × 35 мм + миграция DBF + отчёты.
#
#  Встроено из CCalc_Sync:
#    • Миграция новых заказов DBF -> CCALC (BASE*)
#    • Генерация отчётов в workshop (CCALC, CCALC_FRN, CCALC_SP, CCALC_TBL)
#    • Отдельное окно лога, запуск в фоне (не блокирует UI).
# ============================================================

import io
import os
import sys
import time
import queue
import threading
import logging
import tempfile
import argparse
from pathlib import Path
from datetime import datetime

import mysql.connector
import qrcode
from PIL import Image, ImageFile
from dbfread import DBF

from reportlab.lib.units import mm
from reportlab.lib.colors import black, grey
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader


ImageFile.LOAD_TRUNCATED_IMAGES = True

VERSION = 'NAKL4.5'
DEFAULT_PRINTER = 'HPRT LPG4'


# ================= КОНФИГ =================
DB_HOST = '192.168.0.10'
DB_USER = 'workshop'
DB_PASSWORD = 'w0rK5h0p'
DB_CHARSET = 'utf8mb4'
SRC_DB = 'CCALC'
DST_DB = 'workshop'

MDB_PATH = r'E:\ACCESS\BACKUP\15.10.2026\ACCESS\2010.mdb'
MDB_TABLE = 'лист1'
MDB_CACHE_TTL = 15 * 60

DBF_PATH = Path(r'\\sp-roshan-03\Ccalc6.80\Data')
MIGRATE_TABLES = ['BASE1', 'BASE3', 'BASE4', 'BASE6',
                  'BASE7', 'BASE8', 'BASE10', 'BASE14']
BATCH_SIZE = 500

_mdb_cache = None
_mdb_cache_time = 0

if sys.platform == 'linux':
    IMG_DIR = Path(r'/mnt/smb/ПВХ/JPG')
elif sys.platform == 'win32':
    IMG_DIR = Path(r'\\Synas\work\ПВХ\JPG')
else:
    IMG_DIR = None

# ---- Этикетка ----
LABEL_W, LABEL_H = 100 * mm, 35 * mm
PAGE_SIZE = (LABEL_W, LABEL_H)

QR_SIZE, QR_X, QR_Y, DETAL_Y = 22 * mm, 1.5 * mm, 3.5 * mm, 2.2 * mm
LEFT_X = 1.5 * mm
TXT_X = QR_X + QR_SIZE + 2 * mm

IMG_RIGHT_EDGE = 115 * mm
IMG_MAX_H = 45 * mm
IMG_V_CENTER = True

FRN_MAX_W, FRN_MAX_LINES = 45 * mm, 3
PRIM_MAX_LINES = 2

Y_L1, Y_L2, Y_L3 = 29.0 * mm, 25.5 * mm, 20.0 * mm
Y_L4, Y_L4B, Y_L4C = 16.5 * mm, 13.5 * mm, 10.5 * mm
Y_L7, Y_L8, Y_L8B = 8.0 * mm, 5.0 * mm, 2.5 * mm

FONT_NAME = 'Arial'
FONT_BOLD = 'Arial-Bold'
ARIAL = r'C:\Windows\Fonts\arial.ttf'
ARIAL_BOLD = r'C:\Windows\Fonts\arialbd.ttf'

SZ_ZAKNUM, SZ_ROLE, SZ_MATNAME, SZ_COLOR = 24, 18, 12, 12
SZ_LENGTH, SZ_ANGLES = 14, 14
SZ_FRN, SZ_SEAL, SZ_PRIM = 8, 8, 8
SZ_QR_VAL, SZ_NOFILE = 7, 7

ORIENT_ORDER = {'В': 0, 'Н': 1, 'Л': 2, 'П': 3}

# ================= SQL ОТЧЁТОВ =================
SQL_CCALC = f"""
INSERT INTO `{DST_DB}`.`CCALC`
    (ZAKNUM, KONNAME, NUMIZDELIE, DETAL, PRIMPROIZV, BCI, PROFILTYPE,
     KONID, GRI, RAMAID, MATNAME, MATART, ELEMID, OBJID, MATCOLOR,
     LENGTH, ZAKID, ANGLE_L, ANGLE_R, STAR, ORIENT, ORIENT2,
     IMPOST, RUCHKA, ZAMOK, DUGA)
SELECT
    b7.ZAKNUM, b7.KONNAME, b7.NUMIZDELIE,
    b7.BARCODE AS DETAL, b4.PRIMPROIZV, b7.BARCODEIZD AS BCI,
    b7.PROFILTYPE, b7.KONID, n.`Вариант` AS GRI,
    b7.RAMAID, b7.MATNAME, b7.MATART, b7.ELEMID, b7.OBJID,
    b7.MATCOLOR, b7.LENGTH, b7.ZAKID,
    b7.ANGLE_L, b7.ANGLE_R, b7.STAR,
    b7.ORIENT, b7.ORIENT2, b7.IMPOST, b7.RUCHKA, b7.ZAMOK, b7.DUGA
FROM `{SRC_DB}`.`BASE7` b7
LEFT JOIN `{SRC_DB}`.`NOM`   n  ON n.MATART  = b7.MATART
LEFT JOIN `{SRC_DB}`.`BASE4` b4 ON b4.ZAKNUM = b7.ZAKNUM
                               AND b4.KONNAME = b7.KONNAME
LEFT JOIN `{DST_DB}`.`CCALC` c  ON c.DETAL   = b7.BARCODE
WHERE b7.PROFILTYPE <= 3 AND c.DETAL IS NULL
"""

SQL_FRN = f"""
INSERT INTO `{DST_DB}`.`CCALC_FRN`
    (ZAKNUM, ZAKID, ADRESS, KONNUM, KONID, KONNAME, NAME, NAMEID,
     L, H, CODE_DETAL, BK, BO, ML, MP, TYPENAME, TYPEID, NAPRNAME, NAPRID)
SELECT
    b.ZAKNUM, b.ZAKID, b.ADRESS, b.KONNUM, b.KONID, b.KONNAME,
    b.NAME, b.NAMEID, b.L, b.H, b.CODE_DETAL,
    b.BK, b.BO, b.ML, b.MP,
    b.TypeName, b.TYPEID, b.NAPRNAME, b.NAPRID
FROM `{SRC_DB}`.`BASE8` b
LEFT JOIN `{DST_DB}`.`CCALC_FRN` c ON c.ZAKNUM = b.ZAKNUM
WHERE c.ZAKNUM IS NULL
"""

SQL_SP = f"""
INSERT IGNORE INTO `{DST_DB}`.`CCALC_SP`
    (BCP, BARCODEIZD, KONNUM, ELEMID, ZAKNUM, KONNAME,
     KONID, SPNAME, L, H, KOLVO, dat)
SELECT
    CONCAT_WS(' ', b7.BARCODEIZD, b6.KONNUM, b6.ELEMID) AS BCP,
    b7.BARCODEIZD, b6.KONNUM, b6.ELEMID,
    b6.ZAKNUM, b6.KONNAME, 0 AS KONID,
    b6.SPNAME, b6.L, b6.H, b6.KOLVO, NOW() AS dat
FROM (
    SELECT DISTINCT ZAKNUM, KONNAME, RAMAID, BARCODEIZD
    FROM `{SRC_DB}`.`BASE7`
) b7
INNER JOIN `{SRC_DB}`.`BASE6` b6
     ON b6.RAMAID  = b7.RAMAID
    AND b6.KONNAME = b7.KONNAME
    AND b6.ZAKNUM  = b7.ZAKNUM
LEFT JOIN `{DST_DB}`.`CCALC_SP` c ON c.ZAKNUM = b6.ZAKNUM
WHERE c.ZAKNUM IS NULL
"""

SQL_TBL = f"""
INSERT INTO `{DST_DB}`.`CCALC_TBL`
    (GR, ZAKNUM, ZAKID, ADRESS, KONNUM, KONID, KONNAME,
     ART, MATART, MATNAME, MATCOLOR, MATCOLORID, MATID,
     L, L4CALC, ELEMID, OBJID, SOURCETYPE, SOURCEID, RAMAID)
SELECT
    n.`Вариант` AS GR,
    b.ZAKNUM, b.ZAKID, b.ADRESS, b.KONNUM, b.KONID, b.KONNAME,
    b.ART, b.MATART, b.MATNAME, b.MATCOLOR, b.MATCOLORID, b.MATID,
    b.L, b.L4CALC, b.ELEMID, b.OBJID, b.SOURCETYPE, b.SOURCEID, b.RAMAID
FROM `{SRC_DB}`.`BASE14` b
LEFT JOIN `{SRC_DB}`.`NOM`    n ON n.MATART = b.MATART
LEFT JOIN `{DST_DB}`.`CCALC_TBL` c ON c.ZAKNUM = b.ZAKNUM
WHERE b.KONID > 0 AND c.ZAKNUM IS NULL
"""

REPORTS = [
    ('CCALC',     SQL_CCALC),
    ('CCALC_FRN', SQL_FRN),
    ('CCALC_SP',  SQL_SP),
    ('CCALC_TBL', SQL_TBL),
]


# ================= ВСПОМОГАТЕЛЬНОЕ =================
def register_fonts():
    pdfmetrics.registerFont(TTFont(FONT_NAME, ARIAL))
    pdfmetrics.registerFont(TTFont(FONT_BOLD, ARIAL_BOLD))


def connect_db(db):
    return mysql.connector.connect(
        host=DB_HOST, user=DB_USER, password=DB_PASSWORD,
        database=db, charset=DB_CHARSET, autocommit=False,
    )


def clean_db_value(v):
    if v is None:
        return None
    if isinstance(v, (bytes, bytearray)):
        return None
    if isinstance(v, str):
        return v.strip()[:65535]
    return v


def find_dbf(table):
    for ext in ('.DBF', '.dbf'):
        p = DBF_PATH / f'{table}{ext}'
        if p.exists():
            return p
    return None


# ================= МИГРАТОР =================
class DbfMigrator:
    def __init__(self, dbf_path, logger):
        self.dbf_path = Path(dbf_path)
        self.log = logger

    # ---- DBF -> MySQL ----
    def _get_dbf_zaknums(self, dbf_file):
        dbf = DBF(dbf_file, encoding='cp866', ignore_missing_memofile=True)
        zak_field = next((f.name for f in dbf.fields if f.name.upper() == 'ZAKNUM'), None)
        if not zak_field:
            return set()
        return {str(r.get(zak_field)).strip() for r in dbf
                if r.get(zak_field) not in (None, '')}

    def _get_mysql_zaknums(self, conn, table='BASE1'):
        cur = conn.cursor()
        cur.execute(f'SELECT ZAKNUM FROM `{SRC_DB}`.`{table}` WHERE ZAKNUM IS NOT NULL')
        s = {str(r[0]).strip() for r in cur.fetchall()}
        cur.close()
        return s

    def _migrate_table(self, conn, table, dbf_file, new_zaknums):
        dbf = DBF(dbf_file, encoding='cp866', ignore_missing_memofile=True)
        fields = [f.name for f in dbf.fields]
        zak_field = next((f for f in fields if f.upper() == 'ZAKNUM'), None)
        if not zak_field:
            self.log.warning(f'   {table}: нет ZAKNUM — пропуск')
            return 0

        cols = ', '.join(f'`{f}`' for f in fields)
        ph = ', '.join(['%s'] * len(fields))
        sql = f'INSERT INTO `{SRC_DB}`.`{table}` ({cols}) VALUES ({ph})'

        cur = conn.cursor()
        batch, total = [], 0
        for rec in dbf:
            zak = rec.get(zak_field)
            if zak is None or str(zak).strip() not in new_zaknums:
                continue
            batch.append(tuple(clean_db_value(rec.get(f)) for f in fields))
            total += 1
            if len(batch) >= BATCH_SIZE:
                cur.executemany(sql, batch)
                conn.commit()
                batch = []
        if batch:
            cur.executemany(sql, batch)
            conn.commit()
        cur.close()
        return total

    def run_migrate(self):
        self.log.info('=' * 60)
        self.log.info(f'▶ МИГРАЦИЯ DBF -> {SRC_DB}')
        self.log.info(f'   Папка: {self.dbf_path}')
        t0 = time.time()

        base1 = find_dbf('BASE1')
        if not base1:
            self.log.error(f'❌ BASE1.DBF не найден в {self.dbf_path}')
            return

        self.log.info('   Читаю BASE1.DBF...')
        dbf_zak = self._get_dbf_zaknums(base1)
        self.log.info(f'   Заказов в DBF  : {len(dbf_zak):,}')

        conn = connect_db(SRC_DB)
        try:
            mysql_zak = self._get_mysql_zaknums(conn, 'BASE1')
            self.log.info(f'   Заказов в БД   : {len(mysql_zak):,}')

            new_zak = dbf_zak - mysql_zak
            self.log.info(f'   Новых заказов  : {len(new_zak):,}')

            if not new_zak:
                self.log.info('✅ Новых данных нет')
                return

            for table in MIGRATE_TABLES:
                dbf = find_dbf(table)
                if not dbf:
                    self.log.warning(f'   {table}: файл не найден — пропуск')
                    continue
                self.log.info(f'   {table}: мигрирую...')
                n = self._migrate_table(conn, table, dbf, new_zak)
                self.log.info(f'   {table}: добавлено {n:,}')

            self.log.info(f'✅ Миграция завершена за {time.time() - t0:.1f} сек')
        finally:
            conn.close()

    # ---- отчёты ----
    def _count(self, cur, table):
        cur.execute(f'SELECT COUNT(*) FROM `{DST_DB}`.`{table}`')
        return cur.fetchone()[0]

    def _run_one_report(self, conn, label, sql):
        self.log.info('─' * 60)
        self.log.info(f'▶ Отчёт {label}')
        t0 = time.time()
        cur = conn.cursor()
        try:
            before = self._count(cur, label)
            self.log.info(f'   Строк до     : {before:,}')
            cur.execute(sql)
            added = cur.rowcount
            conn.commit()
            after = self._count(cur, label)
            self.log.info(f'   Добавлено    : {added:,}')
            self.log.info(f'   Строк после  : {after:,}  (+{after - before:,})')
            self.log.info(f'✅ {label} за {time.time() - t0:.1f} сек')
        except Exception as e:
            conn.rollback()
            self.log.error(f'❌ Ошибка в {label}: {e}')
            raise
        finally:
            cur.close()

    def run_reports(self):
        self.log.info('=' * 60)
        self.log.info(f'▶ ГЕНЕРАЦИЯ ОТЧЁТОВ в {DST_DB}')
        self.log.info(f'   Состав: {", ".join(label for label, _ in REPORTS)}')
        t0 = time.time()
        conn = connect_db(DST_DB)
        try:
            for label, sql in REPORTS:
                self._run_one_report(conn, label, sql)
        finally:
            conn.close()
        self.log.info(f'✅ Все отчёты за {time.time() - t0:.1f} сек')

    def run_all(self):
        self.run_migrate()
        self.run_reports()


# ================= ACCESS (кэш) =================
def fetch_orders_from_mdb(force_refresh=False):
    global _mdb_cache, _mdb_cache_time
    now = time.time()
    if not force_refresh and _mdb_cache is not None and (now - _mdb_cache_time < MDB_CACHE_TTL):
        return _mdb_cache
    if not os.path.exists(MDB_PATH):
        raise RuntimeError(f'Файл не найден: {MDB_PATH}')
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise RuntimeError('Не установлен pywin32. Выполните: pip install pywin32')

    pythoncom.CoInitialize()
    try:
        providers = ['Microsoft.ACE.OLEDB.12.0', 'Microsoft.Jet.OLEDB.4.0']
        conn = None
        last_err = None
        for provider in providers:
            try:
                conn_str = (f'Provider={provider};Data Source={MDB_PATH};'
                            f'Persist Security Info=False;')
                conn = win32com.client.Dispatch("ADODB.Connection")
                conn.Open(conn_str)
                break
            except Exception as e:
                last_err = e
                conn = None
        if not conn:
            raise RuntimeError(f'Не удалось открыть Access:\n{last_err}')

        sql = f"""
            SELECT DISTINCT [Заказчи] AS ZAKNUM,
                   [Дата заявки] AS DateOrder,
                   [Дата отгрузки] AS DateShip
            FROM [{MDB_TABLE}]
            WHERE [Заказчи] IS NOT NULL AND [Заказчи] <> ''
            ORDER BY [Дата заявки] DESC
        """
        rs = conn.Execute(sql)[0]
        rows = []
        while not rs.EOF:
            rows.append((rs.Fields("ZAKNUM").Value,
                         rs.Fields("DateOrder").Value,
                         rs.Fields("DateShip").Value))
            rs.MoveNext()
        rs.Close(); conn.Close()
    finally:
        pythoncom.CoUninitialize()

    def _fmt(v):
        if v is None: return ''
        if hasattr(v, 'strftime'): return v.strftime('%d.%m.%Y')
        s = str(v).strip()
        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try: return datetime.strptime(s[:19], fmt).strftime('%d.%m.%Y')
            except Exception: pass
        return s

    result = []
    for r in rows:
        zak = str(r[0]).strip()
        if not zak: continue
        result.append((zak, _fmt(r[1]), _fmt(r[2])))

    _mdb_cache = result
    _mdb_cache_time = time.time()
    return result


# ================= МИГРАЦИЯ В GUI =================
class QueueLogHandler(logging.Handler):
    def __init__(self, q):
        super().__init__()
        self.q = q
    def emit(self, record):
        try:
            self.q.put(self.format(record))
        except Exception:
            pass


# ================= MYSQL для наклеек =================
def fetch_parts(zaknum):
    conn = connect_db(DST_DB)
    cur = conn.cursor(dictionary=True)
    cur.execute(f"""
        SELECT ZAKNUM, KONNAME, NUMIZDELIE, DETAL, PRIMPROIZV,
               BCI, PROFILTYPE, KONID, GRI, RAMAID,
               MATNAME, MATART, ELEMID, OBJID, MATCOLOR,
               LENGTH, ZAKID, ANGLE_L, ANGLE_R, STAR,
               ORIENT, ORIENT2, IMPOST, RUCHKA, ZAMOK, DUGA
        FROM `{DST_DB}`.`CCALC` WHERE ZAKNUM = %s
    """, (zaknum,))
    parts = cur.fetchall()
    if not parts:
        cur.close(); conn.close(); return []

    cur.execute(f"""
        SELECT ZAKNUM, KONID, NAME, TYPENAME, NAPRNAME
        FROM `{DST_DB}`.`CCALC_FRN` WHERE ZAKNUM = %s
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
    def gri_of(p): return str(p.get('GRI') or '').strip().lower()
    def is_allowed(p):
        if gri_of(p) in ALLOWED_GRI: return True
        if 'штульп' in str(p.get('MATNAME') or '').lower(): return True
        return False

    filtered = [p for p in parts if is_allowed(p)]

    def key_rs(p):
        return (_int_or_inf(p.get('KONID')), _int_or_inf(p.get('OBJID')),
                ORIENT_ORDER.get((p.get('ORIENT') or '').strip(), 99),
                str(p.get('DETAL') or ''))

    rams    = sorted([p for p in filtered if gri_of(p) == 'r'], key=key_rs)
    stvorki = sorted([p for p in filtered if gri_of(p) == 's'], key=key_rs)
    imposts = [p for p in filtered if gri_of(p) == '3']
    porogi  = [p for p in filtered if gri_of(p) == 'п']
    shtulps = [p for p in filtered if gri_of(p) not in ('r', 's', '3', 'п')]
    groups = [('РАМЫ', rams), ('СТВОРКИ', stvorki), ('ИМПОСТА', imposts),
              ('ШТУЛЬПЫ', shtulps), ('ПОРОГИ', porogi)]
    return [(n, l) for n, l in groups if l]


# ================= РИСОВАНИЕ =================
_image_cache, _qr_cache = {}, {}


def get_imagereader(zaknum, konid):
    if IMG_DIR is None: return None
    key = (str(zaknum), str(konid))
    if key in _image_cache: return _image_cache[key]
    filepath = IMG_DIR / f"{zaknum} ({konid}).jpg"
    if not filepath.exists():
        _image_cache[key] = None
        return None
    try:
        img = Image.open(str(filepath)); img.load()
        w, h = img.size
        if w > 80 and h > 80: img = img.crop((40, 40, w - 40, h - 40))
        if img.mode not in ('RGB', 'L'): img = img.convert('RGB')
        buf = io.BytesIO(); img.save(buf, format='JPEG', quality=88)
        buf.seek(0); reader = ImageReader(buf)
    except Exception:
        reader = None
    _image_cache[key] = reader
    return reader


def get_qr_reader(data):
    data = str(data or '')
    if data in _qr_cache: return _qr_cache[data]
    if not data:
        _qr_cache[data] = None; return None
    try:
        qr = qrcode.QRCode(version=None,
                           error_correction=qrcode.constants.ERROR_CORRECT_M,
                           box_size=10, border=1)
        qr.add_data(data); qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white').convert('RGB')
        buf = io.BytesIO(); img.save(buf, format='PNG')
        buf.seek(0); reader = ImageReader(buf)
    except Exception:
        reader = None
    _qr_cache[data] = reader
    return reader


def clean_color(v):
    if not v: return ''
    s = str(v).strip()
    return s.lstrip('#').strip() if s.startswith('#') else s


def build_role(rec):
    kon, obj = rec.get('KONID') or '', rec.get('OBJID') or ''
    gri = str(rec.get('GRI') or '').strip().lower()
    letter = (rec.get('ORIENT2') or '') if gri == '3' else (rec.get('ORIENT') or '')
    letter = str(letter).strip()
    suffix = 'Рчк' if rec.get('RUCHKA') == -1 else ''
    base = f"K{kon}.{obj}"
    if letter: base += f" {letter}"
    return base + suffix


def build_orient(rec):
    t = (rec.get('FRN_TYPENAME') or '').strip()
    n = (rec.get('FRN_NAPRNAME') or '').strip()
    return f"{t} – {n}" if t and n else (t or n)


def fit_text(c, text, font, size, max_w):
    text = str(text or '')
    if not text or max_w <= 0: return ''
    while text and c.stringWidth(text, font, size) > max_w: text = text[:-1]
    return text


def wrap_text(c, text, font, size, max_w, max_lines=3):
    text = str(text or '').strip()
    if not text: return []
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
                    cur = w; break
            cur = w
    if cur and len(lines) < max_lines: lines.append(cur)
    elif cur and len(lines) == max_lines:
        tail = lines[-1]
        while tail and c.stringWidth(tail + '…', font, size) > max_w:
            tail = tail[:-1]
        lines[-1] = (tail + '…') if tail else '…'
    return lines[:max_lines]


def draw_label(c, rec):
    # Картинка
    reader = get_imagereader(rec.get('ZAKNUM'), rec.get('KONID'))
    if reader is not None:
        try:
            iw0, ih0 = reader.getSize()
            h = IMG_MAX_H; w = iw0 * (h / ih0)
            img_x = IMG_RIGHT_EDGE - w
            img_y = (LABEL_H - h) / 2 if IMG_V_CENTER else 0
            c.drawImage(reader, img_x, img_y, width=w, height=h, mask='auto')
        except Exception:
            c.setStrokeColor(grey); c.rect(IMG_RIGHT_EDGE - 42 * mm, 0, 42 * mm, LABEL_H)
    else:
        c.setStrokeColor(grey); c.rect(IMG_RIGHT_EDGE - 42 * mm, 0, 42 * mm, LABEL_H)
        c.setFont(FONT_BOLD, SZ_NOFILE); c.setFillColor(grey)
        c.drawCentredString(IMG_RIGHT_EDGE - 21 * mm, LABEL_H / 2, 'нет файла')
        c.setFillColor(black)

    # QR
    bc_val = str(rec.get('DETAL') or '')
    qr_reader = get_qr_reader(bc_val)
    if qr_reader is not None:
        try:
            c.drawImage(qr_reader, QR_X, QR_Y, width=QR_SIZE, height=QR_SIZE,
                        preserveAspectRatio=True, mask='auto')
        except Exception:
            c.setStrokeColor(grey); c.rect(QR_X, QR_Y, QR_SIZE, QR_SIZE)
    else:
        c.setStrokeColor(grey); c.rect(QR_X, QR_Y, QR_SIZE, QR_SIZE)

    if bc_val:
        c.setFont(FONT_BOLD, SZ_QR_VAL); c.setFillColor(black)
        c.drawString(QR_X, DETAL_Y, bc_val)

    LW_EDGE = LABEL_W - 1.5 * mm

    # Строка 1
    zak = str(rec.get('ZAKNUM') or '')
    c.setFont(FONT_BOLD, SZ_ZAKNUM); c.drawString(LEFT_X, Y_L1, zak)
    zak_w = c.stringWidth(zak, FONT_BOLD, SZ_ZAKNUM)
    role = build_role(rec)
    c.setFont(FONT_BOLD, SZ_ROLE)
    gap_role = c.stringWidth(' ', FONT_BOLD, SZ_ROLE)
    c.drawString(LEFT_X + zak_w + gap_role, Y_L1, role)

    # Строка 2
    matname = (rec.get('MATNAME') or '').strip()
    color_txt = clean_color(rec.get('MATCOLOR'))
    c.setFont(FONT_BOLD, SZ_MATNAME); c.drawString(LEFT_X, Y_L2, matname)
    mat_w = c.stringWidth(matname, FONT_BOLD, SZ_MATNAME)
    if color_txt:
        c.setFont(FONT_BOLD, SZ_COLOR)
        gap_color = c.stringWidth(' ', FONT_BOLD, SZ_MATNAME)
        c.drawString(LEFT_X + mat_w + gap_color, Y_L2, color_txt)

    # Строка 3
    l_str = f"L:{int(rec.get('LENGTH') or 0)}"
    al, ar = int(rec.get('ANGLE_L') or 0), int(rec.get('ANGLE_R') or 0)
    c.setFont(FONT_BOLD, SZ_LENGTH); c.drawString(TXT_X, Y_L3, l_str)
    l_w = c.stringWidth(l_str, FONT_BOLD, SZ_LENGTH)
    gap3 = c.stringWidth('  ', FONT_BOLD, SZ_LENGTH)
    c.setFont(FONT_BOLD, SZ_ANGLES)
    c.drawString(TXT_X + l_w + gap3, Y_L3, f"{al}° {ar}°")

    # Строки 4-6: фурнитура + ПО-Лв
    orient = build_orient(rec)
    frn_name = (rec.get('FRN_NAME') or '').strip()
    c.setFont(FONT_BOLD, SZ_FRN)
    lines = wrap_text(c, frn_name, FONT_BOLD, SZ_FRN, FRN_MAX_W, max_lines=FRN_MAX_LINES)
    y_positions = [Y_L4, Y_L4B, Y_L4C]
    last_x, last_y = TXT_X, Y_L4
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

    # Строка 7: Упл. + ШТП
    seal = clean_color(rec.get('SEAL_COLOR'))
    stap = (rec.get('STAP_NAME') or '').strip()
    parts7 = []
    if seal: parts7.append(f"Упл. {seal}")
    if stap: parts7.append(f"ШТП: {stap}")
    if parts7:
        c.setFont(FONT_BOLD, SZ_SEAL)
        c.drawString(TXT_X, Y_L7, "  ".join(parts7))

    # Строка 8
    prim = (rec.get('PRIMPROIZV') or '').strip()
    if prim:
        c.setFont(FONT_BOLD, SZ_PRIM)
        prim_w = LW_EDGE - TXT_X
        lines = wrap_text(c, prim, FONT_BOLD, SZ_PRIM, prim_w, max_lines=PRIM_MAX_LINES)
        if lines: c.drawString(TXT_X, Y_L8, lines[0])
        if len(lines) > 1: c.drawString(TXT_X, Y_L8B, lines[1])


# ================= PDF =================
def build_pdf_to(target, zaknum, parts):
    register_fonts()
    c = Canvas(target, pagesize=PAGE_SIZE)
    c.setTitle(f'{VERSION} — Наклейки {zaknum}')
    for group_name, items in sort_for_print(parts):
        for rec in items:
            try: draw_label(c, rec)
            except Exception as e: print(f"⚠ {rec.get('DETAL')}: {e}")
            finally: c.showPage()
    c.save()


def build_pdf_file(zaknum, out_path, parts):
    build_pdf_to(str(out_path), zaknum, parts)


def build_pdf_bytes(zaknum, parts):
    buf = io.BytesIO(); build_pdf_to(buf, zaknum, parts); buf.seek(0)
    return buf


# ================= ПРИНТЕРЫ =================
def list_printers():
    try:
        import win32print
        ps = [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        if ps: return sorted(set(ps))
    except Exception: pass
    try:
        import subprocess
        out = subprocess.check_output(['wmic', 'printer', 'get', 'name'],
                                      stderr=subprocess.DEVNULL, text=True, encoding='cp866')
        lines = [l.strip() for l in out.splitlines()[1:] if l.strip()]
        if lines: return sorted(set(lines))
    except Exception: pass
    return []


def print_to_printer(pdf_path, printer_name=None):
    pdf_path = str(pdf_path)
    if printer_name:
        try:
            import win32api
            win32api.ShellExecute(0, 'printto', pdf_path, f'"{printer_name}"', '.', 0)
            return True
        except Exception: pass
    try:
        os.startfile(pdf_path, 'print'); return True
    except Exception: return False


# ================= CLI =================
def run_cli(zaknum, out_path):
    print(f'📖 {VERSION}  заказ {zaknum}')
    recs = fetch_parts(zaknum)
    if not recs:
        print('❌ Деталей не найдено'); sys.exit(1)
    total = sum(len(l) for _, l in sort_for_print(recs))
    print(f'   Деталей: {len(recs)}   страниц: {total}')
    build_pdf_file(zaknum, out_path, recs)
    print(f'✅ Готово: {out_path}')


# ================= GUI =================
def run_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog, scrolledtext
    from PIL import ImageTk, Image as PILImage
    import pypdfium2 as pdfium

    class OrderPickerWindow:
        def __init__(self, parent, on_select):
            self.on_select = on_select
            self.win = tk.Toplevel(parent)
            self.win.title('Заказы из Access (лист1) — OLEDB')
            self.win.geometry('750x550')
            self.win.transient(parent); self.win.grab_set()

            top = ttk.Frame(self.win, padding=6); top.pack(side=tk.TOP, fill=tk.X)
            ttk.Label(top, text=f'База: {MDB_PATH}', font=('Arial', 9),
                      foreground='gray').pack(side=tk.LEFT)
            self.search_var = tk.StringVar()
            self.search_var.trace('w', lambda *a: self.filter_rows())
            ttk.Label(top, text='Поиск:').pack(side=tk.LEFT, padx=(20, 4))
            ttk.Entry(top, textvariable=self.search_var, width=20,
                      font=('Arial', 10)).pack(side=tk.LEFT)
            ttk.Button(top, text='🔄 Обновить из БД',
                       command=self.refresh_cache).pack(side=tk.RIGHT, padx=6)
            self.last_update_lbl = ttk.Label(top, text='', font=('Arial', 9),
                                              foreground='blue')
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

            bar = ttk.Frame(self.win, padding=6); bar.pack(side=tk.BOTTOM, fill=tk.X)
            ttk.Button(bar, text='Выбрать', command=self.pick).pack(side=tk.LEFT)
            ttk.Button(bar, text='Закрыть', command=self.win.destroy).pack(side=tk.LEFT, padx=6)

            self.rows_all = []
            self.load()

        def load(self, force=False):
            self.tree.delete(*self.tree.get_children()); self.rows_all = []
            try:
                self.rows_all = fetch_orders_from_mdb(force_refresh=force)
                t = datetime.fromtimestamp(_mdb_cache_time).strftime('%H:%M:%S')
                self.last_update_lbl.config(text=f'Кэш от {t}')
            except Exception as e:
                messagebox.showerror('Ошибка чтения Access', str(e), parent=self.win); return
            for zak, d1, d2 in self.rows_all:
                self.tree.insert('', 'end', values=(zak, d1, d2))

        def refresh_cache(self): self.load(force=True)

        def filter_rows(self):
            q = self.search_var.get().strip().lower()
            self.tree.delete(*self.tree.get_children())
            for zak, d1, d2 in self.rows_all:
                if q and q not in zak.lower(): continue
                self.tree.insert('', 'end', values=(zak, d1, d2))

        def pick(self):
            sel = self.tree.selection()
            if not sel: return
            vals = self.tree.item(sel[0], 'values')
            self.on_select(vals[0]); self.win.destroy()

    class SyncWindow:
        """Окно миграции DBF + отчётов."""
        def __init__(self, parent):
            self.parent = parent
            self.win = tk.Toplevel(parent)
            self.win.title('🔄 Синхронизация БД: DBF → CCALC → workshop')
            self.win.geometry('900x600')
            self.win.transient(parent)

            self.log_queue = queue.Queue()
            self.running = False

            # --- кнопки ---
            top = ttk.Frame(self.win, padding=6)
            top.pack(side=tk.TOP, fill=tk.X)

            self.btn_full = ttk.Button(top, text='▶ Миграция + отчёты',
                                        command=self.run_full)
            self.btn_full.pack(side=tk.LEFT)

            self.btn_reports = ttk.Button(top, text='📊 Только отчёты',
                                           command=self.run_reports_only)
            self.btn_reports.pack(side=tk.LEFT, padx=6)

            ttk.Button(top, text='🧹 Очистить лог',
                       command=self.clear_log).pack(side=tk.LEFT, padx=6)

            self.status = ttk.Label(top, text='Готово', foreground='blue',
                                     font=('Arial', 10))
            self.status.pack(side=tk.RIGHT)

            # --- лог ---
            self.log_text = scrolledtext.ScrolledText(
                self.win, wrap=tk.WORD, font=('Consolas', 9),
                bg='#1e1e1e', fg='#d4d4d4', insertbackground='#d4d4d4',
                state='disabled')
            self.log_text.pack(side=tk.TOP, fill=tk.BOTH, expand=True,
                                padx=6, pady=6)

            self.win.protocol("WM_DELETE_WINDOW", self._on_close)

            # --- логгер ---
            self.logger = self._setup_logger()
            self._poll_queue()

        def _setup_logger(self):
            lg = logging.getLogger(f'naKL4.5_sync_{id(self)}')
            lg.setLevel(logging.INFO)
            lg.propagate = False

            # Файловый
            base = Path(sys.executable).parent if getattr(sys, 'frozen', False) \
                   else Path(__file__).parent
            log_path = base / f'sync_{datetime.now():%Y%m%d}.log'
            fh = logging.FileHandler(log_path, encoding='utf-8')
            fh.setFormatter(logging.Formatter('%(asctime)s | %(message)s',
                                              datefmt='%H:%M:%S'))
            lg.addHandler(fh)

            # В UI
            qh = QueueLogHandler(self.log_queue)
            qh.setFormatter(logging.Formatter('%(asctime)s | %(message)s',
                                              datefmt='%H:%M:%S'))
            lg.addHandler(qh)
            return lg

        def _poll_queue(self):
            try:
                while True:
                    msg = self.log_queue.get_nowait()
                    if msg == '__done__':
                        self._finish()
                        continue
                    self._append(msg)
            except queue.Empty:
                pass
            self.win.after(100, self._poll_queue)

        def _append(self, text):
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, text + '\n')
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')

        def clear_log(self):
            self.log_text.config(state='normal')
            self.log_text.delete('1.0', tk.END)
            self.log_text.config(state='disabled')

        def _finish(self):
            self.running = False
            self.btn_full.config(state='normal')
            self.btn_reports.config(state='normal')
            self.status.config(text=f'Готово: {datetime.now().strftime("%H:%M:%S")}',
                                foreground='green')

        def _start(self):
            self.running = True
            self.btn_full.config(state='disabled')
            self.btn_reports.config(state='disabled')
            self.status.config(text='Выполняется...', foreground='orange')

        def _worker(self, mode):
            try:
                mig = DbfMigrator(DBF_PATH, self.logger)
                if mode == 'full':
                    mig.run_all()
                else:
                    mig.run_reports()
            except Exception as e:
                self.logger.exception(f'❌ Ошибка: {e}')
            finally:
                self.log_queue.put('__done__')

        def run_full(self):
            if self.running: return
            self._start()
            self.logger.info('▶ Запуск полной синхронизации (DBF + отчёты)')
            threading.Thread(target=self._worker, args=('full',), daemon=True).start()

        def run_reports_only(self):
            if self.running: return
            self._start()
            self.logger.info('▶ Запуск генерации отчётов (без миграции DBF)')
            threading.Thread(target=self._worker, args=('reports',), daemon=True).start()

        def _on_close(self):
            if self.running:
                if not messagebox.askyesno('Синхронизация идёт',
                                            'Процесс ещё выполняется. Закрыть окно?',
                                            parent=self.win):
                    return
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
            ttk.Button(top, text='💾 Сохранить PDF',
                       command=self.save_pdf).pack(side=tk.LEFT, padx=3)
            ttk.Button(top, text='🔄 Синхронизация БД',
                       command=self.open_sync).pack(side=tk.LEFT, padx=3)
            ttk.Button(top, text='🗑 Очистить',
                       command=self.clear).pack(side=tk.LEFT, padx=3)

            printers = list_printers()
            self.printer_var = tk.StringVar()
            ttk.Label(top, text='Принтер:',
                      font=('Arial', 10)).pack(side=tk.LEFT, padx=(20, 4))
            self.printer_cb = ttk.Combobox(top, textvariable=self.printer_var,
                                            values=printers, width=28,
                                            state='readonly')
            if printers:
                idx = 0
                for i, p in enumerate(printers):
                    if DEFAULT_PRINTER.lower() in p.lower():
                        idx = i; break
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

        def open_sync(self):
            SyncWindow(self.root)

        def open_order_picker(self):
            OrderPickerWindow(self.root, self.on_order_selected)

        def on_order_selected(self, zaknum):
            self.zak_entry.delete(0, tk.END)
            self.zak_entry.insert(0, zaknum)
            self.zak_entry.focus_set()
            self.set_status(f'Выбран заказ: {zaknum}', 'green')
            self.load()

        def paste_from_clipboard(self):
            try: text = self.root.clipboard_get()
            except tk.TclError: text = ''
            if text:
                self.zak_entry.delete(0, tk.END)
                self.zak_entry.insert(0, text.strip())
                self.zak_entry.focus_set()
                self.load()

        def _on_ctrl_v(self, event=None):
            try: text = self.root.clipboard_get()
            except tk.TclError: return 'break'
            if text:
                self.zak_entry.delete(0, tk.END)
                self.zak_entry.insert(0, text.strip())
                self.zak_entry.focus_set()
                self.load()
            return 'break'

        def set_status(self, text, color='blue'):
            self.status.configure(text=text, foreground=color)
            self.root.update_idletasks()

        def load(self):
            zaknum = self.zak_entry.get().strip()
            if not zaknum:
                messagebox.showwarning('Внимание', 'Введите номер заказа'); return
            self.last_zaknum = zaknum
            self.set_status('Чтение данных...')
            try: recs = fetch_parts(zaknum)
            except Exception as e:
                messagebox.showerror('Ошибка БД', str(e))
                self.set_status('Ошибка БД', 'red'); return
            if not recs:
                self.set_status('Деталей не найдено', 'red')
                messagebox.showinfo('Нет данных',
                                    f'По заказу {zaknum} деталей не найдено'); return

            self.set_status(f'Найдено {len(recs)} деталей, генерация PDF...')
            try: buf = build_pdf_bytes(zaknum, recs)
            except Exception as e:
                messagebox.showerror('Ошибка генерации', str(e))
                self.set_status('Ошибка генерации', 'red'); return

            tmp = Path(tempfile.gettempdir()) / f'labels_{zaknum}.pdf'
            tmp.write_bytes(buf.getvalue())
            self.pdf_path = tmp
            try: self.open_preview(buf.getvalue())
            except Exception as e:
                messagebox.showerror('Ошибка предпросмотра', str(e)); return

            groups = sort_for_print(recs)
            summary = ', '.join(f'{n}: {len(l)}' for n, l in groups)
            total = sum(len(l) for _, l in groups)
            filtered_out = len(recs) - total
            self.set_status(f'Готово: {summary} — этикеток {total} '
                             f'(отфильтровано {filtered_out}) — {tmp}', 'green')

        def open_preview(self, pdf_bytes):
            if self.pdf_doc:
                try: self.pdf_doc.close()
                except Exception: pass
            self.pdf_bytes = pdf_bytes
            self.pdf_doc = pdfium.PdfDocument(self.pdf_bytes)
            self.total_pages = len(self.pdf_doc)
            self.total_label.config(text=str(self.total_pages))
            self.show_page(0)

        def show_page(self, index):
            if not self.pdf_doc: return
            index = max(0, min(index, self.total_pages - 1))
            self.current_page = index
            self.page_var.set(str(index + 1))
            try:
                page = self.pdf_doc[index]
                bitmap = page.render(scale=3.0)
                pil = bitmap.to_pil()
                w = self.image_label.winfo_width()
                h = self.image_label.winfo_height()
                if w > 100 and h > 100:
                    pil.thumbnail((w - 4, h - 4), PILImage.LANCZOS)
                self.preview_img = ImageTk.PhotoImage(pil)
                self.image_label.config(image=self.preview_img)
            except Exception as e:
                self.set_status(f'Ошибка отрисовки {index + 1}: {e}', 'red')

        def _on_resize(self, event):
            if self.pdf_doc and self.total_pages > 0:
                self.root.after(200, lambda: self.show_page(self.current_page))

        def prev_page(self):
            if self.current_page > 0: self.show_page(self.current_page - 1)

        def next_page(self):
            if self.current_page < self.total_pages - 1:
                self.show_page(self.current_page + 1)

        def goto_page(self):
            try:
                p = int(self.page_var.get()) - 1
                if 0 <= p < self.total_pages: self.show_page(p)
                else: self.page_var.set(str(self.current_page + 1))
            except ValueError:
                self.page_var.set(str(self.current_page + 1))

        def _current_printer(self):
            return self.printer_var.get().strip() or None

        def print_single_page(self, page_index):
            if not self.pdf_path or not self.pdf_path.exists():
                messagebox.showinfo('Нет данных', 'Сначала выберите заказ'); return
            try:
                src = pdfium.PdfDocument(str(self.pdf_path))
                if page_index < 0 or page_index >= len(src):
                    src.close(); return
                new_pdf = pdfium.PdfDocument.new()
                new_pdf.import_pages(src, [page_index])
                zaknum = self.last_zaknum or 'zakaz'
                dst = (Path(tempfile.gettempdir())
                       / f'labels_{zaknum}_p{page_index + 1}.pdf')
                new_pdf.save(str(dst)); new_pdf.close(); src.close()
                ok = print_to_printer(dst, self._current_printer())
                self.set_status(f'{"Отправлено" if ok else "Открыто"}: '
                                 f'стр. {page_index + 1}', 'green')
            except Exception as e:
                messagebox.showerror('Ошибка печати', str(e))

        def print_pdf(self):
            if not self.pdf_path or not self.pdf_path.exists():
                messagebox.showinfo('Нет данных', 'Сначала выберите заказ'); return
            printer = self._current_printer()
            try:
                src = pdfium.PdfDocument(str(self.pdf_path))
                n = len(src)
                if n == 0: src.close(); return
                new_pdf = pdfium.PdfDocument.new()
                indices = list(range(n - 1, -1, -1))
                new_pdf.import_pages(src, indices)
                zaknum = self.last_zaknum or 'zakaz'
                dst = (Path(tempfile.gettempdir())
                       / f'labels_{zaknum}_reversed.pdf')
                new_pdf.save(str(dst)); new_pdf.close(); src.close()
                ok = print_to_printer(dst, printer)
                self.set_status(f'{"Отправлено" if ok else "Открыто"} '
                                 f'(в обратном порядке)'
                                 + (f' ({printer})' if printer else ''), 'green')
            except Exception as e:
                messagebox.showerror('Ошибка печати',
                                      f'Не удалось развернуть PDF: {e}')

        def save_pdf(self):
            if not self.pdf_path or not self.pdf_path.exists():
                messagebox.showinfo('Нет данных', 'Сначала выберите заказ'); return
            zaknum = self.last_zaknum or 'zakaz'
            dst = filedialog.asksaveasfilename(
                defaultextension='.pdf',
                initialfile=f'labels_{zaknum}.pdf',
                filetypes=[('PDF', '*.pdf')])
            if not dst: return
            Path(dst).write_bytes(self.pdf_path.read_bytes())
            self.set_status(f'Сохранено: {dst}', 'green')

        def clear(self):
            if self.pdf_doc:
                try: self.pdf_doc.close()
                except Exception: pass
            self.pdf_doc = None; self.pdf_bytes = None; self.pdf_path = None
            self.last_zaknum = ''
            self.current_page = 0; self.total_pages = 0
            self.page_var.set('1'); self.total_label.config(text='0')
            self.image_label.config(image=''); self.preview_img = None
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
# -*- coding: utf-8 -*-
# ============================================================
#  CCalc_Sync.py
#
#  Единственный режим работы:
#    1) Миграция новых заказов из DBF в CCALC (BASE*)
#    2) Генерация отчётов в workshop:
#         CCALC, CCALC_FRN, CCALC_SP, CCALC_TBL
#
#  Только добавление. Удаления / TRUNCATE не выполняются.
#  Запуск: CCalc_Sync.exe
# ============================================================

import sys
import time
import logging
from datetime import datetime
from pathlib import Path

import mysql.connector
from dbfread import DBF


# ================== КОНФИГ ==================
DB_HOST     = '192.168.0.10'
DB_USER     = 'workshop'
DB_PASSWORD = 'w0rK5h0p'
DB_CHARSET  = 'utf8mb4'

SRC_DB = 'CCALC'       # источники (BASE*)
DST_DB = 'workshop'    # отчёты (CCALC, CCALC_FRN, CCALC_SP, CCALC_TBL)

DBF_PATH = Path(r'\\sp-roshan-03\Ccalc6.80\Data')

MIGRATE_TABLES = ['BASE1', 'BASE3', 'BASE4', 'BASE6',
                  'BASE7', 'BASE8', 'BASE10', 'BASE14']

BATCH_SIZE = 500


# ================== SQL ОТЧЁТОВ ==================

SQL_CCALC = f"""
INSERT INTO `{DST_DB}`.`CCALC`
    (ZAKNUM, KONNAME, NUMIZDELIE, DETAL, PRIMPROIZV, BCI, PROFILTYPE,
     KONID, GRI, RAMAID, MATNAME, MATART, ELEMID, OBJID, MATCOLOR,
     LENGTH, ZAKID, ANGLE_L, ANGLE_R, STAR, ORIENT, ORIENT2,
     IMPOST, RUCHKA, ZAMOK, DUGA)
SELECT
    b7.ZAKNUM, b7.KONNAME, b7.NUMIZDELIE,
    b7.BARCODE          AS DETAL,
    b4.PRIMPROIZV       AS PRIMPROIZV,
    b7.BARCODEIZD       AS BCI,
    b7.PROFILTYPE, b7.KONID,
    n.`Вариант`         AS GRI,
    b7.RAMAID, b7.MATNAME, b7.MATART, b7.ELEMID, b7.OBJID,
    b7.MATCOLOR, b7.LENGTH, b7.ZAKID,
    b7.ANGLE_L, b7.ANGLE_R, b7.STAR,
    b7.ORIENT, b7.ORIENT2, b7.IMPOST,
    b7.RUCHKA, b7.ZAMOK, b7.DUGA
FROM      `{SRC_DB}`.`BASE7` b7
LEFT JOIN `{SRC_DB}`.`NOM`   n  ON n.MATART  = b7.MATART
LEFT JOIN `{SRC_DB}`.`BASE4` b4 ON b4.ZAKNUM = b7.ZAKNUM
                               AND b4.KONNAME = b7.KONNAME
LEFT JOIN `{DST_DB}`.`CCALC` c  ON c.DETAL   = b7.BARCODE
WHERE b7.PROFILTYPE <= 3
  AND c.DETAL IS NULL
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
FROM      `{SRC_DB}`.`BASE8` b
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
    b6.ZAKNUM, b6.KONNAME,
    0              AS KONID,
    b6.SPNAME, b6.L, b6.H, b6.KOLVO,
    NOW()          AS dat
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
    n.`Вариант`  AS GR,
    b.ZAKNUM, b.ZAKID, b.ADRESS, b.KONNUM, b.KONID, b.KONNAME,
    b.ART, b.MATART, b.MATNAME, b.MATCOLOR, b.MATCOLORID, b.MATID,
    b.L, b.L4CALC, b.ELEMID, b.OBJID, b.SOURCETYPE, b.SOURCEID, b.RAMAID
FROM      `{SRC_DB}`.`BASE14` b
LEFT JOIN `{SRC_DB}`.`NOM`    n ON n.MATART = b.MATART
LEFT JOIN `{DST_DB}`.`CCALC_TBL` c ON c.ZAKNUM = b.ZAKNUM
WHERE b.KONID > 0
  AND c.ZAKNUM IS NULL
"""

REPORTS = [
    ('CCALC',     SQL_CCALC),
    ('CCALC_FRN', SQL_FRN),
    ('CCALC_SP',  SQL_SP),
    ('CCALC_TBL', SQL_TBL),
]


# ================== СЛУЖЕБНОЕ ==================
log = logging.getLogger('sync')


def setup_logging():
    base = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    log_path = base / f'ccalc_sync_{datetime.now():%Y%m%d}.log'
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s | %(message)s',
        datefmt='%H:%M:%S',
        handlers=[logging.FileHandler(log_path, encoding='utf-8'),
                  logging.StreamHandler(sys.stdout)],
    )
    return log_path


def connect(db):
    return mysql.connector.connect(
        host=DB_HOST, user=DB_USER, password=DB_PASSWORD,
        database=db, charset=DB_CHARSET, autocommit=False,
    )


def clean_value(v):
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


# ================== МИГРАЦИЯ DBF ==================
def get_dbf_zaknums(dbf_file):
    dbf = DBF(dbf_file, encoding='cp866', ignore_missing_memofile=True)
    zak_field = next((f.name for f in dbf.fields if f.name.upper() == 'ZAKNUM'), None)
    if not zak_field:
        return set()
    return {str(r.get(zak_field)).strip() for r in dbf
            if r.get(zak_field) not in (None, '')}


def get_mysql_zaknums(conn, table='BASE1'):
    cur = conn.cursor()
    cur.execute(f'SELECT ZAKNUM FROM `{SRC_DB}`.`{table}` WHERE ZAKNUM IS NOT NULL')
    s = {str(r[0]).strip() for r in cur.fetchall()}
    cur.close()
    return s


def migrate_table(conn, table, dbf_file, new_zaknums):
    dbf = DBF(dbf_file, encoding='cp866', ignore_missing_memofile=True)
    fields = [f.name for f in dbf.fields]
    zak_field = next((f for f in fields if f.upper() == 'ZAKNUM'), None)
    if not zak_field:
        log.warning(f'   {table}: нет ZAKNUM — пропуск')
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
        batch.append(tuple(clean_value(rec.get(f)) for f in fields))
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


def run_migrate():
    log.info('=' * 60)
    log.info(f'▶ МИГРАЦИЯ DBF -> {SRC_DB}')
    log.info(f'   Папка: {DBF_PATH}')
    t0 = time.time()

    base1 = find_dbf('BASE1')
    if not base1:
        log.error(f'❌ BASE1.DBF не найден в {DBF_PATH}')
        return

    log.info('   Читаю BASE1.DBF...')
    dbf_zak = get_dbf_zaknums(base1)
    log.info(f'   Заказов в DBF  : {len(dbf_zak):,}')

    conn = connect(SRC_DB)
    try:
        mysql_zak = get_mysql_zaknums(conn, 'BASE1')
        log.info(f'   Заказов в БД   : {len(mysql_zak):,}')

        new_zak = dbf_zak - mysql_zak
        log.info(f'   Новых заказов  : {len(new_zak):,}')

        if not new_zak:
            log.info('✅ Новых данных нет')
            return

        for table in MIGRATE_TABLES:
            dbf = find_dbf(table)
            if not dbf:
                log.warning(f'   {table}: файл не найден — пропуск')
                continue
            log.info(f'   {table}: мигрирую...')
            n = migrate_table(conn, table, dbf, new_zak)
            log.info(f'   {table}: добавлено {n:,}')

        log.info(f'✅ Миграция завершена за {time.time() - t0:.1f} сек')
    finally:
        conn.close()


# ================== ОТЧЁТЫ ==================
def _count(cur, table):
    cur.execute(f'SELECT COUNT(*) FROM `{DST_DB}`.`{table}`')
    return cur.fetchone()[0]


def run_one_report(conn, label, sql):
    log.info('─' * 60)
    log.info(f'▶ Отчёт {label}')
    t0 = time.time()
    cur = conn.cursor()
    try:
        before = _count(cur, label)
        log.info(f'   Строк до     : {before:,}')
        cur.execute(sql)
        added = cur.rowcount
        conn.commit()
        after = _count(cur, label)
        log.info(f'   Добавлено    : {added:,}')
        log.info(f'   Строк после  : {after:,}  (+{after - before:,})')
        log.info(f'✅ {label} за {time.time() - t0:.1f} сек')
    except Exception as e:
        conn.rollback()
        log.error(f'❌ Ошибка в {label}: {e}')
        raise
    finally:
        cur.close()


def run_reports():
    log.info('=' * 60)
    log.info(f'▶ ГЕНЕРАЦИЯ ОТЧЁТОВ в {DST_DB}')
    log.info(f'   Состав: {", ".join(label for label, _ in REPORTS)}')
    t0 = time.time()
    conn = connect(DST_DB)
    try:
        for label, sql in REPORTS:
            run_one_report(conn, label, sql)
    finally:
        conn.close()
    log.info(f'✅ Все отчёты за {time.time() - t0:.1f} сек')


# ================== MAIN ==================
def main():
    log_path = setup_logging()
    log.info('=' * 60)
    log.info(f'CCalc_Sync  |  {datetime.now():%Y-%m-%d %H:%M:%S}')
    log.info(f'Лог: {log_path}')
    log.info('=' * 60)

    t0 = time.time()
    try:
        run_migrate()
        run_reports()
    except Exception as e:
        log.exception(f'❌ Фатальная ошибка: {e}')
        return 1

    log.info('=' * 60)
    log.info(f'🏁 Всё готово за {time.time() - t0:.1f} сек')
    return 0


if __name__ == '__main__':
    rc = main()
    if getattr(sys, 'frozen', False):
        try:
            input('\nНажмите Enter для выхода...')
        except EOFError:
            pass
    sys.exit(rc)
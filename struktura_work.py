# ============================================================
#  gen_ccalc_frn.py
#  CCALC_FRN = копия BASE8 (по ZAKNUM-анти-join).
#  Источник:  CCALC.BASE8
#  Приёмник:  workshop.CCALC_FRN
#  Только INSERT, без удаления.
# ============================================================

import time
import mysql.connector

DB_CONFIG = {
    'host':       '192.168.0.10',
    'user':       'workshop',
    'password':   'w0rK5h0p',
    'database':   'workshop',
    'charset':    'utf8mb4',
    'autocommit': False,
}

# Access: BASE81 = алиас BASE8
INSERT_COLS = (
    'ZAKNUM, ZAKID, ADRESS, KONNUM, KONID, KONNAME, NAME, NAMEID,'
    ' L, H, CODE_DETAL, BK, BO, ML, MP, TYPENAME, TYPEID, NAPRNAME, NAPRID'
)

INSERT_SQL = f"""
INSERT INTO `workshop`.`CCALC_FRN` ({INSERT_COLS})
SELECT
    b.ZAKNUM,
    b.ZAKID,
    b.ADRESS,
    b.KONNUM,
    b.KONID,
    b.KONNAME,
    b.NAME,
    b.NAMEID,
    b.L,
    b.H,
    b.CODE_DETAL,
    b.BK,
    b.BO,
    b.ML,
    b.MP,
    b.TypeName,
    b.TYPEID,
    b.NAPRNAME,
    b.NAPRID
FROM      `CCALC`.`BASE8` b
LEFT JOIN `workshop`.`CCALC_FRN` c ON c.ZAKNUM = b.ZAKNUM
WHERE c.ZAKNUM IS NULL
"""


def main():
    print('🔧 Генерация workshop.CCALC_FRN — новые ZAKNUM из CCALC.BASE8')
    t0 = time.time()

    conn = mysql.connector.connect(**DB_CONFIG)
    cur  = conn.cursor()

    cur.execute('SELECT COUNT(*) FROM `workshop`.`CCALC_FRN`')
    before = cur.fetchone()[0]
    print(f'   Строк до  : {before:,}')

    print('   INSERT ... SELECT ...')
    cur.execute(INSERT_SQL)
    added = cur.rowcount
    conn.commit()

    cur.execute('SELECT COUNT(*) FROM `workshop`.`CCALC_FRN`')
    after = cur.fetchone()[0]

    print(f'   Добавлено : {added:,}')
    print(f'   Строк после: {after:,}  (+{after - before:,})')
    print(f'✅ Готово за {time.time() - t0:.1f} сек')

    cur.close()
    conn.close()


if __name__ == '__main__':
    main()
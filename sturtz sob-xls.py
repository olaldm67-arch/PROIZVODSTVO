"""
SOB → Excel (.xlsx). Всё в одном файле.

Заполняет 10 листов-таблиц:
    data_projekt, data_projektlinie,
    data_los, data_loszeit,
    data_kstab, data_kstahl, data_kteil,
    data_werkzeug, data_kteil_status, data_krest

Проверка дубликатов:
    * Если filename уже в data_projekt — импорт пропускается.
    * Если losname уже в data_los — эта партия пропускается.
    * Чекбокс «Разрешить дубликаты» отключает проверку.

Установка:
    pip install openpyxl

Запуск:
    python "sturtz sob-xls.py"
"""
import os
import re
import sys
import subprocess
import datetime as dt
from dataclasses import dataclass, field
from typing import List, Dict, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook, load_workbook


# ============================================================
# КОНСТАНТЫ
# ============================================================

DEFAULT_LINES     = [(1, True), (2, False)]
DAUER_K           = 0.02
DAUER_SEC_PER_OP  = 5
PVCTNR_START      = 200000
DEFAULT_XLSX_NAME = "sobs.xlsx"


# ============================================================
# DTO
# ============================================================

@dataclass
class Operation:
    wkznr: str
    verstellung: int
    position: int
    zusatzmass: int


@dataclass
class Part:
    seq: int
    wagennr: int
    fachnr: int
    auftrag: str
    pos: int
    stuck: int
    rahmen: int
    element: str
    lage: str
    profnr: str
    farbe: str
    laenge: int
    cut_left: str
    cut_right: str
    druck: int
    stulp: int
    trolley: int
    stahl_profnr: str = ''
    stahl_laenge: int = 0
    operations: List[Operation] = field(default_factory=list)


@dataclass
class Bar:
    nummer: int
    laenge: int
    profnr: str
    farbe: str
    bezeichnung: str
    lagerort: int


@dataclass
class Rest:
    laenge: int
    art: int
    druck: int


@dataclass
class SobFile:
    version: str = ''
    bars: List[Bar] = field(default_factory=list)
    parts_by_bar: Dict[int, List[Part]] = field(default_factory=dict)
    rests_by_bar: Dict[int, Rest] = field(default_factory=dict)


# ============================================================
# ПАРСЕР .SOB
# ============================================================

RE_KSN = re.compile(
    r'^KSN(?P<nummer>\d{3})'
    r'L(?P<laenge>\d{5})'
    r'T(?P<prof13>\d{13})'
    r'B(?P<bez>.{20})'
    r'O(?P<lager>\d{6})'
)

RE_KTN = re.compile(
    r'^KTN(?P<seq>\d{4})'
    r'C(?P<wagen>\d{3})'
    r'F(?P<fach>\d{3})'
    r'K(?P<auftrag>\d{8})'
    r'P(?P<pos>\d{3})(?P<stuck>\d{3})(?P<rahmen>\d{3})'
    r'E(?P<elem>\d{3})'
    r'U(?P<lage>[A-Z])'
    r'T(?P<prof13>\d{13})'
    r'B(?P<bez>.{20})'
    r'L(?P<laenge>\d{5})'
    r'G(?P<cut_l>\d{4})(?P<cut_r>\d{4})'
    r'D(?P<druck>\d)S(?P<stulp>\d)W(?P<troll>\d)'
)

RE_W   = re.compile(r'^W(\d{4})/(\d{2})/(\d{5})/(\d{5})')
RE_S   = re.compile(r'^S\s*/([^/]+?)\s*/(\d+)')
RE_KRL = re.compile(r'^KRL(\d{5})A(\d)D(\d)')


def parse_sob(content: str, log=print) -> SobFile:
    sf = SobFile()
    cur_bar = -1
    stats = {'ksn': 0, 'ktn': 0, 'w': 0, 's': 0, 'krl': 0, 'skip': 0}

    for raw in content.splitlines():
        line = raw.rstrip()
        if not line:
            continue

        if line.startswith('VER'):
            sf.version = line
            continue
        if line.startswith('ZUH'):
            continue

        if line.startswith('KSN'):
            m = RE_KSN.match(line)
            if not m:
                stats['skip'] += 1
                continue
            g = m.groupdict()
            sf.bars.append(Bar(
                nummer=int(g['nummer']),
                laenge=int(g['laenge']),
                profnr=g['prof13'][:10],
                farbe=g['prof13'][10:13],
                bezeichnung=g['bez'].rstrip(),
                lagerort=int(g['lager']),
            ))
            cur_bar = len(sf.bars) - 1
            sf.parts_by_bar[cur_bar] = []
            stats['ksn'] += 1
            continue

        if line.startswith('KTN'):
            if cur_bar < 0:
                stats['skip'] += 1
                continue
            m = RE_KTN.match(line)
            if not m:
                stats['skip'] += 1
                continue
            g = m.groupdict()
            sf.parts_by_bar[cur_bar].append(Part(
                seq=int(g['seq']),
                wagennr=int(g['wagen']),
                fachnr=int(g['fach']),
                auftrag=g['auftrag'],
                pos=int(g['pos']),
                stuck=int(g['stuck']),
                rahmen=int(g['rahmen']),
                element=g['elem'],
                lage=g['lage'],
                profnr=g['prof13'][:10],
                farbe=g['prof13'][10:13],
                laenge=int(g['laenge']),
                cut_left=g['cut_l'],
                cut_right=g['cut_r'],
                druck=int(g['druck']),
                stulp=int(g['stulp']),
                trolley=int(g['troll']),
            ))
            stats['ktn'] += 1
            continue

        if line.startswith('W') and len(line) > 1 and line[1].isdigit():
            if cur_bar >= 0 and sf.parts_by_bar[cur_bar]:
                m = RE_W.match(line)
                if m:
                    wkz, ver, pos, tail = m.groups()
                    sf.parts_by_bar[cur_bar][-1].operations.append(
                        Operation(int(wkz), int(ver), int(pos), int(tail)))
                    stats['w'] += 1
            continue

        if line.startswith('S'):
            r = RE_S.match(line)
            if r and cur_bar >= 0 and sf.parts_by_bar[cur_bar]:
                sf.parts_by_bar[cur_bar][-1].stahl_profnr = r.group(1).strip()
                sf.parts_by_bar[cur_bar][-1].stahl_laenge = int(r.group(2))
                stats['s'] += 1
            continue

        if line.startswith('KRL'):
            if cur_bar >= 0:
                m = RE_KRL.match(line)
                if m:
                    l, a, d = m.groups()
                    sf.rests_by_bar[cur_bar] = Rest(int(l), int(a), int(d))
                    stats['krl'] += 1
            continue

        stats['skip'] += 1

    log(f"[PARSE] {sf.version or '(no VER)'} | "
        f"bars={stats['ksn']} parts={stats['ktn']} "
        f"W={stats['w']} S={stats['s']} KRL={stats['krl']} "
        f"skipped={stats['skip']}")
    return sf


# ============================================================
# EXCEL-СХЕМА (10 листов)
# ============================================================

TABLE_SCHEMA = {
    'data_projekt': [
        'lfdprojektnr', 'lpdatum', 'ledatum', 'rfolge', 'optimiert',
        'filename', 'filedatum', 'einlesedatum', 'zuordcode',
        'loesch', 'version', 'vrelease', 'vkrelease',
    ],
    'data_projektlinie': ['Projektnr', 'linie', 'wahl'],
    'data_los': [
        'losnr', 'losname', 'pdatum', 'erlaubnis', 'rfolge', 'workrfolge',
    ],
    'data_loszeit': ['losnr', 'stabanr', 'wahl', 'dauer'],
    'data_kstab': [
        'lfdprojektnr', 'lfdstabnr', 'nummer', 'proflaenge',
        'profnr', 'farbe', 'profbezeichnung', 'lagerort',
    ],
    'data_kstahl': [
        'lfdnr', 'pvctnr', 'readpcnr', 'readbar', 'readsort',
        'profnr', 'profkenn', 'laenge', 'wagennr', 'fachnr',
        'status', 'datum', 'zeit', 'neuaktivierung', 'neuaktivnachopti',
    ],
    'data_kteil': [
        'lfdprojektnr', 'lfdstabnr', 'lfdteilenr', 'losnr',
        'textkeynr', 'halle', 'status', 'datum', 'zeit',
        'nummer', 'wagennr', 'fachnr',
        'kommission', 'positionsname', 'elementnr', 'lage', 'utnr',
        'profnr', 'farbe', 'profbezeichnung', 'laenge',
        'anschnitt', 'abschnitt', 'stulpprofil', 'druck',
        'skenn', 'skenn2', 'spnr', 'spnr2', 'slaenge', 'slaenge2',
        'neuaktivierung', 'neuaktivnachopti', 'hilfsmerker', 'smstatus',
        'stahlfertig', 'tour', 'wagentyp', 'rapid',
        'profnr2', 'profbezeichnung2', 'gldnr', 'barcode', 'zusatzangabe',
    ],
    'data_werkzeug': [
        'lfdteilenr', 'reihenfolge', 'wkznr', 'verstellung',
        'position', 'zusatzmass',
    ],
    'data_kteil_status': [
        'lfdteilenr', 'remakeno', 'status', 'pcnr',
        'datum', 'zeit', 'usernr', 'ghu',
    ],
    'data_krest': ['lfdstabnr', 'laenge', 'art', 'druck', 'fertig'],
}


# ============================================================
# ОБЁРТКА EXCEL
# ============================================================

class ExcelDB:
    def __init__(self, path: str, log=print):
        self.path = path
        self.log = log
        if os.path.isfile(path):
            self.wb = load_workbook(path)
            log(f"[XLSX] открыт {path}")
        else:
            self.wb = Workbook()
            self.wb.remove(self.wb.active)
            log(f"[XLSX] создан {path}")
        self._ensure_sheets()

    def _ensure_sheets(self):
        for sheet, cols in TABLE_SCHEMA.items():
            if sheet not in self.wb.sheetnames:
                ws = self.wb.create_sheet(title=sheet)
                ws.append(cols)
                self.log(f"[XLSX] + лист {sheet}")
            else:
                ws = self.wb[sheet]
                existing = [c.value for c in ws[1]] if ws.max_row >= 1 else []
                if existing != cols:
                    for j, col in enumerate(cols, 1):
                        ws.cell(row=1, column=j, value=col)

    def append(self, sheet: str, row_dict: dict):
        cols = TABLE_SCHEMA[sheet]
        row = [row_dict.get(c) for c in cols]
        self.wb[sheet].append(row)

    def max_id(self, sheet: str, col: str) -> int:
        cols = TABLE_SCHEMA[sheet]
        j = cols.index(col) + 1
        ws = self.wb[sheet]
        max_v = 0
        for row in ws.iter_rows(min_row=2, min_col=j, max_col=j,
                                values_only=True):
            v = row[0]
            if isinstance(v, (int, float)) and v > max_v:
                max_v = int(v)
        return max_v

    def next_id(self, sheet: str, col: str, floor: int = 1) -> int:
        return max(floor, self.max_id(sheet, col) + 1)

    # --------------------------------------------------------
    # ПОИСК
    # --------------------------------------------------------

    def find_value(self, sheet: str, col: str, value):
        """
        Ищет value в колонке col листа sheet.
        Возвращает индекс строки (1-based, учитывая шапку), или None.
        """
        if sheet not in self.wb.sheetnames:
            return None
        if col not in TABLE_SCHEMA[sheet]:
            return None
        j = TABLE_SCHEMA[sheet].index(col) + 1
        ws = self.wb[sheet]
        target = str(value).strip() if value is not None else None
        for i, row in enumerate(
                ws.iter_rows(min_row=2, min_col=j, max_col=j,
                             values_only=True), start=2):
            v = row[0]
            if v is None:
                continue
            if str(v).strip() == target:
                return i
        return None

    def exists(self, sheet: str, col: str, value) -> bool:
        return self.find_value(sheet, col, value) is not None

    def load_column(self, sheet: str, col: str) -> set:
        """Множество всех непустых значений в колонке (как str)."""
        if sheet not in self.wb.sheetnames:
            return set()
        if col not in TABLE_SCHEMA[sheet]:
            return set()
        j = TABLE_SCHEMA[sheet].index(col) + 1
        ws = self.wb[sheet]
        out = set()
        for row in ws.iter_rows(min_row=2, min_col=j, max_col=j,
                                values_only=True):
            v = row[0]
            if v is not None:
                out.add(str(v).strip())
        return out

    def save(self):
        self.wb.save(self.path)
        self.log(f"[XLSX] сохранён {self.path}")


# ============================================================
# ПОДГОТОВКА
# ============================================================

def collect_los_groups(sob):
    groups, order = {}, []
    for bar_idx, bar in enumerate(sob.bars):
        prof13 = f"{bar.profnr}{bar.farbe}"
        if prof13 not in groups:
            groups[prof13] = {'prof13': prof13,
                              'total_len_mm': 0.0, 'total_ops': 0}
            order.append(prof13)
        for part in sob.parts_by_bar.get(bar_idx, []):
            groups[prof13]['total_len_mm'] += part.laenge / 10.0
            groups[prof13]['total_ops']    += len(part.operations)
    return [groups[p] for p in order]


def compute_dauer(total_len_mm: float, total_ops: int) -> int:
    return int(round(total_len_mm * DAUER_K + DAUER_SEC_PER_OP * total_ops))


# ============================================================
# ИМПОРТ В EXCEL
# ============================================================

def import_sob_to_excel(sob, xlsx_path, *, sob_path,
                        lpdatum=None, optimize=True,
                        lines=DEFAULT_LINES,
                        force_reimport=False,
                        log=print) -> dict:
    """
    force_reimport=False — проверять дубликаты проекта и партии;
    force_reimport=True  — импортировать без проверок.
    """
    if lpdatum is None:
        lpdatum = dt.date.today()

    filename = os.path.basename(sob_path)
    filedatum = dt.datetime.fromtimestamp(os.path.getmtime(sob_path))
    einlesedatum = dt.datetime.now()

    stats = dict.fromkeys(TABLE_SCHEMA.keys(), 0)
    db = ExcelDB(xlsx_path, log=log)

    # ============================================================
    # ПРОВЕРКА ДУБЛИКАТА ПРОЕКТА
    # ============================================================
    if not force_reimport:
        row_idx = db.find_value('data_projekt', 'filename', filename)
        if row_idx is not None:
            # Берём lfdprojektnr для информативного сообщения
            cols = TABLE_SCHEMA['data_projekt']
            j = cols.index('lfdprojektnr') + 1
            existing_nr = db.wb['data_projekt'].cell(
                row=row_idx, column=j).value
            log(f"[SKIP] Проект '{filename}' уже в базе "
                f"(lfdprojektnr={existing_nr}). Импорт отменён.")
            log("[SKIP] Снимите галочку «Разрешить дубликаты», "
                "если нужно переимпортировать с нуля.")
            return stats   # всё по нулям

    # ============================================================
    # ПРОВЕРКА СУЩЕСТВУЮЩИХ ПАРТИЙ
    # ============================================================
    existing_losnames = db.load_column('data_los', 'losname')
    groups_all = collect_los_groups(sob)
    groups = []
    for g in groups_all:
        losname_candidate = f"{filename} / {g['prof13']}"
        if not force_reimport and losname_candidate in existing_losnames:
            log(f"[SKIP] Партия '{losname_candidate}' уже есть — пропуск.")
            continue
        groups.append(g)

    if not groups:
        log("[SKIP] Все партии уже в базе. Импорт завершён.")
        db.save()
        return stats

    # --------------------------------------------------------
    # 1. data_projekt
    # --------------------------------------------------------
    lfdprojektnr = db.next_id('data_projekt', 'lfdprojektnr')
    db.append('data_projekt', {
        'lfdprojektnr': lfdprojektnr, 'lpdatum': lpdatum, 'ledatum': None,
        'rfolge': 0, 'optimiert': optimize,
        'filename': filename, 'filedatum': filedatum,
        'einlesedatum': einlesedatum, 'zuordcode': 0,
        'loesch': False, 'version': 2, 'vrelease': 0, 'vkrelease': 0,
    })
    stats['data_projekt'] = 1
    log(f"[XLSX] data_projekt        ← lfdprojektnr={lfdprojektnr} "
        f"filename={filename}")

    # --------------------------------------------------------
    # 2. data_projektlinie
    # --------------------------------------------------------
    for linie, wahl in lines:
        db.append('data_projektlinie', {
            'Projektnr': lfdprojektnr, 'linie': linie, 'wahl': wahl,
        })
        stats['data_projektlinie'] += 1
    log(f"[XLSX] data_projektlinie  ← {len(lines)} строк")

    # --------------------------------------------------------
    # 3. data_los (только отфильтрованные партии)
    # --------------------------------------------------------
    losnr_by_profile  = {}
    dauer_by_profile  = {}
    kept_profiles     = {g['prof13'] for g in groups}

    for idx, g in enumerate(groups):
        losnr = db.next_id('data_los', 'losnr')
        losname = f"{filename} / {g['prof13']}"
        dauer = compute_dauer(g['total_len_mm'], g['total_ops'])
        db.append('data_los', {
            'losnr': losnr, 'losname': losname, 'pdatum': lpdatum,
            'erlaubnis': 0, 'rfolge': 0, 'workrfolge': idx,
        })
        stats['data_los'] += 1
        losnr_by_profile[g['prof13']] = losnr
        dauer_by_profile[g['prof13']] = dauer
        log(f"[XLSX] data_los           ← losnr={losnr} "
            f"{losname} dauer={dauer}")

    # --------------------------------------------------------
    # 4. data_kstab (только стержни с сохранёнными партиями)
    # --------------------------------------------------------
    stabnr_by_baridx = {}
    for bar_idx, bar in enumerate(sob.bars):
        prof13 = f"{bar.profnr}{bar.farbe}"
        if prof13 not in kept_profiles:
            continue
        stabnr = db.next_id('data_kstab', 'lfdstabnr')
        db.append('data_kstab', {
            'lfdprojektnr':    lfdprojektnr,
            'lfdstabnr':       stabnr,
            'nummer':          bar.nummer,
            'proflaenge':      bar.laenge,
            'profnr':          prof13,
            'farbe':           bar.farbe,
            'profbezeichnung': bar.bezeichnung,
            'lagerort':        bar.lagerort,
        })
        stabnr_by_baridx[bar_idx] = stabnr
        stats['data_kstab'] += 1
    log(f"[XLSX] data_kstab         ← {stats['data_kstab']} строк")

    # --------------------------------------------------------
    # 5. data_loszeit
    # --------------------------------------------------------
    for prof13, losnr in losnr_by_profile.items():
        dauer = dauer_by_profile[prof13]
        for linie, wahl in lines:
            db.append('data_loszeit', {
                'losnr':   losnr,
                'stabanr': linie,
                'wahl':    wahl,
                'dauer':   dauer,
            })
            stats['data_loszeit'] += 1
    log(f"[XLSX] data_loszeit       ← {stats['data_loszeit']} строк")

    # --------------------------------------------------------
    # 6. data_krest
    # --------------------------------------------------------
    for bar_idx in stabnr_by_baridx:
        rest = sob.rests_by_bar.get(bar_idx)
        if rest and rest.laenge > 0:
            db.append('data_krest', {
                'lfdstabnr': stabnr_by_baridx[bar_idx],
                'laenge':    rest.laenge,
                'art':       rest.art,
                'druck':     rest.druck,
                'fertig':    False,
            })
            stats['data_krest'] += 1
    log(f"[XLSX] data_krest         ← {stats['data_krest']} строк")

    # --------------------------------------------------------
    # 7. data_kstahl
    # --------------------------------------------------------
    pvctnr = db.next_id('data_kstahl', 'pvctnr', floor=PVCTNR_START)
    first_pvctnr = pvctnr
    pvctnr_by_key = {}
    readsort_global = 0

    for bar_idx in stabnr_by_baridx:
        stabnr = stabnr_by_baridx[bar_idx]
        for part_idx, part in enumerate(sob.parts_by_bar.get(bar_idx, [])):
            stahl_profnr = part.stahl_profnr or part.profnr
            stahl_laenge = part.stahl_laenge or part.laenge
            lfdnr = db.next_id('data_kstahl', 'lfdnr')
            db.append('data_kstahl', {
                'lfdnr':             lfdnr,
                'pvctnr':            pvctnr,
                'readpcnr':          lfdprojektnr,
                'readbar':           stabnr,
                'readsort':          readsort_global,
                'profnr':            stahl_profnr,
                'profkenn':          '',
                'laenge':            stahl_laenge,
                'wagennr':           0,
                'fachnr':            str(part.fachnr),
                'status':            0,
                'datum':             lpdatum,
                'zeit':              einlesedatum,
                'neuaktivierung':    0,
                'neuaktivnachopti':  False,
            })
            stats['data_kstahl'] += 1
            pvctnr_by_key[(bar_idx, part_idx)] = pvctnr
            pvctnr += 1
            readsort_global += 1
    log(f"[XLSX] data_kstahl        ← {stats['data_kstahl']} строк, "
        f"pvctnr {first_pvctnr}..{pvctnr - 1}")

    # --------------------------------------------------------
    # 8. data_kteil
    # --------------------------------------------------------
    for bar_idx in stabnr_by_baridx:
        bar = sob.bars[bar_idx]
        stabnr = stabnr_by_baridx[bar_idx]
        prof13 = f"{bar.profnr}{bar.farbe}"
        losnr  = losnr_by_profile.get(prof13)
        if losnr is None:
            continue

        for part_idx, part in enumerate(sob.parts_by_bar.get(bar_idx, [])):
            pv = pvctnr_by_key[(bar_idx, part_idx)]

            db.append('data_kteil', {
                'lfdprojektnr':     lfdprojektnr,
                'lfdstabnr':        stabnr,
                'lfdteilenr':       pv,
                'losnr':            losnr,
                'textkeynr':        0,
                'halle':            0,
                'status':           0,
                'datum':            lpdatum,
                'zeit':             einlesedatum,
                'nummer':           part.seq,
                'wagennr':          part.wagennr,
                'fachnr':           str(part.fachnr),
                'kommission':       part.auftrag,
                'positionsname':    f"{part.pos:03d}{part.stuck:03d}{part.rahmen:03d}",
                'elementnr':        int(part.element),
                'lage':             part.lage,
                'utnr':             0,
                'profnr':           part.profnr,
                'farbe':            part.farbe,
                'profbezeichnung':  bar.bezeichnung,
                'laenge':           part.laenge,
                'anschnitt':        str(int(part.cut_left)),
                'abschnitt':        str(int(part.cut_right)),
                'stulpprofil':      part.stulp,
                'druck':            bool(part.druck),
                'skenn':            '',
                'skenn2':           '',
                'spnr':             part.stahl_profnr,
                'spnr2':            '',
                'slaenge':          part.stahl_laenge,
                'slaenge2':         0,
                'neuaktivierung':   0,
                'neuaktivnachopti': False,
                'hilfsmerker':      False,
                'smstatus':         0,
                'stahlfertig':      False,
                'tour':             0,
                'wagentyp':         0,
                'rapid':            0,
                'profnr2':          '',
                'profbezeichnung2': '',
                'gldnr':            0,
                'barcode':          '',
                'zusatzangabe':     '',
            })
            stats['data_kteil'] += 1
    log(f"[XLSX] data_kteil         ← {stats['data_kteil']} строк")

    # --------------------------------------------------------
    # 9-10. data_werkzeug + data_kteil_status
    # --------------------------------------------------------
    for bar_idx in stabnr_by_baridx:
        for part_idx, part in enumerate(sob.parts_by_bar.get(bar_idx, [])):
            pv = pvctnr_by_key[(bar_idx, part_idx)]

            for i, op in enumerate(part.operations):
                db.append('data_werkzeug', {
                    'lfdteilenr':  pv,
                    'reihenfolge': i,
                    'wkznr':       int(op.wkznr),
                    'verstellung': op.verstellung,
                    'position':    op.position,
                    'zusatzmass':  op.zusatzmass,
                })
                stats['data_werkzeug'] += 1

            db.append('data_kteil_status', {
                'lfdteilenr': pv,
                'remakeno':   0,
                'status':     0,
                'pcnr':       0,
                'datum':      lpdatum,
                'zeit':       einlesedatum,
                'usernr':     0,
                'ghu':        False,
            })
            stats['data_kteil_status'] += 1

    db.save()
    log("[XLSX] ----------------------------------------")
    for k, v in stats.items():
        log(f"[XLSX] {k:<20} {v}")
    return stats


# ============================================================
# ВСПОМОГАТЕЛЬНОЕ ДЛЯ GUI
# ============================================================

def open_folder(path: str):
    if not os.path.isdir(path):
        return
    if sys.platform.startswith('win'):
        os.startfile(path)
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


def open_file(path: str):
    if not os.path.isfile(path):
        return
    if sys.platform.startswith('win'):
        os.startfile(path)
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


# ============================================================
# GUI
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SOB → Excel (.xlsx)")
        self.geometry("940x700")
        self.minsize(840, 580)

        self.src_var      = tk.StringVar()
        self.outdir_var   = tk.StringVar()
        self.xlsxname_var = tk.StringVar(value=DEFAULT_XLSX_NAME)
        self.lp_var       = tk.StringVar()
        self.opt_var      = tk.BooleanVar(value=True)
        self.open_var     = tk.BooleanVar(value=True)
        self.force_var    = tk.BooleanVar(value=False)   # разрешить дубликаты
        self.status_var   = tk.StringVar(value="Целевой файл: —")

        pad = {'padx': 8, 'pady': 4}

        # --- Источник ---
        tk.Label(self, text="Файл .sob:",
                 font=("", 10, "bold")).grid(
            row=0, column=0, sticky='w', **pad)
        tk.Entry(self, textvariable=self.src_var, width=80).grid(
            row=0, column=1, sticky='we', **pad)
        tk.Button(self, text="Выбрать .sob…", width=16,
                  command=self.choose_src).grid(row=0, column=2, **pad)

        # --- Excel-параметры ---
        excel_frame = tk.LabelFrame(self, text="Excel", **pad)
        excel_frame.grid(row=1, column=0, columnspan=3,
                         sticky='we', padx=8, pady=6)

        tk.Label(excel_frame, text="Папка выгрузки:").grid(
            row=0, column=0, sticky='w', padx=8, pady=4)
        tk.Entry(excel_frame, textvariable=self.outdir_var,
                 width=64).grid(
            row=0, column=1, sticky='we', padx=8, pady=4)
        tk.Button(excel_frame, text="Выбрать папку…", width=16,
                  command=self.choose_outdir).grid(
            row=0, column=2, padx=8, pady=4)

        tk.Label(excel_frame, text="Имя XLSX-файла:").grid(
            row=1, column=0, sticky='w', padx=8, pady=4)
        tk.Entry(excel_frame, textvariable=self.xlsxname_var,
                 width=30).grid(
            row=1, column=1, sticky='w', padx=8, pady=4)

        excel_frame.grid_columnconfigure(1, weight=1)

        # --- Общие ---
        common = tk.Frame(self)
        common.grid(row=2, column=0, columnspan=3, sticky='we',
                    padx=8, pady=4)
        tk.Label(common, text="Дата проекта (ГГГГ-ММ-ДД):").pack(
            side='left', padx=(8, 4))
        tk.Entry(common, textvariable=self.lp_var, width=14).pack(
            side='left', padx=4)
        tk.Checkbutton(common, text="optimiert",
                       variable=self.opt_var).pack(side='left', padx=16)
        tk.Checkbutton(common, text="Открыть папку после импорта",
                       variable=self.open_var).pack(side='left', padx=16)

        # --- Проверка дубликатов ---
        tk.Checkbutton(
            self,
            text="Разрешить дубликаты (переимпорт без проверки)",
            variable=self.force_var,
            fg="#B71C1C",
        ).grid(row=3, column=1, sticky='w', padx=8, pady=(0, 4))

        # --- Строка статуса ---
        tk.Label(self, textvariable=self.status_var,
                 anchor='w', fg="#1565C0",
                 font=("Consolas", 9)).grid(
            row=4, column=0, columnspan=3, sticky='we',
            padx=8, pady=(0, 4))

        # --- Кнопки ---
        btns = tk.Frame(self)
        btns.grid(row=5, column=0, columnspan=3, sticky='e',
                  padx=8, pady=6)
        tk.Button(btns, text="Импортировать",
                  bg="#2E7D32", fg="white",
                  font=("Arial", 10, "bold"), width=18,
                  command=self.run).pack(side='left', padx=4)
        tk.Button(btns, text="Открыть целевой файл",
                  command=self.open_target).pack(side='left', padx=4)
        tk.Button(btns, text="Выход",
                  command=self.destroy).pack(side='left', padx=4)

        # --- Журнал ---
        tk.Label(self, text="Журнал:").grid(
            row=6, column=0, sticky='nw', **pad)
        self.log_text = tk.Text(self, height=20, wrap='word',
                                font=("Consolas", 9))
        self.log_text.grid(row=6, column=1, columnspan=2,
                           sticky='nsew', **pad)
        sb = ttk.Scrollbar(self, command=self.log_text.yview)
        sb.grid(row=6, column=3, sticky='ns')
        self.log_text.config(yscrollcommand=sb.set)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(6, weight=1)

        self.after(100, self._update_status)

    def log(self, msg):
        self.log_text.insert('end', msg + '\n')
        self.log_text.see('end')
        self.update_idletasks()

    def _target_path(self):
        d = self.outdir_var.get().strip()
        n = self.xlsxname_var.get().strip() or DEFAULT_XLSX_NAME
        return os.path.join(d, n) if d else ''

    def _update_status(self):
        p = self._target_path()
        self.status_var.set(f"Целевой файл: {p}" if p else "Целевой файл: —")

    def choose_src(self):
        p = filedialog.askopenfilename(
            title="Выберите .sob",
            filetypes=[("Stürtz .sob", "*.sob"), ("Все", "*.*")])
        if p:
            self.src_var.set(p)
            if not self.outdir_var.get():
                self.outdir_var.set(os.path.dirname(p))
            self._update_status()

    def choose_outdir(self):
        initial = self.outdir_var.get().strip() or None
        p = filedialog.askdirectory(
            title="Выберите папку для XLSX", initialdir=initial)
        if p:
            self.outdir_var.set(p)
            self._update_status()

    def open_target(self):
        p = self._target_path()
        if not p:
            messagebox.showwarning("Нет файла", "Целевой файл не задан.")
            return
        if os.path.isfile(p):
            open_file(p)
        else:
            d = os.path.dirname(p)
            if os.path.isdir(d):
                open_folder(d)
            else:
                messagebox.showwarning("Нет папки",
                                       f"Папка не найдена: {d}")

    def run(self):
        self.log_text.delete('1.0', 'end')

        src = self.src_var.get().strip()
        if not src or not os.path.isfile(src):
            messagebox.showerror("Ошибка", "Не найден файл .sob.")
            return

        outdir = self.outdir_var.get().strip()
        xlsxname = self.xlsxname_var.get().strip() or DEFAULT_XLSX_NAME
        if not outdir:
            messagebox.showerror("Ошибка", "Не выбрана папка выгрузки.")
            return
        if not os.path.isdir(outdir):
            try:
                os.makedirs(outdir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Ошибка",
                                     f"Не удалось создать папку:\n{e}")
                return
        target = os.path.join(outdir, xlsxname)

        lp = None
        if self.lp_var.get().strip():
            try:
                lp = dt.datetime.strptime(
                    self.lp_var.get().strip(), "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Ошибка",
                                     "Дата в формате ГГГГ-ММ-ДД")
                return

        try:
            self.log(f"Источник: {src}")
            self.log(f"Цель:     {target}")
            self.log(f"Режим:    "
                     f"{'переимпорт' if self.force_var.get() else 'с проверкой дубликатов'}")
            self.log("Чтение .sob…")

            with open(src, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()

            self.log("Парсинг .sob…")
            sob = parse_sob(content, log=self.log)

            stats = import_sob_to_excel(
                sob, target, sob_path=src,
                lpdatum=lp,
                optimize=bool(self.opt_var.get()),
                force_reimport=bool(self.force_var.get()),
                log=self.log,
            )

            total = sum(stats.values())
            self.log("Готово.")
            self._update_status()

            if self.open_var.get() and total > 0:
                open_folder(outdir)

            if total == 0:
                messagebox.showinfo(
                    "Пропущено",
                    f"Импорт не выполнен.\n"
                    f"Данные уже есть в базе.\n\n"
                    f"Подробности — в журнале.")
            else:
                summary = "\n".join(f"  {k:<20} {v}"
                                    for k, v in stats.items()
                                    if v > 0)
                messagebox.showinfo(
                    "Успех",
                    f"Импорт завершён.\n"
                    f"Файл: {target}\n\n"
                    f"Строк по таблицам:\n{summary}")
        except Exception as e:
            self.log(f"[ОШИБКА] {e!r}")
            messagebox.showerror("Ошибка", str(e))


if __name__ == '__main__':
    App().mainloop()
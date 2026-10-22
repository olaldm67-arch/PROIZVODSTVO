import re
import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ============================================================
# КОНСТАНТЫ
# ============================================================

BR_TO_PROFILE = {
    'BR-05000': '1032130010001',
    'BR-01000': '1012130010001',
    'BR-03000': '1021030010001',
    'BR-07000': '1012080010001',
    'BR-08000': '1022140010001',
    'BR-09000': '1032320010001',
}

PROFILE_TYPES = {
    '1012130010001': 'рама',
    '1012080010001': 'рама',
    '1012330010001': 'рама',
    '1021030010001': 'импост',
    '1022140010001': 'импост',
    '1022350010001': 'импост',
    '1032130010001': 'створка',
    '1032320010001': 'створка',
    '1032350010001': 'створка',
}

LOC_TO_STURTZ = {'T': 'O', 'B': 'U', 'H': 'W', 'V': 'S', 'L': 'L', 'R': 'R'}

LOCATION_MEANING = {
    'T': 'верх', 'B': 'низ', 'L': 'лево', 'R': 'право',
    'V': 'вертикальный', 'H': 'горизонтальный',
}

NEW_VER = 'Ver02.07-00'
SRC_VER = 'ZUH30'
S_CODE  = '113292'

SRC_EXT = '.sob'
DST_EXT = '.sob'

NOTE_TRIGGER_DN  = 'DN'
NOTE_TRIGGER_D25 = 'D25'

# --- W4401 ---
W4401_EDGE_BY_TYPE = {'рама': 1080, 'импост': 1080, 'створка': 1190}
W4401_EDGE_DEFAULT = 1080
W4401_MAX_STEP     = 3000
W4401_MIN_SPAN     = 1000
W4401_SHIFT        = 200
W4401_TOL          = 100
W4401_TAIL         = '00000'

# --- W1901 ---
W1901_TAIL_DEFAULT  = '00350'
W1901_TAIL_D25      = '00450'
W1901_HANDLE_OFFSET = 500

# --- W0101 ---
W0101_TAIL_FROM_0301 = '00000'
W0101_TAIL_FROM_1901 = '00600'

# --- W1101 ---
W1101_EDGE     = 1500
W1101_OFFSET   = 1000
W1101_L_LIMIT  = 10000
W1101_TAIL     = '00200'
W1101_IMPOST_L = 880
W1101_IMPOST_R = 820

# --- W0201 ---
W0201_EDGE        = 1020
W0201_TAIL_STD    = '00200'
W0201_TAIL_IMPOST = '00100'

# --- W1601 ---
W1601_EDGE    = 1820
W1601_L_LIMIT = 10000
W1601_TAIL    = '00200'

# --- W1801 ---
W1801_CENTER_OFFSET = 30
W1801_TAIL          = '00200'

# --- W0999 / W1804 ---
W_RUBBER_OFFSET = 380
W0999_TAIL      = '00000'
W1804_TAIL      = '00000'

# --- W4501 ---
W4501_BY_PROFILE = {
    '1012130010001': 550,
    '1012080010001': 550,
    '1012330010001': 577,
    '1032130010001': 720,
    '1032320010001': 730,
    '1032350010001': 700,
}
W4501_DEFAULT = 600
W4501_TAIL    = '00000'
W4501_SKIP_TYPES = {'импост'}


# ============================================================
# СБОРЩИКИ СТРОК
# ============================================================

def build_ksn(stab, length, profile_code, description, storage="000000"):
    return (
        f"KSN{stab:0>3}"
        f"L{length:0>5}"
        f"T{profile_code:<13}"
        f"B{description:<20}"
        f"O{storage:<6}"
    )


def build_ktn(part, cart, comp, order, pos, elem, location_stz,
              profile_code, description, length,
              g_val, d_val, s_val, w_val, i_val):
    return (
        f"KTN{part:0>4}"
        f"C{cart:0>3}"
        f"F{comp:0>3}"
        f"K{order:<8}"
        f"P{pos:<9}"
        f"E{elem:0>3}"
        f"U{location_stz}"
        f"T{profile_code:<13}"
        f"B{description:<20}"
        f"L{length:0>5}"
        f"G{g_val:<8}"
        f"D{d_val}"
        f"S{s_val}"
        f"W{w_val}"
        f"I{i_val:<18}"
        f"S      Z"
    )


def build_s(length):
    return f"S   /{S_CODE:<11}/{length:0>5}"


def compute_i(order, cart, comp, elem):
    try:
        comp_int = str(int(comp))
    except ValueError:
        comp_int = '0'
    return f"{order:<8}"[:8] + comp_int + f"{cart:0>3}" + "001" + f"{elem:0>3}"


# ============================================================
# УНИКАЛЬНЫЙ КЛЮЧ ЗАКАЗА (K)
# ============================================================

def get_order_key(user_key, source_key, src_filename):
    """
    Возвращает K-поле (8 знаков) с приоритетом:
      1. Пользовательский ключ
      2. Источник (если не 00000000)
      3. Имя исходного файла (alphanumeric)
    """
    # 1. Пользовательский
    if user_key and user_key.strip():
        clean = re.sub(r'[^A-Za-z0-9]', '', user_key.strip())[:8]
        if clean:
            return clean.ljust(8)[:8]

    # 2. Из источника
    sk = (source_key or '').strip()
    if sk and sk != '00000000':
        return sk.ljust(8)[:8]

    # 3. Из имени файла
    if src_filename:
        base = os.path.splitext(os.path.basename(src_filename))[0]
        clean = re.sub(r'[^A-Za-z0-9]', '', base)[:8]
        if clean:
            return clean.ljust(8)[:8]

    return '00000000'


# ============================================================
# ПАРСЕРЫ (принимают S/R/I)
# ============================================================

def parse_ksn(line):
    m = re.match(r'^KSN(\d{3})L(\d{5})T(\S{13})B(.{20})O(\S{6})', line)
    if m:
        return {
            'stab': m.group(1), 'length': m.group(2),
            'profile_code': m.group(3), 'description': m.group(4),
            'storage': m.group(5),
        }
    m = re.match(
        r'^KSN(\d{3})L(\d{5})T\s+(BR-\d{5})(?:#\s*(\S+))?\s+([A-Z])O(\d{6})',
        line
    )
    if m:
        profile_code = BR_TO_PROFILE.get(m.group(3))
        if not profile_code:
            return None
        desc = (m.group(4) or 'bel | bel').ljust(20)[:20]
        return {
            'stab': m.group(1), 'length': m.group(2),
            'profile_code': profile_code, 'description': desc,
            'storage': m.group(6),
        }
    return None


def parse_ktn(line):
    m = re.match(
        r'^KTN(\d{4})C(\d{3})F(\d{3})K(\S{8})P(\S{9})E(\d{3})'
        r'U([A-Z])'
        r'T(\S{13})B(.{20})L(\d{5})G(\S{8})D(\S)S(\S)W(\S)I(\S{18})',
        line
    )
    if m:
        return {
            'part': m.group(1), 'cart': m.group(2), 'comp': m.group(3),
            'order': m.group(4), 'pos': m.group(5), 'elem': m.group(6),
            'location': m.group(7),
            'profile_code': m.group(8), 'description': m.group(9),
            'length': m.group(10),
            'g': m.group(11), 'd': m.group(12), 's': m.group(13),
            'w': m.group(14), 'i': m.group(15),
            'source_format': 'sta',
        }
    m = re.match(
        r'^KTN(\d{4})C(\d{3})F(\d{3})K(\s{8}|\d{8})P(\S{9})E(\d{3})'
        r'U([A-Z])'
        r'T\s+(BR-\d{5})(?:#\s*(\S+))?'
        r'\s+([A-Z])L(\d{5})G(\S{8})D(\S)S(\S)',
        line
    )
    if m:
        profile_code = BR_TO_PROFILE.get(m.group(8))
        if not profile_code:
            return None
        order_raw = m.group(4).strip()
        order = order_raw if order_raw else '00000000'
        desc = (m.group(9) or 'bel | bel').ljust(20)[:20]
        return {
            'part': m.group(1), 'cart': m.group(2), 'comp': m.group(3),
            'order': order, 'pos': m.group(5), 'elem': m.group(6),
            'location': m.group(7),
            'profile_code': profile_code, 'description': desc,
            'length': m.group(11),
            'g': m.group(12),
            'd': '1' if m.group(13) == '0' else m.group(13),
            's': m.group(14),
            'w': '0', 'i': None,
            'source_format': 'test8',
        }
    return None


# ============================================================
# РАСЧЁТЫ ПОЗИЦИЙ
# ============================================================

def get_w4401_edge(profile_type):
    return W4401_EDGE_BY_TYPE.get(profile_type, W4401_EDGE_DEFAULT)


def compute_w4401_positions(length_x10, edge):
    if length_x10 is None or length_x10 < 2 * edge:
        return []
    span = length_x10 - 2 * edge
    if span < W4401_MIN_SPAN:
        return [length_x10 // 2]
    N = max(1, (span + W4401_MAX_STEP - 1) // W4401_MAX_STEP)
    step = span / N
    return [int(round(edge + i * step)) for i in range(N + 1)]


def compute_w4401_with_handle(length_x10, w1901_pos, edge):
    if length_x10 is None:
        return []
    left_end  = w1901_pos - W1901_HANDLE_OFFSET
    right_beg = w1901_pos + W1901_HANDLE_OFFSET
    right_end = length_x10 - edge
    result = []
    if left_end > edge:
        span = left_end - edge
        N = max(1, (span + W4401_MAX_STEP - 1) // W4401_MAX_STEP)
        step = span / N
        for i in range(N + 1):
            result.append(int(round(edge + i * step)))
    else:
        result.append(edge)
    if right_end > right_beg:
        span = right_end - right_beg
        N = max(1, (span + W4401_MAX_STEP - 1) // W4401_MAX_STEP)
        step = span / N
        for i in range(N + 1):
            result.append(int(round(right_beg + i * step)))
    else:
        result.append(right_beg)
    return sorted(set(result))


def filter_w4401_handle(positions, w1901_pos):
    if w1901_pos is None:
        return positions
    zs = w1901_pos - W1901_HANDLE_OFFSET
    ze = w1901_pos + W1901_HANDLE_OFFSET
    return [p for p in positions if not (zs < p < ze)]


def apply_w0301_shift(positions, w0301_positions):
    if not w0301_positions:
        return positions
    shifted = []
    for p in positions:
        s = 0
        for q in w0301_positions:
            if abs(p - q) <= W4401_TOL:
                s = W4401_SHIFT
                break
        shifted.append(p + s)
    return shifted


def compute_w1101_positions(length_x10, w0301_positions):
    positions = [W1101_EDGE, length_x10 - W1101_EDGE]
    if w0301_positions:
        for p in w0301_positions:
            positions.append(p - W1101_OFFSET)
            positions.append(p + W1101_OFFSET)
    else:
        if length_x10 > W1101_L_LIMIT:
            positions.append(length_x10 // 2)
    return sorted(set(p for p in positions if 0 < p < length_x10))


def compute_w0201_positions(length_x10):
    if length_x10 is None or length_x10 < 2 * W0201_EDGE:
        return []
    return [W0201_EDGE, length_x10 - W0201_EDGE]


def compute_w1601_positions(length_x10):
    if length_x10 is None or length_x10 < 2 * W1601_EDGE:
        return []
    positions = [W1601_EDGE, length_x10 - W1601_EDGE]
    if length_x10 > W1601_L_LIMIT:
        positions.append(length_x10 // 2)
    return sorted(set(positions))


# ============================================================
# ПРЕОБРАЗОВАНИЕ
# ============================================================

def transform_content(content: str, order_note: str = "",
                      user_order_key: str = "",
                      src_filename: str = "",
                      add_w4501: bool = False, log=print) -> str:
    note_DN  = NOTE_TRIGGER_DN  in (order_note or '')
    note_D25 = NOTE_TRIGGER_D25 in (order_note or '')
    if order_note:
        log(f"Примечание заказа: {order_note}")
        if note_DN:  log(f"  → триггер {NOTE_TRIGGER_DN}: рама B → W1801")
        if note_D25: log(f"  → триггер {NOTE_TRIGGER_D25}: W1901 хвост → {W1901_TAIL_D25}")
    log(f"Работа 4501: {'включена' if add_w4501 else 'отключена'}")

    # Вычисляем K-ключ один раз для всего файла
    # На вход подаём "какой-то источник" — фактически заменит пустые поля K в KTN
    first_source_k = '00000000'
    for ln in content.splitlines():
        if ln.startswith('KTN'):
            parsed = parse_ktn(ln.rstrip('\n').rstrip('\r'))
            if parsed:
                first_source_k = parsed['order']
                break
    order_key_final = get_order_key(user_order_key, first_source_k, src_filename)
    log(f"Ключ заказа K = '{order_key_final}'")

    lines = content.splitlines()
    out = []
    i = 0
    stab_counter = 0

    while i < len(lines):
        raw = lines[i].rstrip('\n').rstrip('\r')
        if not raw:
            i += 1
            continue

        if raw.strip() == SRC_VER:
            out.append(NEW_VER)
            i += 1
            continue

        # --- KSN ---
        if raw.startswith('KSN'):
            ksn = parse_ksn(raw)
            if ksn:
                stab_counter += 1
                ksn['stab'] = f"{stab_counter:03d}"
                out.append(build_ksn(
                    ksn['stab'], ksn['length'],
                    ksn['profile_code'], ksn['description'],
                    ksn['storage']
                ))
            else:
                log(f"[!] Не распознан KSN: {raw[:60]}")
            i += 1
            continue

        # --- KTN ---
        if raw.startswith('KTN'):
            ktn = parse_ktn(raw)
            if not ktn:
                log(f"[!] Не распознан KTN: {raw[:60]}")
                i += 1
                continue

            profile_code = ktn['profile_code']
            profile_type = PROFILE_TYPES.get(profile_code)
            length_x10 = int(ktn['length']) if ktn['length'] else None
            location = ktn['location']

            log(f"LOCATION = {location} ({LOCATION_MEANING.get(location, '?')})")

            # Сбор строк блока (W/S/пустые до KSN/KTN/KRL)
            i += 1
            w_source = []
            while i < len(lines):
                nxt = lines[i].rstrip('\n').rstrip('\r')
                if not nxt:
                    i += 1
                    continue
                if nxt.startswith('W'):
                    w_source.append(nxt)
                    i += 1
                elif nxt.startswith('S'):
                    i += 1
                elif (nxt.startswith('KTN') or nxt.startswith('KSN')
                      or nxt.startswith('KRL') or nxt.startswith('Ver')):
                    break
                else:
                    i += 1

            # Фильтр W-источника по длине детали
            if length_x10:
                filtered = []
                for w in w_source:
                    mw = re.match(r'W\d+/00/(\d+)/', w)
                    if mw:
                        if int(mw.group(1)) < length_x10:
                            filtered.append(w)
                    else:
                        filtered.append(w)
                w_source = filtered

            # Сбор данных из W
            w0301_positions = []
            w1901_position = None
            existing_0999 = set(); existing_1101 = set()
            existing_0201 = set(); existing_1601 = set()
            existing_1801 = set(); existing_1804 = set()
            for w in w_source:
                mw = re.match(r'W0301/00/(\d+)/', w)
                if mw: w0301_positions.append(int(mw.group(1)))
                mw = re.match(r'W1901/00/(\d+)/', w)
                if mw and w1901_position is None:
                    w1901_position = int(mw.group(1))
                mw = re.match(r'W0999/00/(\d+)/', w)
                if mw: existing_0999.add(int(mw.group(1)))
                mw = re.match(r'W1101/00/(\d+)/', w)
                if mw: existing_1101.add(int(mw.group(1)))
                mw = re.match(r'W0201/00/(\d+)/', w)
                if mw: existing_0201.add(int(mw.group(1)))
                mw = re.match(r'W1601/00/(\d+)/', w)
                if mw: existing_1601.add(int(mw.group(1)))
                mw = re.match(r'W1801/00/(\d+)/', w)
                if mw: existing_1801.add(int(mw.group(1)))
                mw = re.match(r'W1804/00/(\d+)/', w)
                if mw: existing_1804.add(int(mw.group(1)))

            # Устанавливаем финальный K (уникальный)
            ktn['order'] = order_key_final
            # I-поле всегда пересчитываем из актуального K
            ktn['i'] = compute_i(
                order_key_final, ktn['cart'], ktn['comp'], ktn['elem'])

            # Формирование W-элементов
            w_items = []
            for w in w_source:
                m1901_mod = re.match(r'W1901/00/(\d+)/\d+', w)
                if m1901_mod and note_D25:
                    pos = m1901_mod.group(1)
                    w = f"W1901/00/{pos}/{W1901_TAIL_D25}"

                mw = re.match(r'W(\d+)/00/(\d+)/', w)
                if mw:
                    w_items.append((int(mw.group(2)), mw.group(1), w))

                m0301 = re.match(r'W0301/00/(\d+)/', w)
                if m0301:
                    pos = int(m0301.group(1))
                    w_items.append((pos, '0101',
                        f"W0101/00/{pos:05d}/{W0101_TAIL_FROM_0301}"))

                m1901 = re.match(r'W1901/00/(\d+)/', w)
                if m1901:
                    pos = int(m1901.group(1))
                    w_items.append((pos, '0101',
                        f"W0101/00/{pos:05d}/{W0101_TAIL_FROM_1901}"))

            # W4501
            if (add_w4501 and length_x10
                    and profile_type not in W4501_SKIP_TYPES):
                w4501_pos = W4501_BY_PROFILE.get(profile_code, W4501_DEFAULT)
                if w4501_pos < length_x10:
                    w_items.append((w4501_pos, '4501',
                                    f"W4501/00/{w4501_pos:05d}/{W4501_TAIL}"))

            # W1101
            if length_x10:
                if profile_type == 'импост' and location == 'H':
                    for p in (W1101_IMPOST_L, length_x10 - W1101_IMPOST_R):
                        if 0 < p < length_x10 and p not in existing_1101:
                            w_items.append((p, '1101',
                                f"W1101/00/{p:05d}/{W1101_TAIL}"))
                            existing_1101.add(p)
                elif profile_type == 'рама' and location == 'B':
                    for p in compute_w1101_positions(length_x10, w0301_positions):
                        if 0 < p < length_x10 and p not in existing_1101:
                            w_items.append((p, '1101',
                                f"W1101/00/{p:05d}/{W1101_TAIL}"))
                            existing_1101.add(p)

            # W0201 / W1801
            add_w0201 = (
                (profile_type in ('створка', 'рама') and location in ('T', 'B')) or
                (profile_type == 'импост' and location == 'V')
            )
            frame_B_DN = (profile_type == 'рама' and location == 'B' and note_DN)

            if length_x10 and add_w0201:
                if frame_B_DN:
                    for p in compute_w0201_positions(length_x10):
                        if p not in existing_1801:
                            w_items.append((p, '1801',
                                f"W1801/00/{p:05d}/{W1801_TAIL}"))
                            existing_1801.add(p)
                else:
                    tail = W0201_TAIL_IMPOST if profile_type == 'импост' \
                           else W0201_TAIL_STD
                    for p in compute_w0201_positions(length_x10):
                        if p not in existing_0201:
                            w_items.append((p, '0201',
                                f"W0201/00/{p:05d}/{tail}"))
                            existing_0201.add(p)

            # W1601
            if length_x10 and (
                (profile_type == 'створка' and location in ('T', 'B')) or
                (profile_type == 'рама'    and location == 'B')
            ):
                for p in compute_w1601_positions(length_x10):
                    if p not in existing_1601:
                        w_items.append((p, '1601',
                            f"W1601/00/{p:05d}/{W1601_TAIL}"))
                        existing_1601.add(p)

            # W1801 (импост H)
            if length_x10 and profile_type == 'импост' and location == 'H':
                p = length_x10 // 2 + W1801_CENTER_OFFSET
                if 0 < p < length_x10 and p not in existing_1801:
                    w_items.append((p, '1801',
                        f"W1801/00/{p:05d}/{W1801_TAIL}"))
                    existing_1801.add(p)

            # W0999 / W1804
            if length_x10 and w0301_positions:
                for p in w0301_positions:
                    for delta in (-W_RUBBER_OFFSET, W_RUBBER_OFFSET):
                        new_p = p + delta
                        if not (0 < new_p < length_x10):
                            continue
                        if profile_type in ('рама', 'импост'):
                            if new_p not in existing_0999:
                                w_items.append((new_p, '0999',
                                    f"W0999/00/{new_p:05d}/{W0999_TAIL}"))
                                existing_0999.add(new_p)
                        elif profile_type == 'створка':
                            if new_p not in existing_1804:
                                w_items.append((new_p, '1804',
                                    f"W1804/00/{new_p:05d}/{W1804_TAIL}"))
                                existing_1804.add(new_p)

            # W4401
            edge = get_w4401_edge(profile_type)
            if length_x10:
                if w1901_position is not None:
                    positions = compute_w4401_with_handle(
                        length_x10, w1901_position, edge)
                else:
                    positions = compute_w4401_positions(length_x10, edge)
                positions = filter_w4401_handle(positions, w1901_position)
                positions = apply_w0301_shift(positions, w0301_positions)
                for p in positions:
                    if 0 < p < length_x10:
                        w_items.append((p, '4401',
                                        f"W4401/00/{p:05d}/{W4401_TAIL}"))

            # KTN
            loc_stz = LOC_TO_STURTZ.get(location, 'L')
            out.append(build_ktn(
                ktn['part'], ktn['cart'], ktn['comp'],
                ktn['order'], ktn['pos'], ktn['elem'],
                loc_stz,
                ktn['profile_code'], ktn['description'], ktn['length'],
                ktn['g'], ktn['d'], ktn['s'], ktn['w'], ktn['i']
            ))

            # W-строки
            w_items.sort(key=lambda x: (x[0], x[1]))
            for _, _, w_line in w_items:
                out.append(w_line)

            # S-строка
            out.append(build_s(ktn['length']))
            continue

        # S-строки вне KTN игнорируем
        if raw.startswith('S'):
            i += 1
            continue

        out.append(raw)
        i += 1

    return '\n'.join(out) + '\n'


# ============================================================
# ТАБЛИЦА ПРАВИЛ
# ============================================================

RULES_TABLE = [
    ("0101", "Фреза 12 мм", "00000 / 00600", "все", "любая",
     "От W0301 → 00000; от W1901 → 00600"),
    ("0201", "Дренаж наружу вниз", "00200 / 00100",
     "створка, рама, импост", "створка T/B, рама T/B, импост V",
     "102 мм от краёв"),
    ("0301", "Разметка T пластик", "—", "все", "любая", "из источника"),
    ("0999", "Фрез. под резину", "00000", "рама, импост", "любая",
     "W0301 ± 380"),
    ("1101", "Дренаж фальца 45°", "00200", "рама, импост",
     "рама B, импост H", "рама: 150 + ±100 от W0301; импост H: 88/82"),
    ("1331", "Разметка T металл", "00000", "все", "любая", "из источника"),
    ("1601", "Дренаж фальца для створки", "00200", "створка, рама",
     "створка T/B, рама B", "182 мм + центр при L>1000"),
    ("1801", "Дренаж наружу", "00200", "импост, рама",
     "импост H; рама B (при 'DN')", "импост: L/2+30; рама B: 102 мм"),
    ("1804", "Фрез. под резину", "00000", "створка", "любая",
     "W0301 ± 380"),
    ("1901", "Фрез. 3 отв. под ручку", "00350 / 00450", "створка", "любая",
     "L/2. При 'D25' → 00450"),
    ("4401", "Шуруповёрт", "00000", "все", "любая",
     "108/119 мм от краёв, шаг ≤300, сегментация вокруг W1901±50"),
    ("4501", "Сверление под крепёж", "00000", "рама, створка", "любая",
     "Фикс. по профилю. По умолчанию отключена"),
]

RULES_HEADERS = ["Код", "Назначение", "Хвост", "Профили", "LOCATION", "Правило"]


def export_rules_csv(path: str):
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(RULES_HEADERS)
        for row in RULES_TABLE:
            writer.writerow(row)


def rules_to_tsv() -> str:
    lines = ['\t'.join(RULES_HEADERS)]
    for row in RULES_TABLE:
        lines.append('\t'.join(row))
    return '\n'.join(lines)


# ============================================================
# GUI
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("STBAZ-IPT-4000 — конвертер .sob")
        self.geometry("940x720")
        self.resizable(True, True)

        self.src_var      = tk.StringVar()
        self.dst_var      = tk.StringVar()
        self.note_var     = tk.StringVar()
        self.order_key_var = tk.StringVar()
        self.add_4501_var = tk.BooleanVar(value=False)
        self.location_var = tk.StringVar(value="LOCATION: —")

        pad = {'padx': 8, 'pady': 4}

        tk.Label(self, text="Файл-источник (.sob):").grid(
            row=0, column=0, sticky='w', **pad)
        tk.Entry(self, textvariable=self.src_var, width=74).grid(
            row=0, column=1, sticky='we', **pad)
        tk.Button(self, text="Обзор...", command=self.choose_src).grid(
            row=0, column=2, **pad)

        tk.Label(self, text="Результирующий файл (.sob):").grid(
            row=1, column=0, sticky='w', **pad)
        tk.Entry(self, textvariable=self.dst_var, width=74).grid(
            row=1, column=1, sticky='we', **pad)
        tk.Button(self, text="Обзор...", command=self.choose_dst).grid(
            row=1, column=2, **pad)

        tk.Label(self, text="Код заказа (K, до 8 знаков):").grid(
            row=2, column=0, sticky='w', **pad)
        tk.Entry(self, textvariable=self.order_key_var, width=74).grid(
            row=2, column=1, sticky='we', **pad)
        tk.Label(self,
                 text="(пусто → берётся из источника; если тоже пусто → из имени файла)",
                 fg="#666", font=("Arial", 8)).grid(
            row=3, column=1, sticky='w', **pad)

        tk.Label(self, text="Примечание заказа:").grid(
            row=4, column=0, sticky='w', **pad)
        tk.Entry(self, textvariable=self.note_var, width=74).grid(
            row=4, column=1, sticky='we', **pad)

        tk.Label(self,
                 text="(триггеры: DN → W1801 на раме B; D25 → W1901 хвост 00450)",
                 fg="#666", font=("Arial", 8)).grid(
            row=5, column=1, sticky='w', **pad)

        tk.Checkbutton(self, text="Добавить работу 4501 (по умолчанию отключена)",
                       variable=self.add_4501_var).grid(
            row=6, column=1, sticky='w', **pad)

        tk.Label(self, textvariable=self.location_var,
                 font=("Arial", 11, "bold"), fg="#1565C0").grid(
            row=7, column=1, sticky='w', **pad)

        btn_frame = tk.Frame(self)
        btn_frame.grid(row=8, column=1, columnspan=2, sticky='e', **pad)

        tk.Button(btn_frame, text="Сгенерировать выгрузку",
                  command=self.run_generation,
                  bg="#2E7D32", fg="white",
                  font=("Arial", 10, "bold")).pack(side='left', padx=4)
        tk.Button(btn_frame, text="Показать таблицу правил",
                  command=self.show_rules).pack(side='left', padx=4)
        tk.Button(btn_frame, text="Сохранить правила в Excel (CSV)",
                  command=self.save_rules_csv).pack(side='left', padx=4)
        tk.Button(btn_frame, text="Копировать правила (для Excel)",
                  command=self.copy_rules).pack(side='left', padx=4)

        tk.Label(self, text="Журнал:").grid(
            row=9, column=0, sticky='nw', **pad)
        self.log_text = tk.Text(self, height=20, wrap='word')
        self.log_text.grid(row=9, column=1, columnspan=2,
                           sticky='nsew', **pad)
        scroll = ttk.Scrollbar(self, command=self.log_text.yview)
        scroll.grid(row=9, column=3, sticky='ns')
        self.log_text.config(yscrollcommand=scroll.set)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(9, weight=1)

    def log(self, msg):
        self.log_text.insert('end', msg + '\n')
        self.log_text.see('end')
        self.update_idletasks()

    def _log_location(self, msg):
        if msg.startswith("LOCATION"):
            self.location_var.set(msg)
        self.log(msg)

    def choose_src(self):
        path = filedialog.askopenfilename(
            title="Выберите файл-источник",
            filetypes=[("Файлы центра (.sob)", "*.sob"),
                       ("Все файлы", "*.*")])
        if not path:
            return
        self.src_var.set(path)
        folder = os.path.dirname(path)
        base = os.path.splitext(os.path.basename(path))[0]
        self.dst_var.set(os.path.join(folder, f"{base}_out{DST_EXT}"))

    def choose_dst(self):
        path = filedialog.asksaveasfilename(
            title="Сохранить результат как",
            defaultextension=DST_EXT,
            filetypes=[("Файлы центра (.sob)", "*.sob"),
                       ("Все файлы", "*.*")],
            initialfile=os.path.basename(self.dst_var.get()) or f"output{DST_EXT}",
            initialdir=os.path.dirname(self.dst_var.get()) or None)
        if path:
            if not path.lower().endswith(DST_EXT):
                path += DST_EXT
            self.dst_var.set(path)

    def run_generation(self):
        self.log_text.delete('1.0', 'end')
        self.location_var.set("LOCATION: —")

        src  = self.src_var.get().strip()
        dst  = self.dst_var.get().strip()
        note = self.note_var.get().strip()
        okey = self.order_key_var.get().strip()
        add_4501 = bool(self.add_4501_var.get())

        if not src or not os.path.isfile(src):
            messagebox.showerror("Ошибка", "Файл-источник не найден.")
            return
        if not dst:
            messagebox.showerror("Ошибка", "Не указан путь для результата.")
            return
        if not dst.lower().endswith(DST_EXT):
            dst += DST_EXT
            self.dst_var.set(dst)

        try:
            self.log(f"Источник: {src}")
            self.log(f"Результат: {dst}")
            with open(src, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()

            self.log("Преобразование...")
            result = transform_content(content,
                                       order_note=note,
                                       user_order_key=okey,
                                       src_filename=src,
                                       add_w4501=add_4501,
                                       log=self._log_location)

            with open(dst, 'w', encoding='utf-8', newline='\r\n') as f:
                f.write(result)

            self.log("Готово.")
            messagebox.showinfo("Успех", f"Файл сохранён:\n{dst}")
        except Exception as e:
            self.log(f"[ОШИБКА] {e}")
            messagebox.showerror("Ошибка", str(e))

    def show_rules(self):
        win = tk.Toplevel(self)
        win.title("Таблица правил STBAZ-IPT-4000")
        win.geometry("1200x520")
        cols = RULES_HEADERS
        tree = ttk.Treeview(win, columns=cols, show='headings', height=20)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=140, anchor='w')
        tree.column("Правило", width=440)
        for row in RULES_TABLE:
            tree.insert('', 'end', values=row)
        vsb = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(win, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)
        tk.Button(win, text="Копировать всё (для вставки в Excel)",
                  command=self.copy_rules).grid(row=2, column=0, pady=6)
        tk.Button(win, text="Сохранить как CSV",
                  command=self.save_rules_csv).grid(row=2, column=1, pady=6)

    def copy_rules(self):
        self.clipboard_clear()
        self.clipboard_append(rules_to_tsv())
        self.log("Таблица правил скопирована в буфер обмена.")
        messagebox.showinfo("Скопировано", "Вставьте в Excel (Ctrl+V).")

    def save_rules_csv(self):
        path = filedialog.asksaveasfilename(
            title="Сохранить таблицу правил",
            defaultextension=".csv",
            filetypes=[("CSV (Excel)", "*.csv"), ("Все файлы", "*.*")],
            initialfile="STBAZ_rules.csv")
        if not path:
            return
        try:
            export_rules_csv(path)
            self.log(f"Таблица правил сохранена: {path}")
            messagebox.showinfo("Готово", f"Файл сохранён:\n{path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


if __name__ == '__main__':
    App().mainloop()
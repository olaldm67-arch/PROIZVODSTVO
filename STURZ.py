import re
import os
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
    '1021030010001': 'импост',
    '1022140010001': 'импост',
    '1032130010001': 'створка',
    '1032320010001': 'створка',
}

NEW_VER   = 'Ver02.07-00'
SRC_VER   = 'ZUH30'
S_CODE    = '113292'
S_FIELD_W = 11

SRC_EXT = '.sob'
DST_EXT = '.sob'

LOCATION_MEANING = {
    'L': 'лево',
    'R': 'право',
    'T': 'верх',
    'B': 'низ',
    'V': 'вертикальный',
    'H': 'горизонтальный',
}

# --- W4401 (шуруповерт) ---
W4401_EDGE  = 600
W4401_STEP  = 2000
W4401_SHIFT = 200
W4401_TOL   = 100
W4401_TAIL  = '00000'

# --- W0101 (фреза 12мм) ---
W0101_TAIL  = '00000'

# --- W1101 (дренаж фальца 45°) ---
W1101_EDGE    = 1500     # 150.0 мм от края — всегда
W1101_OFFSET  = 1000     # 100.0 мм в обе стороны от W0301
W1101_L_LIMIT = 10000    # 1000.0 мм — порог для центрирования
W1101_TAIL    = '00200'


# ============================================================
# РАСЧЁТЫ ПОЗИЦИЙ
# ============================================================

def compute_w4401_positions(length_x10):
    if length_x10 is None or length_x10 < 2 * W4401_EDGE:
        return []
    last = length_x10 - W4401_EDGE
    positions = [W4401_EDGE]
    pos = W4401_EDGE + W4401_STEP
    while pos < last:
        positions.append(pos)
        pos += W4401_STEP
    if positions[-1] != last:
        positions.append(last)
    return positions


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
    """
    W1101:
    - ВСЕГДА: по 150 мм от краёв
    - Если есть W0301: дополнительно ±100 мм от каждой W0301
    - Если нет W0301 и L>1000 мм: дополнительно по центру
    """
    positions = [
        W1101_EDGE,
        length_x10 - W1101_EDGE,
    ]
    if w0301_positions:
        for p in w0301_positions:
            positions.append(p - W1101_OFFSET)
            positions.append(p + W1101_OFFSET)
    else:
        if length_x10 > W1101_L_LIMIT:
            positions.append(length_x10 // 2)
    positions = sorted(set(p for p in positions if 0 < p < length_x10))
    return positions


# ============================================================
# ПРЕОБРАЗОВАНИЕ
# ============================================================

def transform_content(content: str, log=print) -> str:
    out_lines = []
    lines = content.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()

        if line == SRC_VER:
            out_lines.append(NEW_VER)
            i += 1
            continue

        if not line:
            i += 1
            continue

        # --- KSN ---
        if line.startswith('KSN'):
            for br, code in BR_TO_PROFILE.items():
                line = line.replace(br, code)
            line = re.sub(r'S(O\d{6})', r'\1', line)
            out_lines.append(line)
            i += 1
            continue

        # --- KTN + W ---
        if line.startswith('KTN'):
            profile_code = None
            for br, code in BR_TO_PROFILE.items():
                if br in line:
                    profile_code = code
                    line = line.replace(br, code)
            line = line.replace('D0', 'D1')

            m = re.search(r'E\d+U([A-Z])', line)
            location = m.group(1) if m else None
            if location:
                meaning = LOCATION_MEANING.get(location, '?')
                log(f"LOCATION = {location} ({meaning})")

            lm = re.search(r'L(\d+)', line)
            length_x10 = int(lm.group(1)) if lm else None

            out_lines.append(line)
            i += 1

            w_source = []
            while i < len(lines) and lines[i].startswith('W'):
                w_source.append(lines[i].rstrip())
                i += 1

            w0301_positions = []
            for w in w_source:
                mw = re.match(r'W0301/00/(\d+)/', w)
                if mw:
                    w0301_positions.append(int(mw.group(1)))

            # Источник + W0101 после каждого W0301
            for w in w_source:
                out_lines.append(w)
                mw = re.match(r'W0301/00/(\d+)/', w)
                if mw:
                    pos = mw.group(1)
                    out_lines.append(f"W0101/00/{pos}/{W0101_TAIL}")

            # --- W1101 ---
            profile_type = PROFILE_TYPES.get(profile_code)
            add_w1101 = (
                (profile_type == 'рама'   and location == 'B') or
                (profile_type == 'импост' and location == 'H')
            )
            if add_w1101 and length_x10:
                w1101_positions = compute_w1101_positions(
                    length_x10, w0301_positions)
                existing = set()
                for w in w_source:
                    mw = re.match(r'W1101/00/(\d+)/', w)
                    if mw:
                        existing.add(int(mw.group(1)))
                for p in w1101_positions:
                    if p not in existing:
                        out_lines.append(
                            f"W1101/00/{p:05d}/{W1101_TAIL}")
                        existing.add(p)

            # --- W4401 ---
            positions = compute_w4401_positions(length_x10)
            positions = apply_w0301_shift(positions, w0301_positions)
            for p in positions:
                out_lines.append(f"W4401/00/{p:05d}/{W4401_TAIL}")

            continue

        # --- S ---
        if line.startswith('S'):
            first = line.find('/')
            last = line.rfind('/')
            if first != -1 and last != -1 and first < last:
                code_field = S_CODE.ljust(S_FIELD_W)
                line = line[:first] + '/' + code_field + line[last:]
            out_lines.append(line)
            i += 1
            continue

        out_lines.append(line)
        i += 1

    return '\n'.join(out_lines) + '\n'


# ============================================================
# GUI
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("STBAZ-IPT-4000 — конвертер .sob")
        self.geometry("860x560")
        self.resizable(True, True)

        self.src_var = tk.StringVar()
        self.dst_var = tk.StringVar()
        self.location_var = tk.StringVar(value="LOCATION: —")

        pad = {'padx': 8, 'pady': 4}

        tk.Label(self, text="Файл-источник (.sob):").grid(
            row=0, column=0, sticky='w', **pad)
        tk.Entry(self, textvariable=self.src_var, width=72).grid(
            row=0, column=1, sticky='we', **pad)
        tk.Button(self, text="Обзор...", command=self.choose_src).grid(
            row=0, column=2, **pad)

        tk.Label(self, text="Результирующий файл (.sob):").grid(
            row=1, column=0, sticky='w', **pad)
        tk.Entry(self, textvariable=self.dst_var, width=72).grid(
            row=1, column=1, sticky='we', **pad)
        tk.Button(self, text="Обзор...", command=self.choose_dst).grid(
            row=1, column=2, **pad)

        tk.Label(self, textvariable=self.location_var,
                 font=("Arial", 11, "bold"), fg="#1565C0").grid(
            row=2, column=1, sticky='w', **pad)

        tk.Button(self, text="Сгенерировать выгрузку",
                  command=self.run_generation,
                  bg="#2E7D32", fg="white", height=1,
                  font=("Arial", 10, "bold")).grid(
            row=3, column=1, sticky='e', **pad)

        tk.Label(self, text="Журнал:").grid(
            row=4, column=0, sticky='nw', **pad)
        self.log_text = tk.Text(self, height=20, wrap='word')
        self.log_text.grid(row=4, column=1, columnspan=2,
                           sticky='nsew', **pad)
        scroll = ttk.Scrollbar(self, command=self.log_text.yview)
        scroll.grid(row=4, column=3, sticky='ns')
        self.log_text.config(yscrollcommand=scroll.set)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(4, weight=1)

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

        src = self.src_var.get().strip()
        dst = self.dst_var.get().strip()

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
            self.log("Чтение...")
            with open(src, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()

            self.log("Преобразование...")
            result = transform_content(content, log=self._log_location)

            with open(dst, 'w', encoding='utf-8') as f:
                f.write(result)

            self.log("Готово.")
            messagebox.showinfo("Успех", f"Файл сохранён:\n{dst}")
        except Exception as e:
            self.log(f"[ОШИБКА] {e}")
            messagebox.showerror("Ошибка", str(e))


# ============================================================
# ТОЧКА ВХОДА
# ============================================================

if __name__ == '__main__':
    App().mainloop()
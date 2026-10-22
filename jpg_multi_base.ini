import os
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from dbfread import DBF

# --- НАСТРОЙКА ПУТЕЙ ---
SOURCE_DIR = r"\\sp-roshan-03\Ccalc6.80\Data"
DEFAULT_OUTPUT_DIR = r"\\Synas\work\ПВХ\JPG"
# -----------------------

BASES = ["base4", "base6", "base7", "base10", "base15"]


def extract_jpg(base_name, output_dir, log, set_status, root):
    dbf_file = os.path.join(SOURCE_DIR, f"{base_name}.dbf")
    dbt_file = os.path.join(SOURCE_DIR, f"{base_name}.dbt")

    if not os.path.exists(dbf_file) or not os.path.exists(dbt_file):
        set_status("Ошибка: файлы не найдены!")
        log(f"Ошибка: {dbf_file} или {dbt_file} не найдены")
        return

    try:
        os.makedirs(output_dir, exist_ok=True)
    except Exception as e:
        set_status(f"Ошибка создания папки: {e}")
        log(f"Не удалось создать папку выгрузки: {e}")
        return

    set_status("Чтение DBF...")
    log(f"База: {base_name}")
    log(f"Источник: {SOURCE_DIR}")
    log(f"Выгрузка: {output_dir}")

    try:
        table = DBF(dbf_file, encoding='cp1251', raw=True, load=True)
    except Exception as e:
        set_status(f"Ошибка чтения DBF: {e}")
        log(f"Ошибка чтения DBF: {e}")
        return

    picture_field = None
    for field in table.fields:
        if 'PICT' in field.name.upper():
            picture_field = field.name
            break
    if not picture_field:
        set_status("Ошибка: поле PICTURE не найдено!")
        log("Поле PICTURE не найдено!")
        return

    field_names = [f.name for f in table.fields]
    use_zaknum = 'ZAKNUM' in field_names
    use_konnum = 'KONNUM' in field_names

    records_data = []
    for record in table.records:
        raw = record.get(picture_field)
        if raw is None:
            continue
        try:
            s = raw.decode('cp1251').strip('\x00')
            if not s.isdigit():
                continue
            block_num = int(s)
        except:
            continue

        parts = []
        if use_zaknum:
            val = record.get('ZAKNUM')
            if val is not None:
                try:
                    s_val = val.decode('cp866', errors='ignore').strip('\x00').strip()
                except:
                    s_val = str(val)
                if s_val:
                    parts.append(s_val)
        if use_konnum:
            val = record.get('KONNUM')
            if val is not None:
                try:
                    s_val = val.decode('cp866', errors='ignore').strip('\x00').strip()
                except:
                    s_val = str(val)
                if s_val:
                    parts.append(f"({s_val})")
        base_fn = " ".join(parts) if parts else f"image_{block_num}"

        for ch in '<>:"/\\|?*':
            base_fn = base_fn.replace(ch, '_')
        base_fn = ' '.join(base_fn.split())
        if not base_fn:
            base_fn = f"image_{block_num}"

        records_data.append((block_num, base_fn))

    if not records_data:
        set_status("Нет данных в поле PICTURE.")
        log("Нет данных в поле PICTURE.")
        return

    records_data.sort(key=lambda x: x[0])

    with open(dbt_file, 'rb') as f:
        dbt_data = f.read()
    dbt_size = len(dbt_data)

    log(f"Найдено записей с изображениями: {len(records_data)}")
    log(f"Размер DBT: {dbt_size} байт")
    log("-" * 40)

    set_status("Извлечение начато")
    log("=== НАЧАЛО ИЗВЛЕЧЕНИЯ ===")

    saved = 0
    total = len(records_data)

    for i, (block_num, base_fn) in enumerate(records_data):
        start = (block_num - 1) * 512
        if start >= dbt_size:
            continue

        if i + 1 < len(records_data):
            end = (records_data[i+1][0] - 1) * 512
        else:
            end = dbt_size

        if end <= start:
            continue

        segment = dbt_data[start:end]
        if len(segment) < 100:
            continue

        pos = segment.find(b'\xff\xd8')
        if pos == -1:
            continue

        jpeg = segment[pos:]
        end_jpeg = jpeg.find(b'\xff\xd9')
        if end_jpeg != -1:
            jpeg = jpeg[:end_jpeg+2]

        fname = f"{base_fn}.jpg"
        fpath = os.path.join(output_dir, fname)

        # Перезаписываем файл, если он существует
        with open(fpath, 'wb') as f:
            f.write(jpeg)
        saved += 1

        set_status(f"Сохранён {saved} из {total}")
        log(f"[{saved}/{total}] {os.path.basename(fpath)} (блок {block_num}, {len(jpeg)} байт)")

    set_status("Извлечение завершено")
    log("=== КОНЕЦ ИЗВЛЕЧЕНИЯ ===")
    log(f"Готово. Сохранено {saved} JPG-файлов в {output_dir}.")
    log("Окно закроется автоматически через 10 секунд...")

    # Автозакрытие через 10 секунд
    root.after(10000, root.destroy)


# --- Создаём окно ---
root = tk.Tk()
root.title("Извлечение JPG из DBF/DBT")
root.geometry("650x480")
root.resizable(True, True)

# --- Панель выбора ---
top_frame = tk.Frame(root)
top_frame.pack(fill=tk.X, padx=10, pady=8)

tk.Label(top_frame, text="База:", font=("Arial", 10)).pack(side=tk.LEFT)

base_var = tk.StringVar(value=BASES[0])
base_combo = ttk.Combobox(top_frame, textvariable=base_var, values=BASES,
                          state="readonly", width=12)
base_combo.pack(side=tk.LEFT, padx=5)

out_var = tk.StringVar(value=DEFAULT_OUTPUT_DIR)


def choose_output():
    folder = filedialog.askdirectory(initialdir=out_var.get(), title="Выберите папку для выгрузки JPG")
    if folder:
        out_var.set(folder)


tk.Button(top_frame, text="Папка выгрузки...", command=choose_output).pack(side=tk.LEFT, padx=5)

# --- Строка статуса ---
status_var = tk.StringVar()
status_var.set("Готов к работе. Выберите базу и нажмите «Старт».")

status_label = tk.Label(root, textvariable=status_var, font=("Arial", 11, "bold"), fg="blue")
status_label.pack(pady=5)

# --- Окно лога ---
log_text = scrolledtext.ScrolledText(root, height=18, wrap=tk.WORD)
log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)


def log(msg):
    log_text.insert(tk.END, msg + "\n")
    log_text.see(tk.END)
    root.update()


def set_status(msg):
    status_var.set(msg)
    root.update()


# --- Кнопка Старт ---
def on_start():
    start_btn.config(state=tk.DISABLED)
    base_combo.config(state=tk.DISABLED)
    log_text.delete("1.0", tk.END)
    extract_jpg(base_var.get(), out_var.get(), log, set_status, root)


start_btn = tk.Button(root, text="СТАРТ", font=("Arial", 12, "bold"),
                      bg="#4CAF50", fg="white", height=1, width=15,
                      command=on_start)
start_btn.pack(pady=8)

root.mainloop()
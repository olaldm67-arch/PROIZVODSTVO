import os
from dbfread import DBF

# --- НАСТРОЙКА ПУТЕЙ ---
SOURCE_DIR = r"\\sp-roshan-03\Ccalc6.80\Data"
OUTPUT_DIR = r"\\Synas\work\ПВХ\JPG"

dbf_file = os.path.join(SOURCE_DIR, "base4.dbf")
dbt_file = os.path.join(SOURCE_DIR, "base4.dbt")
# -----------------------

if not os.path.exists(dbf_file) or not os.path.exists(dbt_file):
    print(f"Ошибка: файлы не найдены в {SOURCE_DIR}")
    exit()

# Создаём папку выгрузки, если её нет
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Читаем DBF в raw-режиме
table = DBF(dbf_file, encoding='cp1251', raw=True, load=True)

# Находим поле PICTURE
picture_field = None
for field in table.fields:
    if 'PICT' in field.name.upper():
        picture_field = field.name
        break
if not picture_field:
    print("Поле PICTURE не найдено!")
    exit()

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
    base_name = " ".join(parts) if parts else f"image_{block_num}"

    for ch in '<>:"/\\|?*':
        base_name = base_name.replace(ch, '_')
    base_name = ' '.join(base_name.split())
    if not base_name:
        base_name = f"image_{block_num}"

    records_data.append((block_num, base_name))

if not records_data:
    print("Нет данных в поле PICTURE.")
    exit()

print(f"Найдено записей с изображениями: {len(records_data)}")

records_data.sort(key=lambda x: x[0])

with open(dbt_file, 'rb') as f:
    dbt_data = f.read()
dbt_size = len(dbt_data)
print(f"Размер DBT: {dbt_size} байт")

out_dir = OUTPUT_DIR
saved = 0

for i, (block_num, base_name) in enumerate(records_data):
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

    fname = f"{base_name}.jpg"
    fpath = os.path.join(out_dir, fname)

    # Перезаписываем файл, если он существует (без создания копий)
    with open(fpath, 'wb') as f:
        f.write(jpeg)
    saved += 1
    print(f"Сохранён {os.path.basename(fpath)} (блок {block_num}, размер {len(jpeg)} байт)")

print(f"\nГотово. Сохранено {saved} JPG-файлов в {OUTPUT_DIR}.")
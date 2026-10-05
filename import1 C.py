import xlrd
import openpyxl
from openpyxl.styles import Font
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import csv
import datetime
import traceback
from collections import defaultdict

try:
    import dbf
    DBF_AVAILABLE = True
except ImportError:
    DBF_AVAILABLE = False


def write_log(log_path, message):
    with open(log_path, 'a', encoding='utf-8') as f:
        f.write(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")


def safe_float(value, default=0.0):
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    if isinstance(value, str):
        val = value.strip().replace(',', '.')
        if val == '':
            return round(default, 2)
        try:
            return round(float(val), 2)
        except ValueError:
            return round(default, 2)
    return round(default, 2)


def read_excel_file(input_path):
    """
    Читает Excel-файл (xls или xlsx) и возвращает лист (sheet) в виде списка списков.
    """
    ext = os.path.splitext(input_path)[1].lower()
    if ext == '.xls':
        workbook = xlrd.open_workbook(input_path, formatting_info=False)
        sheet = workbook.sheet_by_index(0)
        data = []
        for row_idx in range(sheet.nrows):
            data.append([str(cell).strip() for cell in sheet.row_values(row_idx)])
        return data
    elif ext == '.xlsx':
        workbook = openpyxl.load_workbook(input_path, data_only=True)
        sheet = workbook.worksheets[0]
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append([str(cell).strip() if cell is not None else '' for cell in row])
        return data
    else:
        raise ValueError("Неподдерживаемый формат файла. Используйте .xls или .xlsx.")


def create_or_open_dbf(dbf_path, log_path):
    if not DBF_AVAILABLE:
        raise RuntimeError("Библиотека 'dbf' не установлена.")

    if not os.path.exists(dbf_path):
        write_log(log_path, f"Создаём новый DBF: {dbf_path}")
        table = dbf.Table(
            dbf_path,
            'KONTRAG C(255); NAKLAD C(255); NAIMEN C(255); DDATE D; ARTIKUL C(100); NDS N(15,2); KOLVO N(15,2); CENA N(15,2)',
            codepage='cp866'
        )
        table.open(mode=dbf.READ_WRITE)
        return table
    else:
        write_log(log_path, f"Открываем существующий DBF: {dbf_path}")
        table = dbf.Table(dbf_path, codepage='cp866')
        table.open(mode=dbf.READ_WRITE)
        try:
            write_log(log_path, "Выполняем pack() для удаления помеченных записей...")
            table.pack()
            write_log(log_path, "pack() выполнен.")
        except Exception as e:
            write_log(log_path, f"pack() не удался: {e}, продолжаем без удаления.")
        return table


def get_existing_invoices(table, log_path):
    invoices = set()
    for record in table:
        try:
            if hasattr(record, 'deleted') and record.deleted:
                continue
        except Exception:
            pass
        invoices.add(str(record.NAKLAD).strip())
    write_log(log_path, f"Найдено {len(invoices)} активных накладных.")
    return invoices


def append_new_records(table, grouped_rows, existing_invoices, log_path):
    added = 0
    skipped = 0
    for invoice, rows in grouped_rows.items():
        if invoice in existing_invoices:
            skipped += len(rows)
            continue
        for row in rows:
            try:
                date_str = row[2]
                if ' ' in date_str:
                    date_part = date_str.split(' ')[0]
                else:
                    date_part = date_str
                day, month, year = date_part.split('.')
                dt = datetime.date(int(year), int(month), int(day))
            except Exception as e:
                write_log(log_path, f"Ошибка даты '{row[2]}': {e}, ставим сегодня.")
                dt = datetime.date.today()

            table.append({
                'KONTRAG': row[0],
                'NAKLAD': row[1],
                'NAIMEN': row[4],
                'DDATE': dt,
                'ARTIKUL': row[3],
                'NDS': row[5],
                'KOLVO': row[6],
                'CENA': row[7]
            })
            added += 1
        existing_invoices.add(invoice)
    return added, skipped


def transform_purchase_data(input_path, log_path=None):
    if log_path is None:
        base, _ = os.path.splitext(input_path)
        log_path = base + "_log.txt"

    with open(log_path, 'w', encoding='utf-8') as f:
        f.write(f"=== Лог преобразования ===\nВходной файл: {input_path}\n\n")

    write_log(log_path, "Начало обработки")

    try:
        sheet_data = read_excel_file(input_path)
        write_log(log_path, f"Прочитан файл, строк: {len(sheet_data)}")

        if len(sheet_data) < 8:
            raise ValueError("Файл содержит меньше 8 строк.")

        header_row = sheet_data[7]  # 8-я строка (индекс 7)
        if len(header_row) < 10:
            raise ValueError("Строка заголовка слишком короткая.")

        col_a = header_row[0].lower()
        col_g = header_row[6].lower()
        if 'контрагент' not in col_a or 'артикул' not in col_g:
            write_log(log_path, f"Предупреждение: в строке 8 не найдены ожидаемые заголовки. "
                                f"A='{header_row[0]}', G='{header_row[6]}'.")

        data_rows = []
        for row_idx in range(8, len(sheet_data)):
            row = sheet_data[row_idx]
            if len(row) < 10:
                continue
            if row[0].strip() == 'Итого' or row[2].strip() == 'Итого':
                write_log(log_path, f"Строка {row_idx+1} — 'Итого', остановка.")
                break
            data_rows.append([
                row[0].strip(), row[3].strip(), row[5].strip(),
                row[6].strip(), row[7].strip(), row[8].strip(), row[9].strip()
            ])

        write_log(log_path, f"Собрано {len(data_rows)} строк.")

        if not data_rows:
            return 0, 0

        transformed = []
        for row in data_rows:
            raw_art_nom = row[3]
            if ',' in raw_art_nom:
                art, nom = raw_art_nom.split(',', 1)
                art = art.strip()
                nom = nom.strip()
            else:
                art = raw_art_nom
                nom = ''

            nds_val = row[4]
            if nds_val.lower() == 'без ндс':
                nds = 0.0
            else:
                nds_val = nds_val.replace('%', '').strip()
                try:
                    nds = float(nds_val) if nds_val else 0.0
                except ValueError:
                    nds = 0.0

            transformed.append([
                row[0], row[1], row[2],
                art, nom,
                round(nds, 2),
                safe_float(row[5]),
                safe_float(row[6])
            ])

        grouped = defaultdict(list)
        for r in transformed:
            grouped[r[1].strip()].append(r)

        write_log(log_path, f"Сгруппировано по накладным: {len(grouped)} шт.")

        dbf_path = r"\\Synas\work\ПВХ\УЧЕТ-1С\1C.dbf"
        if not DBF_AVAILABLE:
            raise RuntimeError("Библиотека 'dbf' не установлена.")

        dbf_dir = os.path.dirname(dbf_path)
        if not os.path.exists(dbf_dir):
            os.makedirs(dbf_dir)
            write_log(log_path, f"Создана папка: {dbf_dir}")

        table = create_or_open_dbf(dbf_path, log_path)
        existing = get_existing_invoices(table, log_path)
        added, skipped = append_new_records(table, grouped, existing, log_path)
        table.close()

        write_log(log_path, f"Добавлено строк: {added}, пропущено строк: {skipped}")
        return added, skipped

    except Exception as e:
        write_log(log_path, f"Ошибка: {e}\n{traceback.format_exc()}")
        raise


class App:
    def __init__(self, root):
        self.root = root
        root.title("Загрузка в 1С (DBF)")
        root.geometry("650x200")

        self.input_path = tk.StringVar()

        tk.Label(root, text="Входной файл (.xls/.xlsx):").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        tk.Entry(root, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(root, text="Обзор...", command=self.browse_input).grid(row=0, column=2, padx=5, pady=5)

        tk.Button(root, text="Загрузить в 1С", command=self.run, bg="lightblue", font=("Arial", 12)).grid(row=1, column=1, pady=20)

        self.info_label = tk.Label(root, text="Выберите файл выгрузки (.xls или .xlsx). Данные добавятся в \\Synas\work\ПВХ\УЧЕТ-1С\1C.dbf", fg="gray")
        self.info_label.grid(row=2, column=0, columnspan=3, pady=5)

        if not DBF_AVAILABLE:
            self.info_label.config(text="ВНИМАНИЕ: библиотека 'dbf' не установлена. Установите: pip install dbf", fg="orange")

    def browse_input(self):
        fname = filedialog.askopenfilename(
            title="Выберите файл выгрузки",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("Excel 97-2003", "*.xls"), ("Excel files", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if fname:
            self.input_path.set(fname)

    def run(self):
        inp = self.input_path.get().strip()
        if not inp:
            messagebox.showerror("Ошибка", "Входной файл не выбран")
            return

        base_in, _ = os.path.splitext(inp)
        log_file = base_in + "_log.txt"

        try:
            added, skipped = transform_purchase_data(inp, log_file)
            msg = f"Загрузка завершена!\nДобавлено новых строк: {added}\nПропущено строк (накладные уже есть): {skipped}"
            self.info_label.config(text=msg, fg="green")
            messagebox.showinfo("Успех", msg)
        except Exception as e:
            self.info_label.config(text=f"Ошибка: {e}\nПодробности в логе: {log_file}", fg="red")
            messagebox.showerror("Ошибка", f"Произошла ошибка. Подробности в логе:\n{log_file}")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
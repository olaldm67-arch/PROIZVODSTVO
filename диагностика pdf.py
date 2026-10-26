import re
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Опциональные зависимости
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

try:
    import zxingcpp
    HAS_ZXING = True
except ImportError:
    HAS_ZXING = False

try:
    import numpy as np
    from PIL import Image
    import io
    HAS_IMG = True
except ImportError:
    HAS_IMG = False


# ============================================================
# ДИАГНОСТИКА
# ============================================================

def diagnose_pdf(pdf_path, log):
    """Полная диагностика PDF — все слои и все методы извлечения."""

    log("=" * 70)
    log(f"ФАЙЛ: {pdf_path}")
    log(f"Размер: {os.path.getsize(pdf_path)} байт")
    log("=" * 70)

    log("\n[ЗАВИСИМОСТИ]")
    log(f"  pdfplumber: {'✅' if HAS_PDFPLUMBER else '❌ не установлен'}")
    log(f"  PyMuPDF (fitz): {'✅' if HAS_FITZ else '❌ не установлен'}")
    log(f"  zxing-cpp: {'✅' if HAS_ZXING else '❌ не установлен'}")
    log(f"  PIL+numpy: {'✅' if HAS_IMG else '❌ не установлен'}")

    if not HAS_PDFPLUMBER:
        log("\n[ОШИБКА] pdfplumber не установлен. Установите:")
        log("  pip install pdfplumber")
        return

    # ============================================================
    # ЧАСТЬ 1: pdfplumber — анализ текстовых слоёв
    # ============================================================
    log("\n" + "=" * 70)
    log("ЧАСТЬ 1: pdfplumber — анализ текста")
    log("=" * 70)

    with pdfplumber.open(pdf_path) as pdf:
        log(f"Страниц: {len(pdf.pages)}")

        for pn, page in enumerate(pdf.pages, 1):
            log(f"\n--- Страница {pn} ---")
            log(f"Размер страницы: {page.width:.0f} x {page.height:.0f} pt")

            # 1.1 Сырой текст
            text = page.extract_text() or ''
            log(f"\n[1.1] extract_text(): {len(text)} символов")
            nums_12 = re.findall(r'\d{12,20}', text)
            log(f"  Числа 12+ знаков: {nums_12 if nums_12 else '(нет)'}")

            # Первые 300 символов текста
            preview = text[:300].replace('\n', ' | ')
            log(f"  Превью: {preview}...")

            # 1.2 extract_words с координатами
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
            log(f"\n[1.2] extract_words(): {len(words)} слов")
            digit_words = [w for w in words if re.search(r'\d', w['text'])]
            log(f"  Слов с цифрами: {len(digit_words)}")

            # Показываем слова с 10+ знаками или рядом с "ДЕТАЛ"
            interesting = []
            for w in words:
                txt = w['text']
                if re.fullmatch(r'\d{10,20}', txt):
                    interesting.append(w)
                elif 'ДЕТАЛ' in txt.upper() or 'КОД' in txt.upper():
                    interesting.append(w)

            log(f"  Интересных слов (числа 10+ или 'ДЕТАЛ'/'КОД'): {len(interesting)}")
            for w in interesting[:30]:
                log(f"    x0={w['x0']:7.1f} top={w['top']:7.1f} "
                    f"x1={w['x1']:7.1f}  '{w['text']}'")

            # 1.3 Сырые символы (chars) — низкий уровень
            chars = page.chars
            log(f"\n[1.3] chars: {len(chars)} символов всего")
            digit_chars = [c for c in chars if c['text'].isdigit()]
            log(f"  Цифровых символов: {len(digit_chars)}")

            if digit_chars:
                # Сортируем по top, затем по x0
                digit_chars_sorted = sorted(digit_chars, key=lambda c: (round(c['top']), c['x0']))
                log(f"  Первые 40 цифровых символов:")
                for c in digit_chars_sorted[:40]:
                    color = c.get('non_stroking_color', '?')
                    log(f"    x0={c['x0']:7.1f} top={c['top']:7.1f} "
                        f"size={c.get('size', 0):.1f} "
                        f"color={color}  '{c['text']}'")

                # Собираем числа из chars
                numbers = collect_barcode_digits_from_chars(page)
                log(f"\n  Числа, собранные из chars: {numbers if numbers else '(нет)'}")

            # 1.4 Изображения
            log(f"\n[1.4] Изображения: {len(page.images)}")
            for idx, img in enumerate(page.images):
                bbox = img.get('bbox')
                w = img.get('width')
                h = img.get('height')
                log(f"    #{idx}: bbox={bbox}  исходный размер={w}x{h}px")

            # 1.5 Кривые (vector graphics)
            curves = page.curves
            log(f"\n[1.5] Векторных кривых: {len(curves)}")
            if curves:
                log(f"    (штрих-код может быть векторным — не только raster)")

            # 1.6 Rectangles
            rects = page.rects
            log(f"    Прямоугольников: {len(rects)}")

    # ============================================================
    # ЧАСТЬ 2: PyMuPDF — рендер и декодирование штрих-кодов
    # ============================================================
    log("\n" + "=" * 70)
    log("ЧАСТЬ 2: PyMuPDF — рендер страниц и декодирование")
    log("=" * 70)

    if not HAS_FITZ:
        log("[!] PyMuPDF не установлен. Установите: pip install PyMuPDF")
    else:
        with fitz.open(pdf_path) as doc:
            for pn, page in enumerate(doc, 1):
                log(f"\n--- Страница {pn} ---")
                log(f"Размер: {page.rect.width:.0f} x {page.rect.height:.0f} pt")

                # 2.1 Сырой текст через fitz
                text_fitz = page.get_text()
                nums_fitz = re.findall(r'\d{12,20}', text_fitz)
                log(f"[2.1] fitz.get_text(): {len(text_fitz)} символов")
                log(f"  Числа 12+ знаков: {nums_fitz if nums_fitz else '(нет)'}")

                # 2.2 Текст с координатами через fitz
                blocks = page.get_text("dict")["blocks"]
                log(f"[2.2] Блоков текста: {len(blocks)}")
                for b_idx, block in enumerate(blocks):
                    if block.get('type') == 0:  # text block
                        for line in block.get('lines', []):
                            for span in line.get('spans', []):
                                txt = span.get('text', '')
                                if re.search(r'\d{10,}', txt):
                                    bbox = span.get('bbox', (0,0,0,0))
                                    color = span.get('color', 0)
                                    size = span.get('size', 0)
                                    log(f"    span: bbox={tuple(round(v,1) for v in bbox)} "
                                        f"size={size:.1f} color=0x{color:06x} '{txt}'")

                # 2.3 Растровые изображения
                images_info = page.get_images(full=True)
                log(f"[2.3] Растровых изображений: {len(images_info)}")

                # 2.4 Рендер + декодирование zxing
                if HAS_ZXING and HAS_IMG:
                    log(f"[2.4] Попытка декодирования штрих-кодов...")
                    for dpi in (200, 300, 400, 600):
                        try:
                            mat = fitz.Matrix(dpi/72, dpi/72)
                            pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
                            img = Image.open(io.BytesIO(pix.tobytes("png")))
                            arr = np.array(img)
                            results = zxingcpp.read_barcodes(arr)
                            codes = [r.text for r in results if r.valid]
                            log(f"    dpi={dpi}: {len(codes)} штрих-кодов: {codes}")
                            if codes:
                                break  # нашли — хватит
                        except Exception as e:
                            log(f"    dpi={dpi}: ошибка {e}")
                else:
                    log("[2.4] Декодирование недоступно (zxing-cpp или PIL/numpy не установлены)")

    # ============================================================
    # ЧАСТЬ 3: Итоговое сравнение
    # ============================================================
    log("\n" + "=" * 70)
    log("ЧАСТЬ 3: Сводка")
    log("=" * 70)

    # Повторно берём из pdfplumber и собираем все числа
    with pdfplumber.open(pdf_path) as pdf:
        total_chars = 0
        total_digits = 0
        all_numbers = []
        for page in pdf.pages:
            total_chars += len(page.chars)
            for c in page.chars:
                if c['text'].isdigit():
                    total_digits += 1
            nums = collect_barcode_digits_from_chars(page)
            all_numbers.extend(nums)

        log(f"Всего символов в текстовом слое: {total_chars}")
        log(f"Из них цифр: {total_digits}")
        log(f"Собранных чисел-кандидатов (12-20 знаков): {len(all_numbers)}")
        for n in all_numbers:
            log(f"  • {n}  (длина {len(n)})")

    log("\n" + "=" * 70)
    log("ДИАГНОСТИКА ЗАВЕРШЕНА")
    log("=" * 70)
    log("\nВывод:")
    if total_digits >= 100:
        log("  ✅ Цифры присутствуют в текстовом слое — используем extract_chars")
    elif total_digits > 0:
        log("  ⚠ Цифры есть, но мало — возможно часть под картинкой")
    else:
        log("  ❌ Цифр в текстовом слое нет — нужен zxing-cpp")


def collect_barcode_digits_from_chars(page, min_len=12, max_len=20):
    """
    Собирает последовательности цифр из символов страницы.
    Работает, даже если цифры лежат под картинкой.
    """
    chars = page.chars
    digit_chars = [c for c in chars if c['text'].isdigit()]
    if not digit_chars:
        return []

    by_line = {}
    for c in digit_chars:
        key = round(c['top'] / 2) * 2
        by_line.setdefault(key, []).append(c)

    results = []
    for top, group in by_line.items():
        group.sort(key=lambda c: c['x0'])
        buf = ''
        prev_right = None
        for c in group:
            if prev_right is not None:
                gap = c['x0'] - prev_right
                if gap > c.get('size', 10) * 0.8:
                    if min_len <= len(buf) <= max_len:
                        results.append(buf)
                    buf = ''
            buf += c['text']
            prev_right = c['x1']
        if min_len <= len(buf) <= max_len:
            results.append(buf)

    return results


# ============================================================
# GUI
# ============================================================

class DiagApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Диагностика PDF конструктора")
        self.geometry("980x720")

        self.path_var = tk.StringVar()

        pad = {'padx': 8, 'pady': 4}

        tk.Label(self, text="PDF-файл:").grid(row=0, column=0, sticky='w', **pad)
        tk.Entry(self, textvariable=self.path_var, width=80).grid(
            row=0, column=1, sticky='we', **pad)
        tk.Button(self, text="Обзор...", command=self.choose_file).grid(
            row=0, column=2, **pad)
        tk.Button(self, text="Запустить диагностику", command=self.run_diag,
                  bg="#2E7D32", fg="white",
                  font=("Arial", 10, "bold")).grid(
            row=1, column=1, sticky='e', **pad)
        tk.Button(self, text="Сохранить лог в файл", command=self.save_log).grid(
            row=1, column=2, **pad)

        self.log_text = tk.Text(self, wrap='word', height=35, font=("Consolas", 9))
        self.log_text.grid(row=2, column=0, columnspan=3, sticky='nsew', **pad)

        scroll = ttk.Scrollbar(self, command=self.log_text.yview)
        scroll.grid(row=2, column=3, sticky='ns')
        self.log_text.config(yscrollcommand=scroll.set)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)

    def log(self, msg):
        self.log_text.insert('end', str(msg) + '\n')
        self.log_text.see('end')
        self.update_idletasks()

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Выберите PDF конструктора",
            filetypes=[("PDF", "*.pdf"), ("Все файлы", "*.*")])
        if path:
            self.path_var.set(path)

    def run_diag(self):
        path = self.path_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Ошибка", "PDF-файл не найден.")
            return

        self.log_text.delete('1.0', 'end')
        try:
            diagnose_pdf(path, log=self.log)
        except Exception as e:
            self.log(f"\n[КРИТИЧЕСКАЯ ОШИБКА] {e}")
            import traceback
            self.log(traceback.format_exc())

    def save_log(self):
        content = self.log_text.get('1.0', 'end')
        if not content.strip():
            messagebox.showwarning("Пусто", "Лог пуст — сначала запустите диагностику.")
            return
        path = filedialog.asksaveasfilename(
            title="Сохранить лог", defaultextension=".txt",
            filetypes=[("Текст", "*.txt"), ("Все файлы", "*.*")],
            initialfile="pdf_diagnostic.txt")
        if not path:
            return
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        messagebox.showinfo("Готово", f"Лог сохранён:\n{path}")


if __name__ == '__main__':
    App = DiagApp()
    App.mainloop()
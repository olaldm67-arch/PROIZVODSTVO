import os
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk, ImageFile

# Разрешаем загрузку обрезанных JPEG
ImageFile.LOAD_TRUNCATED_IMAGES = True

class ImageViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Просмотр изображений (поддержка обрезанных JPEG)")
        self.root.geometry("1000x700")

        self.current_folder = ""
        self.image_files = []
        self.current_index = 0

        # Хранение оригинального изображения и текущего состояния обрезки
        self.original_image = None      # PIL Image объект
        self.crop_offset = 0            # количество пикселей обрезки со всех сторон

        # Левая панель: список файлов
        left_frame = tk.Frame(root, width=200)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        tk.Button(left_frame, text="Выбрать папку", command=self.select_folder).pack(pady=5)
        self.listbox = tk.Listbox(left_frame, font=("Arial", 10), selectmode=tk.SINGLE)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)

        # Правая панель: изображение
        right_frame = tk.Frame(root)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.image_label = tk.Label(right_frame, text="Выберите папку и файл", font=("Arial", 20))
        self.image_label.pack(expand=True)

        # Нижняя панель: навигация, информация и кнопки обрезки
        nav_frame = tk.Frame(root)
        nav_frame.pack(side=tk.BOTTOM, fill=tk.X)

        tk.Button(nav_frame, text="◀ Предыдущее", command=self.prev_image).pack(side=tk.LEFT, padx=10, pady=5)
        tk.Button(nav_frame, text="Следующее ▶", command=self.next_image).pack(side=tk.LEFT, padx=10, pady=5)

        # Кнопки обрезки
        tk.Button(nav_frame, text="Обрезать (+10)", command=self.crop_image).pack(side=tk.LEFT, padx=10, pady=5)
        tk.Button(nav_frame, text="Сбросить обрезку", command=self.reset_crop).pack(side=tk.LEFT, padx=10, pady=5)

        self.info_label = tk.Label(nav_frame, text="", font=("Arial", 10))
        self.info_label.pack(side=tk.RIGHT, padx=10)

        # Поле для логов (внизу)
        self.log_text = tk.Text(root, height=6, font=("Courier", 8))
        self.log_text.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return
        self.current_folder = folder
        self.load_images()

    def load_images(self):
        if not os.path.exists(self.current_folder):
            self.log(f"Папка не найдена: {self.current_folder}")
            return

        extensions = ('.jpg', '.jpeg', '.jpe', '.jfif', '.jif')
        self.image_files = [f for f in os.listdir(self.current_folder)
                            if f.lower().endswith(extensions)]
        self.image_files.sort()

        self.listbox.delete(0, tk.END)
        for f in self.image_files:
            self.listbox.insert(tk.END, f)

        if self.image_files:
            self.current_index = 0
            self.show_image()
            self.log(f"Загружено {len(self.image_files)} файлов")
        else:
            self.image_label.config(text="Нет изображений в папке")
            self.log("Нет изображений в папке")

    def show_image(self):
        if not self.image_files:
            return
        if self.current_index < 0 or self.current_index >= len(self.image_files):
            self.current_index = 0

        filename = self.image_files[self.current_index]
        path = os.path.join(self.current_folder, filename)

        try:
            self.original_image = Image.open(path)
            self.original_image.load()  # принудительная загрузка
            self.crop_offset = 0
            self.display_image()
        except Exception as e:
            self.image_label.config(text=f"Ошибка загрузки {filename}", image="")
            self.log(f"Ошибка открытия {filename}: {e}")

    def display_image(self):
        """Отображает изображение с применённой обрезкой."""
        if self.original_image is None:
            return

        # Применяем обрезку, если offset > 0 и размеры позволяют
        img = self.original_image.copy()
        if self.crop_offset > 0:
            w, h = img.size
            if w > 2 * self.crop_offset and h > 2 * self.crop_offset:
                img = img.crop((self.crop_offset, self.crop_offset,
                                w - self.crop_offset, h - self.crop_offset))
            else:
                self.log("Обрезка невозможна: изображение слишком мало")
                self.crop_offset = 0

        # Масштабирование с отступами 50 пикселей
        padding = 50
        try:
            win_width = self.root.winfo_width()
            win_height = self.root.winfo_height()
        except:
            win_width = 1000
            win_height = 700
        left_width = 200
        display_width = win_width - left_width - 2 * padding
        nav_height = 80  # приблизительная высота нижней панели
        display_height = win_height - nav_height - 2 * padding
        if display_width < 100:
            display_width = 800
            display_height = 600

        img.thumbnail((display_width, display_height), Image.Resampling.LANCZOS)

        tk_img = ImageTk.PhotoImage(img)
        self.image_label.config(image=tk_img, text="")
        self.image_label.image = tk_img

        # Обновляем информацию
        filename = self.image_files[self.current_index]
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(self.current_index)
        self.listbox.see(self.current_index)

        self.info_label.config(text=f"{self.current_index+1}/{len(self.image_files)}  {filename}  |  Обрезка: {self.crop_offset}px")
        self.log(f"Отображено: {filename} (размер {img.width}x{img.height}, обрезка {self.crop_offset}px)")

    def crop_image(self):
        """Увеличивает обрезку на 10 пикселей."""
        if self.original_image is None:
            return
        self.crop_offset += 10
        self.display_image()

    def reset_crop(self):
        """Сбрасывает обрезку."""
        if self.original_image is None:
            return
        self.crop_offset = 0
        self.display_image()

    def on_select(self, event):
        selection = self.listbox.curselection()
        if selection:
            self.current_index = selection[0]
            self.show_image()

    def prev_image(self):
        if self.image_files:
            self.current_index = (self.current_index - 1) % len(self.image_files)
            self.show_image()

    def next_image(self):
        if self.image_files:
            self.current_index = (self.current_index + 1) % len(self.image_files)
            self.show_image()


if __name__ == "__main__":
    root = tk.Tk()
    app = ImageViewer(root)
    root.mainloop()
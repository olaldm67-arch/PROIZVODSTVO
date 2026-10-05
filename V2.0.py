# -*- coding: utf-8 -*-
import sys
import socket
import pymysql
import tkinter as tk
from tkinter import ttk, messagebox

# ----------------------------------------------------------------------
# Параметры подключения к базе данных (исправлен пароль)
# ----------------------------------------------------------------------
DB_HOST = '192.168.0.10'
DB_USER = 'workshop'
DB_PASSWORD = 'w0rK5h0p'          # правильный пароль (латинская p)
DB_NAME = 'workshop'              # уточните имя базы
DB_CHARSET = 'utf8'

# ----------------------------------------------------------------------
# Вспомогательные функции работы с БД
# ----------------------------------------------------------------------

def get_local_ip_octet():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception:
        ip = '127.0.0.1'
    return int(ip.split('.')[-1])

def worker_exists(cursor, rab):
    cursor.execute("SELECT FIO FROM WORKER WHERE RAB = %s", (rab,))
    return cursor.fetchone() is not None

def get_open_session(cursor, rab):
    cursor.execute(
        "SELECT IP, TIME_START_W FROM WORKER_TIME "
        "WHERE RAB = %s AND DATE(TIME_START_W) = CURDATE() AND TIME_END_W IS NULL",
        (rab,)
    )
    return cursor.fetchone()

def close_session(cursor, rab):
    cursor.execute(
        "UPDATE WORKER_TIME SET TIME_END_W = NOW() "
        "WHERE RAB = %s AND TIME_END_W IS NULL",
        (rab,)
    )

def decrement_workplace_cur(cursor, ip):
    cursor.execute(
        "UPDATE WORKPLACE SET RAB_CUR = RAB_CUR - 1 WHERE IP = %s AND RAB_CUR > 0",
        (ip,)
    )
    return cursor.rowcount > 0

def increment_workplace_cur(cursor, ip):
    cursor.execute(
        "UPDATE WORKPLACE SET RAB_CUR = RAB_CUR + 1 WHERE IP = %s",
        (ip,)
    )

def create_session(cursor, rab, ip):
    cursor.execute(
        "INSERT INTO WORKER_TIME (RAB, IP, TIME_START_W, TIME_END_W) "
        "VALUES (%s, %s, NOW(), NULL)",
        (rab, ip)
    )

def get_workplace_info(cursor, ip):
    cursor.execute(
        "SELECT NAME, RAB_MAX, RAB_CUR FROM WORKPLACE WHERE IP = %s",
        (ip,)
    )
    return cursor.fetchone()

def get_current_workers(cursor, ip):
    cursor.execute(
        "SELECT w.RAB, w.FIO FROM WORKER_TIME wt "
        "JOIN WORKER w ON wt.RAB = w.RAB "
        "WHERE wt.IP = %s AND DATE(wt.TIME_START_W) = CURDATE() AND wt.TIME_END_W IS NULL",
        (ip,)
    )
    return cursor.fetchall()

# ----------------------------------------------------------------------
# Основной класс графического интерфейса
# ----------------------------------------------------------------------

class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт рабочего времени")
        self.root.geometry("600x500")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Определяем IP и подключаемся к БД
        self.current_ip = get_local_ip_octet()
        self.conn = None
        self.cursor = None

        try:
            self.conn = pymysql.connect(
                host=DB_HOST,
                user=DB_USER,
                password=DB_PASSWORD,
                database=DB_NAME,
                charset=DB_CHARSET,
                autocommit=False
            )
            self.cursor = self.conn.cursor()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться к БД:\n{e}")
            sys.exit(1)

        # Проверяем существование рабочего места
        workplace = get_workplace_info(self.cursor, self.current_ip)
        if not workplace:
            messagebox.showerror("Ошибка", f"Рабочее место с IP {self.current_ip} не найдено в БД.")
            sys.exit(1)

        self.workplace_name, self.rab_max, self.rab_cur = workplace

        # ------------------- Интерфейс -------------------

        # Верхняя панель: название места
        self.label_place = tk.Label(
            root,
            text=self.workplace_name,
            font=("Arial", 36, "bold"),
            fg="blue"
        )
        self.label_place.pack(pady=10)

        # Информация о вместимости
        self.info_label = tk.Label(
            root,
            text=f"Максимум: {self.rab_max} | Текущее занято: {self.rab_cur}",
            font=("Arial", 12)
        )
        self.info_label.pack()

        # Список работников
        self.list_frame = tk.Frame(root)
        self.list_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(self.list_frame, font=("Arial", 14), height=10)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(self.list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        # Поле ввода штрих-кода (сканер)
        self.entry_label = tk.Label(root, text="Сканируйте штрих-код:", font=("Arial", 12))
        self.entry_label.pack(pady=(10, 0))

        self.entry = tk.Entry(root, font=("Arial", 14), width=20)
        self.entry.pack(pady=5)
        self.entry.bind("<Return>", self.on_scan)
        self.entry.focus_set()

        # Кнопка выхода
        self.exit_button = tk.Button(root, text="Выход", command=self.on_closing, font=("Arial", 12))
        self.exit_button.pack(pady=10)

        # Первоначальное обновление списка
        self.update_workers_list()

    # ------------------------------------------------------------------
    # Функция для автозакрывающегося информационного окна
    # ------------------------------------------------------------------
    def show_info_auto(self, message):
        """Показывает информационное окно, которое закрывается через 5 секунд."""
        win = tk.Toplevel(self.root)
        win.title("Информация")
        win.geometry("300x120")
        win.resizable(False, False)
        # Делаем окно поверх всех окон
        win.transient(self.root)
        win.grab_set()  # блокирует взаимодействие с главным окном

        tk.Label(win, text=message, font=("Arial", 12), wraplength=280).pack(pady=20)
        # Автоматическое закрытие через 5000 мс
        win.after(5000, win.destroy)

    # ------------------------------------------------------------------
    # Обновление списка и информации
    # ------------------------------------------------------------------
    def update_workers_list(self):
        """Обновляет список работников и информацию о занятости."""
        self.listbox.delete(0, tk.END)
        try:
            workers = get_current_workers(self.cursor, self.current_ip)
            for rab, fio in workers:
                self.listbox.insert(tk.END, f"{rab} – {fio}")
            # Обновляем информацию о количестве занятых мест
            info = get_workplace_info(self.cursor, self.current_ip)
            if info:
                _, _, cur = info
                self.rab_cur = cur
                self.info_label.config(text=f"Максимум: {self.rab_max} | Текущее занято: {self.rab_cur}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить список:\n{e}")

    # ------------------------------------------------------------------
    # Обработчик сканирования
    # ------------------------------------------------------------------
    def on_scan(self, event):
        rab = self.entry.get().strip()
        self.entry.delete(0, tk.END)

        if not rab:
            return

        try:
            # 1. Проверка существования работника
            if not worker_exists(self.cursor, rab):
                messagebox.showerror("Ошибка", f"Работник {rab} не найден в базе.")
                self.conn.rollback()
                return

            # 2. Поиск открытой сессии
            open_session = get_open_session(self.cursor, rab)

            if open_session:
                old_ip, start_time = open_session

                if old_ip == self.current_ip:
                    # Завершение работы на этом же месте
                    close_session(self.cursor, rab)
                    if not decrement_workplace_cur(self.cursor, self.current_ip):
                        self.conn.rollback()
                        messagebox.showerror("Ошибка", f"Счётчик RAB_CUR на месте {self.workplace_name} уже равен 0. Операция отменена.")
                        return
                    self.conn.commit()
                    self.show_info_auto(f"Работник {rab} завершил работу на {self.workplace_name}.")
                else:
                    # Перемещение на текущее место
                    info = get_workplace_info(self.cursor, self.current_ip)
                    if info is None:
                        self.conn.rollback()
                        messagebox.showerror("Ошибка", "Текущее рабочее место не найдено.")
                        return
                    _, max_cur, cur_cur = info
                    if cur_cur >= max_cur:
                        messagebox.showwarning("Предупреждение", f"Место {self.workplace_name} заполнено (максимум {max_cur}). Перемещение невозможно.")
                        self.conn.rollback()
                        return

                    # Закрываем старую сессию, уменьшаем счётчик старого места
                    close_session(self.cursor, rab)
                    if not decrement_workplace_cur(self.cursor, old_ip):
                        self.conn.rollback()
                        messagebox.showerror("Ошибка", f"Счётчик RAB_CUR на старом месте уже равен 0. Операция отменена.")
                        return
                    # Создаём новую сессию, увеличиваем счётчик нового места
                    create_session(self.cursor, rab, self.current_ip)
                    increment_workplace_cur(self.cursor, self.current_ip)
                    self.conn.commit()
                    self.show_info_auto(f"Работник {rab} перемещён на {self.workplace_name}.")
            else:
                # Новая регистрация
                info = get_workplace_info(self.cursor, self.current_ip)
                if info is None:
                    self.conn.rollback()
                    messagebox.showerror("Ошибка", "Текущее рабочее место не найдено.")
                    return
                _, max_cur, cur_cur = info
                if cur_cur >= max_cur:
                    messagebox.showwarning("Предупреждение", f"Место {self.workplace_name} заполнено (максимум {max_cur}). Регистрация невозможна.")
                    self.conn.rollback()
                    return

                create_session(self.cursor, rab, self.current_ip)
                increment_workplace_cur(self.cursor, self.current_ip)
                self.conn.commit()
                self.show_info_auto(f"Работник {rab} зарегистрирован на {self.workplace_name}.")

            # Обновляем список после успешной операции
            self.update_workers_list()

        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")

    # ------------------------------------------------------------------
    # Закрытие приложения
    # ------------------------------------------------------------------
    def on_closing(self):
        if self.conn:
            self.conn.close()
        self.root.destroy()


# ----------------------------------------------------------------------
# Запуск приложения
# ----------------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(root)
    root.mainloop()
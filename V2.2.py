# -*- coding: utf-8 -*-
import sys
import socket
import pymysql
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

# ----------------------------------------------------------------------
# Параметры подключения к базе данных
# ----------------------------------------------------------------------
DB_HOST = '192.168.0.10'
DB_USER = 'workshop'
DB_PASSWORD = 'w0rK5h0p'
DB_NAME = 'workshop'
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

def get_worker_fio(cursor, rab):
    cursor.execute("SELECT FIO FROM WORKER WHERE RAB = %s", (rab,))
    row = cursor.fetchone()
    return row[0] if row else None

def get_workplace_name(cursor, ip):
    cursor.execute("SELECT NAME FROM WORKPLACE WHERE IP = %s", (ip,))
    row = cursor.fetchone()
    return row[0] if row else str(ip)

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
        "SELECT wt.RAB, w.FIO, wt.TIME_START_W "
        "FROM WORKER_TIME wt "
        "JOIN WORKER w ON wt.RAB = w.RAB "
        "WHERE wt.IP = %s AND DATE(wt.TIME_START_W) = CURDATE() AND wt.TIME_END_W IS NULL",
        (ip,)
    )
    rows = cursor.fetchall()
    result = []
    for rab, fio, start_time in rows:
        cursor.execute(
            "SELECT SUM(TIMESTAMPDIFF(SECOND, TIME_START_W, IFNULL(TIME_END_W, NOW()))) "
            "FROM WORKER_TIME "
            "WHERE RAB = %s AND IP = %s AND DATE(TIME_START_W) = CURDATE()",
            (rab, ip)
        )
        total_sec = cursor.fetchone()[0]
        if total_sec is None:
            total_sec = 0
        result.append((rab, fio, start_time, total_sec))
    return result

# ----------------------------------------------------------------------
# Основной класс графического интерфейса
# ----------------------------------------------------------------------

class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт рабочего времени")
        self.root.geometry("900x650")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

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

        workplace = get_workplace_info(self.cursor, self.current_ip)
        if not workplace:
            messagebox.showerror("Ошибка", f"Рабочее место с IP {self.current_ip} не найдено в БД.")
            sys.exit(1)

        self.workplace_name, self.rab_max, self.rab_cur = workplace

        # ------------------- Интерфейс -------------------

        self.label_place = tk.Label(
            root,
            text=self.workplace_name,
            font=("Arial", 40, "bold"),
            fg="blue"
        )
        self.label_place.pack(pady=10)

        self.info_label = tk.Label(
            root,
            text=f"Максимум: {self.rab_max} | Текущее занято: {self.rab_cur}",
            font=("Arial", 14)
        )
        self.info_label.pack()

        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 16), rowheight=30)
        style.configure("Treeview.Heading", font=("Arial", 16, "bold"))

        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=("num", "fio", "start", "total"),
            show="headings",
            height=12
        )
        self.tree.heading("num", text="№")
        self.tree.heading("fio", text="ФИО")
        self.tree.heading("start", text="Время начала")
        self.tree.heading("total", text="Отработано")
        self.tree.column("num", width=70, anchor="center")
        self.tree.column("fio", width=400)
        self.tree.column("start", width=160, anchor="center")
        self.tree.column("total", width=160, anchor="center")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.config(yscrollcommand=scrollbar.set)

        self.entry_label = tk.Label(root, text="Сканируйте штрих-код:", font=("Arial", 14))
        self.entry_label.pack(pady=(10, 0))

        self.entry = tk.Entry(root, font=("Arial", 16), width=20)
        self.entry.pack(pady=5)
        self.entry.bind("<Return>", self.on_scan)
        self.entry.focus_set()

        self.exit_button = tk.Button(root, text="Выход", command=self.on_closing, font=("Arial", 14))
        self.exit_button.pack(pady=15)

        self.close_previous_sessions()
        self.update_workers_list()
        self.auto_update()

    # ------------------------------------------------------------------
    def close_previous_sessions(self):
        try:
            self.cursor.execute(
                "SELECT RAB, IP FROM WORKER_TIME "
                "WHERE TIME_END_W IS NULL AND DATE(TIME_START_W) < CURDATE()"
            )
            sessions = self.cursor.fetchall()
            if not sessions:
                return

            count = 0
            for rab, ip in sessions:
                self.cursor.execute(
                    "UPDATE WORKER_TIME SET TIME_END_W = NOW() "
                    "WHERE RAB = %s AND IP = %s AND TIME_END_W IS NULL",
                    (rab, ip)
                )
                decrement_workplace_cur(self.cursor, ip)
                count += 1

            self.conn.commit()
            if count > 0:
                self.show_info_auto(f"Закрыто {count} сессий за предыдущие дни.")
        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Ошибка", f"Не удалось закрыть предыдущие сессии:\n{e}")

    def auto_update(self):
        self.update_workers_list()
        self.root.after(60000, self.auto_update)

    def show_info_auto(self, message):
        win = tk.Toplevel(self.root)
        win.title("Информация")
        win.geometry("400x150")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        tk.Label(win, text=message, font=("Arial", 14), wraplength=370).pack(pady=20)
        win.after(3000, win.destroy)

    def update_workers_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            workers = get_current_workers(self.cursor, self.current_ip)
            idx = 1
            for rab, fio, start_time, total_sec in workers:
                start_str = start_time.strftime("%H:%M") if start_time else "--:--"
                hours = total_sec // 3600
                minutes = (total_sec % 3600) // 60
                if hours > 0:
                    time_str = f"{hours}ч {minutes}м"
                else:
                    time_str = f"{minutes}м"
                self.tree.insert("", "end", values=(idx, fio, start_str, time_str))
                idx += 1

            info = get_workplace_info(self.cursor, self.current_ip)
            if info:
                _, _, cur = info
                self.rab_cur = cur
                self.info_label.config(text=f"Максимум: {self.rab_max} | Текущее занято: {self.rab_cur}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить таблицу:\n{e}")

    # ------------------------------------------------------------------
    # Обработчик сканирования (с улучшенной диагностикой заполнения)
    # ------------------------------------------------------------------
    def on_scan(self, event):
        rab = self.entry.get().strip()
        self.entry.delete(0, tk.END)

        if not rab:
            return

        try:
            fio = get_worker_fio(self.cursor, rab)
            if fio is None:
                messagebox.showerror("Ошибка", f"Работник с кодом {rab} не найден в базе.")
                self.conn.rollback()
                return

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
                    self.show_info_auto(f"Работник {fio} завершил работу на {self.workplace_name}.")
                else:
                    # Перемещение на текущее место
                    # Получаем свежие данные о месте
                    info = get_workplace_info(self.cursor, self.current_ip)
                    if info is None:
                        self.conn.rollback()
                        messagebox.showerror("Ошибка", "Текущее рабочее место не найдено.")
                        return
                    name, max_cur, cur_cur = info

                    if cur_cur >= max_cur:
                        messagebox.showwarning(
                            "Предупреждение",
                            f"Место '{name}' заполнено (занято {cur_cur} из {max_cur}). Перемещение невозможно."
                        )
                        self.conn.rollback()
                        return

                    old_place_name = get_workplace_name(self.cursor, old_ip)

                    close_session(self.cursor, rab)
                    if not decrement_workplace_cur(self.cursor, old_ip):
                        self.conn.rollback()
                        messagebox.showerror("Ошибка", f"Счётчик RAB_CUR на старом месте уже равен 0. Операция отменена.")
                        return
                    create_session(self.cursor, rab, self.current_ip)
                    increment_workplace_cur(self.cursor, self.current_ip)
                    self.conn.commit()
                    self.show_info_auto(f"Работник {fio} перемещён с '{old_place_name}' на '{name}'.")
            else:
                # Новая регистрация
                info = get_workplace_info(self.cursor, self.current_ip)
                if info is None:
                    self.conn.rollback()
                    messagebox.showerror("Ошибка", "Текущее рабочее место не найдено.")
                    return
                name, max_cur, cur_cur = info

                if cur_cur >= max_cur:
                    messagebox.showwarning(
                        "Предупреждение",
                        f"Место '{name}' заполнено (занято {cur_cur} из {max_cur}). Регистрация невозможна."
                    )
                    self.conn.rollback()
                    return

                create_session(self.cursor, rab, self.current_ip)
                increment_workplace_cur(self.cursor, self.current_ip)
                self.conn.commit()
                self.show_info_auto(f"Работник {fio} зарегистрирован на {name}.")

            self.update_workers_list()

        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")

    def on_closing(self):
        if self.conn:
            self.conn.close()
        self.root.destroy()


# ----------------------------------------------------------------------
# Запуск
# ----------------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(root)
    root.mainloop()
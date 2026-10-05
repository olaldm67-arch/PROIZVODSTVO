# -*- coding: utf-8 -*-
import sys
import socket
import pymysql
import tkinter as tk
from tkinter import ttk
from datetime import datetime, time

# ----------------------------------------------------------------------
# Параметры подключения к базе данных
# ----------------------------------------------------------------------
DB_HOST = '192.168.0.10'
DB_USER = 'workshop'
DB_PASSWORD = 'w0rK5h0p'
DB_NAME = 'workshop'
DB_CHARSET = 'utf8'

# ----------------------------------------------------------------------
# Логирование
# ----------------------------------------------------------------------
LOG_FILE = 'app.log'

def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception as e:
        print(f"Ошибка записи лога: {e}")

# ----------------------------------------------------------------------
# Вспомогательные функции для сотрудников
# ----------------------------------------------------------------------
def get_local_ip_octet():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception as e:
        log_message(f"Ошибка получения IP: {e}")
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
        "SELECT NAME, RAB_MAX, RAB_CUR, GR FROM WORKPLACE WHERE IP = %s",
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
            "WHERE RAB = %s AND DATE(TIME_START_W) = CURDATE()",
            (rab,)
        )
        total_sec = cursor.fetchone()[0]
        if total_sec is None:
            total_sec = 0
        result.append((rab, fio, start_time, total_sec))
    return result

# ----------------------------------------------------------------------
# Функции для работы с деталями
# ----------------------------------------------------------------------
def get_detal_full_info(cursor, detal):
    cursor.execute(
        "SELECT ZAKNUM, KONNAME, PRIMPROIZV, MATNAME, MATCOLOR, LENGTH "
        "FROM CCALC WHERE DETAL = %s",
        (detal,)
    )
    row = cursor.fetchone()
    if row:
        return {
            'zaknum': row[0],
            'konname': row[1],
            'primproizv': row[2],
            'matname': row[3],
            'matcolor': row[4],
            'length': row[5]
        }
    return None

def get_open_detal_session_any(cursor, detal):
    cursor.execute(
        "SELECT IP, TIME_START_D FROM DETAL_TIME "
        "WHERE DETAL = %s AND TIME_END_D IS NULL",
        (detal,)
    )
    return cursor.fetchone()

def close_detal_session(cursor, detal, ip):
    cursor.execute(
        "UPDATE DETAL_TIME SET TIME_END_D = NOW() "
        "WHERE DETAL = %s AND IP = %s AND TIME_END_D IS NULL",
        (detal, ip)
    )

def create_detal_session(cursor, detal, ip, gr):
    cursor.execute(
        "INSERT INTO DETAL_TIME (IP, DETAL, TIME_START_D, TIME_END_D, GR) "
        "VALUES (%s, %s, NOW(), NULL, %s)",
        (ip, detal, gr)
    )

# ----------------------------------------------------------------------
# Основной класс приложения
# ----------------------------------------------------------------------
class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт рабочего времени и деталей")
        self.root.geometry("1100x750")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        log_message("=== Запуск программы ===")
        self.current_ip = get_local_ip_octet()
        log_message(f"Определён IP: {self.current_ip}")

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
            log_message("Подключение к БД успешно")
        except Exception as e:
            msg = f"Не удалось подключиться к БД: {e}"
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000)
            sys.exit(1)

        workplace = get_workplace_info(self.cursor, self.current_ip)
        if not workplace:
            msg = f"Рабочее место с IP {self.current_ip} не найдено в БД."
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000)
            sys.exit(1)

        self.workplace_name, self.rab_max, self.rab_cur, self.gr = workplace
        log_message(f"Рабочее место: {self.workplace_name}, GR={self.gr}, RAB_MAX={self.rab_max}, RAB_CUR={self.rab_cur}")

        # ------------------- Интерфейс -------------------
        # Верхняя панель: название места и информация о занятости
        top_panel = tk.Frame(root)
        top_panel.pack(pady=5, fill=tk.X)

        self.label_place = tk.Label(
            top_panel,
            text=self.workplace_name,
            font=("Arial", 36, "bold"),
            fg="blue"
        )
        self.label_place.pack(side=tk.LEFT, padx=20)

        self.info_label = tk.Label(
            top_panel,
            text=f"Максимум сотрудников: {self.rab_max} | Текущее занято: {self.rab_cur}",
            font=("Arial", 14)
        )
        self.info_label.pack(side=tk.RIGHT, padx=20)

        # Блок информации о детали (крупный текст без заголовков)
        self.detal_frame = ttk.LabelFrame(root, text="Информация о детали", padding=10)
        self.detal_frame.pack(fill=tk.X, padx=10, pady=5)

        # Три строки с крупным шрифтом
        self.info_line1 = tk.Label(self.detal_frame, text="", font=("Arial", 20, "bold"), fg="darkblue")
        self.info_line2 = tk.Label(self.detal_frame, text="", font=("Arial", 20, "bold"), fg="darkblue")
        self.info_line3 = tk.Label(self.detal_frame, text="", font=("Arial", 20, "bold"), fg="darkblue")
        self.info_line1.pack(pady=2)
        self.info_line2.pack(pady=2)
        self.info_line3.pack(pady=2)

        # Создаём разделитель
        ttk.Separator(root, orient='horizontal').pack(fill=tk.X, padx=10, pady=5)

        # Таблица сотрудников (нижняя часть)
        self.workers_frame = ttk.LabelFrame(root, text="Сотрудники на рабочем месте", padding=5)
        self.workers_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 16), rowheight=30)
        style.configure("Treeview.Heading", font=("Arial", 16, "bold"))

        tree_frame = tk.Frame(self.workers_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree_workers = ttk.Treeview(
            tree_frame,
            columns=("num", "fio", "start", "total"),
            show="headings",
            height=10
        )
        self.tree_workers.heading("num", text="№")
        self.tree_workers.heading("fio", text="ФИО")
        self.tree_workers.heading("start", text="Начало (тек. место)")
        self.tree_workers.heading("total", text="Отработано (всего)")
        self.tree_workers.column("num", width=70, anchor="center")
        self.tree_workers.column("fio", width=400)
        self.tree_workers.column("start", width=200, anchor="center")
        self.tree_workers.column("total", width=200, anchor="center")
        self.tree_workers.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_workers.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_workers.config(yscrollcommand=scrollbar.set)

        # Поле ввода штрих-кода
        entry_frame = tk.Frame(root)
        entry_frame.pack(pady=10, fill=tk.X)

        self.entry_label = tk.Label(entry_frame, text="Сканируйте штрих-код:", font=("Arial", 14))
        self.entry_label.pack(side=tk.LEFT, padx=10)

        self.entry = tk.Entry(entry_frame, font=("Arial", 16), width=25)
        self.entry.pack(side=tk.LEFT, padx=10)
        self.entry.bind("<Return>", self.on_scan)
        self.entry.focus_set()

        self.exit_button = tk.Button(root, text="Выход", command=self.on_closing, font=("Arial", 14))
        self.exit_button.pack(pady=10)

        # Закрываем предыдущие сессии сотрудников и обновляем список
        self.close_previous_sessions()
        self.update_workers_list()
        self.auto_update()

        # Очищаем информационный блок детали
        self.clear_detal_info()

    # ------------------------------------------------------------------
    # Всплывающее окно без кнопок (автозакрывается через timeout мс)
    # ------------------------------------------------------------------
    def show_auto_message(self, message, title="Информация", timeout=5000):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("500x150")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        tk.Label(win, text=message, font=("Arial", 14), wraplength=470).pack(pady=20)
        win.after(timeout, win.destroy)

    # ------------------------------------------------------------------
    # Очистка информационного блока детали
    # ------------------------------------------------------------------
    def clear_detal_info(self):
        self.info_line1.config(text="")
        self.info_line2.config(text="")
        self.info_line3.config(text="")

    def set_detal_info(self, info):
        line1 = f"{info.get('zaknum', '')}  {info.get('konname', '')}  {info.get('matcolor', '')}"
        length = info.get('length', 0)
        length_minus_6 = length - 6 if length is not None else 0
        line2 = f"{info.get('matname', '')}  {length_minus_6}"
        line3 = info.get('primproizv', '')
        self.info_line1.config(text=line1)
        self.info_line2.config(text=line2)
        self.info_line3.config(text=line3)

    # ------------------------------------------------------------------
    # Закрытие предыдущих сессий сотрудников
    # ------------------------------------------------------------------
    def close_previous_sessions(self):
        try:
            self.cursor.execute(
                "SELECT RAB, IP, TIME_START_W FROM WORKER_TIME "
                "WHERE TIME_END_W IS NULL AND DATE(TIME_START_W) < CURDATE()"
            )
            sessions = self.cursor.fetchall()
            if not sessions:
                log_message("Нет зависших сессий сотрудников за предыдущие дни.")
                return

            count = 0
            for rab, ip, start_time in sessions:
                close_time = start_time.replace(hour=18, minute=0, second=0, microsecond=0)
                if start_time > close_time:
                    close_time = start_time
                self.cursor.execute(
                    "UPDATE WORKER_TIME SET TIME_END_W = %s "
                    "WHERE RAB = %s AND IP = %s AND TIME_END_W IS NULL",
                    (close_time, rab, ip)
                )
                decrement_workplace_cur(self.cursor, ip)
                count += 1
                log_message(f"Закрыта сессия сотрудника {rab} на IP {ip} (время окончания {close_time})")

            self.conn.commit()
            if count > 0:
                msg = f"Закрыто {count} сессий сотрудников за предыдущие дни (в 18:00 дня открытия)."
                log_message(msg)
                self.show_auto_message(msg, "Информация", 5000)
        except Exception as e:
            self.conn.rollback()
            msg = f"Не удалось закрыть предыдущие сессии сотрудников: {e}"
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000)

    # ------------------------------------------------------------------
    # Автообновление каждую минуту
    # ------------------------------------------------------------------
    def auto_update(self):
        self.update_workers_list()
        self.root.after(60000, self.auto_update)

    # ------------------------------------------------------------------
    # Обновление списка сотрудников
    # ------------------------------------------------------------------
    def update_workers_list(self):
        for item in self.tree_workers.get_children():
            self.tree_workers.delete(item)

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
                self.tree_workers.insert("", "end", values=(idx, fio, start_str, time_str))
                idx += 1

            info = get_workplace_info(self.cursor, self.current_ip)
            if info:
                _, _, cur, _ = info
                self.rab_cur = cur
                self.info_label.config(text=f"Максимум сотрудников: {self.rab_max} | Текущее занято: {self.rab_cur}")
        except Exception as e:
            msg = f"Не удалось обновить список сотрудников: {e}"
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000)

    # ------------------------------------------------------------------
    # Обработчик сканирования
    # ------------------------------------------------------------------
    def on_scan(self, event):
        code = self.entry.get().strip()
        self.entry.delete(0, tk.END)

        if not code:
            return

        if code.startswith('2200'):
            self.process_worker(code)
        else:
            self.process_detal(code)

    # ------------------------------------------------------------------
    # Обработка сотрудника
    # ------------------------------------------------------------------
    def process_worker(self, rab):
        try:
            fio = get_worker_fio(self.cursor, rab)
            if fio is None:
                msg = f"Работник с кодом {rab} не найден в базе."
                log_message(f"ОШИБКА: {msg}")
                self.show_auto_message(msg, "Ошибка", 5000)
                self.conn.rollback()
                return

            open_session = get_open_session(self.cursor, rab)

            if open_session:
                old_ip, start_time = open_session

                if old_ip == self.current_ip:
                    close_session(self.cursor, rab)
                    if not decrement_workplace_cur(self.cursor, self.current_ip):
                        self.conn.rollback()
                        msg = f"Счётчик RAB_CUR на месте {self.workplace_name} уже равен 0. Операция отменена."
                        log_message(f"ОШИБКА: {msg}")
                        self.show_auto_message(msg, "Ошибка", 5000)
                        return
                    self.conn.commit()
                    msg = f"Работник {fio} завершил работу на {self.workplace_name}."
                    log_message(msg)
                    self.show_auto_message(msg, "Информация", 5000)
                else:
                    info = get_workplace_info(self.cursor, self.current_ip)
                    if info is None:
                        self.conn.rollback()
                        msg = "Текущее рабочее место не найдено."
                        log_message(f"ОШИБКА: {msg}")
                        self.show_auto_message(msg, "Ошибка", 5000)
                        return
                    name, max_cur, cur_cur, gr = info

                    if cur_cur >= max_cur:
                        msg = f"Место '{name}' заполнено (занято {cur_cur} из {max_cur}). Перемещение невозможно."
                        log_message(f"ПРЕДУПРЕЖДЕНИЕ: {msg}")
                        self.show_auto_message(msg, "Предупреждение", 5000)
                        self.conn.rollback()
                        return

                    old_place_name = get_workplace_name(self.cursor, old_ip)

                    close_session(self.cursor, rab)
                    if not decrement_workplace_cur(self.cursor, old_ip):
                        self.conn.rollback()
                        msg = f"Счётчик RAB_CUR на старом месте уже равен 0. Операция отменена."
                        log_message(f"ОШИБКА: {msg}")
                        self.show_auto_message(msg, "Ошибка", 5000)
                        return
                    create_session(self.cursor, rab, self.current_ip)
                    increment_workplace_cur(self.cursor, self.current_ip)
                    self.conn.commit()
                    msg = f"Работник {fio} перемещён с '{old_place_name}' на '{name}'."
                    log_message(msg)
                    self.show_auto_message(msg, "Информация", 5000)
            else:
                info = get_workplace_info(self.cursor, self.current_ip)
                if info is None:
                    self.conn.rollback()
                    msg = "Текущее рабочее место не найдено."
                    log_message(f"ОШИБКА: {msg}")
                    self.show_auto_message(msg, "Ошибка", 5000)
                    return
                name, max_cur, cur_cur, gr = info

                if cur_cur >= max_cur:
                    msg = f"Место '{name}' заполнено (занято {cur_cur} из {max_cur}). Регистрация невозможна."
                    log_message(f"ПРЕДУПРЕЖДЕНИЕ: {msg}")
                    self.show_auto_message(msg, "Предупреждение", 5000)
                    self.conn.rollback()
                    return

                create_session(self.cursor, rab, self.current_ip)
                increment_workplace_cur(self.cursor, self.current_ip)
                self.conn.commit()
                msg = f"Работник {fio} зарегистрирован на {name}."
                log_message(msg)
                self.show_auto_message(msg, "Информация", 5000)

            self.update_workers_list()

        except Exception as e:
            self.conn.rollback()
            msg = f"Ошибка обработки сотрудника: {e}"
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000)

    # ------------------------------------------------------------------
    # Обработка детали (без успешных сообщений)
    # ------------------------------------------------------------------
    def process_detal(self, detal):
        try:
            detal_info = get_detal_full_info(self.cursor, detal)
            if detal_info is None:
                msg = f"Деталь с кодом {detal} не найдена в CCALC."
                log_message(f"ОШИБКА: {msg}")
                self.show_auto_message(msg, "Ошибка", 5000)
                self.conn.rollback()
                return

            open_session = get_open_detal_session_any(self.cursor, detal)

            if open_session:
                old_ip, start_time = open_session

                if old_ip == self.current_ip:
                    msg = (f"Деталь {detal} (ZAKNUM:{detal_info['zaknum']}, "
                           f"KONNAME:{detal_info['konname']}) уже зарегистрирована на {self.workplace_name}.")
                    log_message(f"ПРЕДУПРЕЖДЕНИЕ: {msg}")
                    self.show_auto_message(msg, "Предупреждение", 5000)
                    self.conn.rollback()
                    return
                else:
                    close_detal_session(self.cursor, detal, old_ip)
                    create_detal_session(self.cursor, detal, self.current_ip, self.gr)
                    self.conn.commit()
                    # Успешное перемещение – сообщение не выводим
                    self.set_detal_info(detal_info)  # обновляем информацию на экране
                    log_message(f"Деталь {detal} перемещена с {old_ip} на {self.current_ip}")
            else:
                create_detal_session(self.cursor, detal, self.current_ip, self.gr)
                self.conn.commit()
                # Успешная регистрация – сообщение не выводим
                self.set_detal_info(detal_info)
                log_message(f"Деталь {detal} зарегистрирована на {self.current_ip}")

            self.update_workers_list()

        except Exception as e:
            self.conn.rollback()
            msg = f"Ошибка обработки детали: {e}"
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000)

    # ------------------------------------------------------------------
    # Закрытие приложения
    # ------------------------------------------------------------------
    def on_closing(self):
        log_message("=== Завершение программы ===")
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
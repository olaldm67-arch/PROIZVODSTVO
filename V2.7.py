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
# Функции для работы с деталями и статистикой
# ----------------------------------------------------------------------
def get_detal_full_info(cursor, detal):
    cursor.execute(
        "SELECT ZAKNUM, KONNAME, PRIMPROIZV, MATNAME, MATCOLOR, LENGTH, BCI, GRI "
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
            'length': row[5],
            'bci': row[6],
            'gri': row[7]
        }
    return None

def get_open_detal_session_at_ip(cursor, detal, ip):
    cursor.execute(
        "SELECT TIME_START_D FROM DETAL_TIME "
        "WHERE DETAL = %s AND IP = %s AND TIME_END_D IS NULL",
        (detal, ip)
    )
    return cursor.fetchone()

def has_open_detal_sessions_for_bci_gri(cursor, bci, gri, ip, exclude_detal=None):
    if exclude_detal:
        cursor.execute(
            "SELECT COUNT(*) FROM DETAL_TIME dt "
            "JOIN CCALC cc ON dt.DETAL = cc.DETAL "
            "WHERE dt.IP = %s AND dt.TIME_END_D IS NULL "
            "AND cc.BCI = %s AND cc.GRI = %s AND dt.DETAL != %s",
            (ip, bci, gri, exclude_detal)
        )
    else:
        cursor.execute(
            "SELECT COUNT(*) FROM DETAL_TIME dt "
            "JOIN CCALC cc ON dt.DETAL = cc.DETAL "
            "WHERE dt.IP = %s AND dt.TIME_END_D IS NULL "
            "AND cc.BCI = %s AND cc.GRI = %s",
            (ip, bci, gri)
        )
    count = cursor.fetchone()[0]
    return count > 0

def get_dimensions_and_perimeter_area_for_bci_gri(cursor, bci, gri):
    # Сначала пробуем взять раму (GRI='r')
    cursor.execute(
        "SELECT ORIENT, LENGTH FROM CCALC WHERE BCI = %s AND GRI = 'r' AND ORIENT IN ('Н', 'П')",
        (bci,)
    )
    rows = cursor.fetchall()
    width_mm = None
    height_mm = None
    if rows:
        for orient, length in rows:
            if orient == 'Н':
                width_mm = length
            elif orient == 'П':
                height_mm = length
        if width_mm is not None and height_mm is not None:
            w_m = (width_mm - 6) / 1000.0
            h_m = (height_mm - 6) / 1000.0
            perimeter = 2 * (w_m + h_m)
            area = w_m * h_m
            return width_mm, height_mm, perimeter, area

    # Если рамы нет, пробуем створку (GRI='s')
    cursor.execute(
        "SELECT ORIENT, LENGTH FROM CCALC WHERE BCI = %s AND GRI = 's' AND ORIENT IN ('Н', 'П')",
        (bci,)
    )
    rows = cursor.fetchall()
    if rows:
        for orient, length in rows:
            if orient == 'Н':
                width_mm = length
            elif orient == 'П':
                height_mm = length
        if width_mm is not None and height_mm is not None:
            w_m = (width_mm - 6) / 1000.0
            h_m = (height_mm - 6) / 1000.0
            perimeter = 2 * (w_m + h_m)
            area = w_m * h_m
            return width_mm, height_mm, perimeter, area

    # Если и створки нет, берём любую другую группу
    cursor.execute(
        "SELECT ORIENT, LENGTH FROM CCALC WHERE BCI = %s AND ORIENT IN ('Н', 'П')",
        (bci,)
    )
    rows = cursor.fetchall()
    if rows:
        for orient, length in rows:
            if orient == 'Н':
                width_mm = length
            elif orient == 'П':
                height_mm = length
        if width_mm is not None and height_mm is not None:
            w_m = (width_mm - 6) / 1000.0
            h_m = (height_mm - 6) / 1000.0
            perimeter = 2 * (w_m + h_m)
            area = w_m * h_m
            return width_mm, height_mm, perimeter, area

    # Если ничего не найдено
    return None, None, None, None

def create_detal_session(cursor, detal, ip, gr, bci, gri, perimeter, area):
    cursor.execute(
        "INSERT INTO DETAL_TIME (IP, DETAL, TIME_START_D, TIME_END_D, GR, BCI, GRI, PP, SS) "
        "VALUES (%s, %s, NOW(), NULL, %s, %s, %s, %s, %s)",
        (ip, detal, gr, bci, gri, perimeter, area)
    )

def get_stats_for_shift(cursor, ip):
    cursor.execute(
        "SELECT COUNT(DISTINCT BCI), SUM(PP), SUM(SS) "
        "FROM DETAL_TIME "
        "WHERE IP = %s AND DATE(TIME_START_D) = CURDATE() AND BCI IS NOT NULL",
        (ip,)
    )
    row = cursor.fetchone()
    count = row[0] if row[0] is not None else 0
    total_perimeter = row[1] if row[1] is not None else 0.0
    total_area = row[2] if row[2] is not None else 0.0
    return count, total_perimeter, total_area

# ----------------------------------------------------------------------
# Основной класс приложения
# ----------------------------------------------------------------------
class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт рабочего времени и деталей")
        self.root.state('zoomed')

        self.root.bind('<Button-1>', self.on_click)
        self.root.bind('<FocusOut>', self.on_focus_out)

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
            self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')
            sys.exit(1)

        workplace = get_workplace_info(self.cursor, self.current_ip)
        if not workplace:
            msg = f"Рабочее место с IP {self.current_ip} не найдено в БД."
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')
            sys.exit(1)

        self.workplace_name, self.rab_max, self.rab_cur, self.gr = workplace
        log_message(f"Рабочее место: {self.workplace_name}, GR={self.gr}, RAB_MAX={self.rab_max}, RAB_CUR={self.rab_cur}")

        # Определяем разрешённые группы для данного рабочего места по GR
        if self.gr in (1, 2, 5):
            self.allowed_gri = ['r']
        elif self.gr in (3, 4):
            self.allowed_gri = ['s']
        elif self.gr in (5, 6):
            self.allowed_gri = ['r', 's']
        else:
            self.allowed_gri = []

        group_names = {'r': 'РАМЫ', 's': 'СТВОРКИ'}
        if len(self.allowed_gri) == 0:
            allowed_desc = "не разрешена регистрация деталей"
        elif len(self.allowed_gri) == 1:
            allowed_desc = f"разрешена только {group_names[self.allowed_gri[0]]}"
        else:
            allowed_desc = "разрешены РАМЫ и СТВОРКИ"
        log_message(f"Разрешённые группы: {allowed_desc}")

        # ------------------- Интерфейс -------------------
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

        # Блок информации о детали (5 строк)
        self.detal_frame = ttk.LabelFrame(root, text="Информация о детали / изделии", padding=10)
        self.detal_frame.pack(fill=tk.X, padx=10, pady=5)

        self.info_line1 = tk.Label(self.detal_frame, text="", font=("Arial", 20, "bold"), fg="darkblue", anchor='w')
        self.info_line2 = tk.Label(self.detal_frame, text="", font=("Arial", 20, "bold"), fg="darkblue", anchor='w')
        self.info_line3 = tk.Label(self.detal_frame, text="", font=("Arial", 20, "bold"), fg="darkblue", anchor='w')
        self.info_line4 = tk.Label(self.detal_frame, text="", font=("Arial", 18, "bold"), fg="green", anchor='w')
        self.info_line5 = tk.Label(self.detal_frame, text="", font=("Arial", 18, "bold"), fg="darkred", anchor='w')

        self.info_line1.pack(fill=tk.X, padx=10, pady=2, anchor='w')
        self.info_line2.pack(fill=tk.X, padx=10, pady=2, anchor='w')
        self.info_line3.pack(fill=tk.X, padx=10, pady=2, anchor='w')
        self.info_line4.pack(fill=tk.X, padx=10, pady=2, anchor='w')
        self.info_line5.pack(fill=tk.X, padx=10, pady=2, anchor='w')

        # Блок статистики за смену
        self.stats_frame = ttk.LabelFrame(root, text="Статистика за смену", padding=10)
        self.stats_frame.pack(fill=tk.X, padx=10, pady=5)

        self.stats_line1 = tk.Label(self.stats_frame, text="Количество изделий: 0", font=("Arial", 16), anchor='w')
        self.stats_line2 = tk.Label(self.stats_frame, text="Суммарный периметр: 0.000 м", font=("Arial", 16), anchor='w')
        self.stats_line3 = tk.Label(self.stats_frame, text="Суммарная площадь: 0.000 м²", font=("Arial", 16), anchor='w')
        self.stats_line1.pack(fill=tk.X, padx=10, pady=2, anchor='w')
        self.stats_line2.pack(fill=tk.X, padx=10, pady=2, anchor='w')
        self.stats_line3.pack(fill=tk.X, padx=10, pady=2, anchor='w')

        ttk.Separator(root, orient='horizontal').pack(fill=tk.X, padx=10, pady=5)

        # Таблица сотрудников
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

        # Нижняя панель: поле ввода и кнопка выхода
        bottom_frame = tk.Frame(root)
        bottom_frame.pack(pady=10, fill=tk.X)

        self.entry_label = tk.Label(bottom_frame, text="Сканируйте штрих-код (или введите вручную):", font=("Arial", 14))
        self.entry_label.pack(side=tk.LEFT, padx=10)

        self.entry = tk.Entry(bottom_frame, font=("Arial", 16), width=25)
        self.entry.pack(side=tk.LEFT, padx=10)
        self.entry.bind("<Return>", self.on_scan)
        self.entry.focus_set()

        self.exit_button = tk.Button(bottom_frame, text="Выход", command=self.on_closing, font=("Arial", 14))
        self.exit_button.pack(side=tk.LEFT, padx=10)

        # Инициализация
        self.close_previous_sessions()
        self.update_workers_list()
        self.auto_update()
        self.clear_detal_info()
        self.update_stats_panel()

    # ------------------------------------------------------------------
    # Обработчики событий для возврата фокуса на поле ввода
    # ------------------------------------------------------------------
    def on_click(self, event):
        if event.widget != self.entry and event.widget != self.exit_button:
            self.entry.focus_set()

    def on_focus_out(self, event):
        self.root.after(10, self.entry.focus_set)

    # ------------------------------------------------------------------
    # Всплывающее окно без кнопок с заданным стилем, увеличенное и центрированное
    # ------------------------------------------------------------------
    def show_auto_message(self, message, title="Информация", timeout=5000, msg_type='info'):
        """
        Создаёт окно-сообщение без кнопок с центрированным текстом крупным полужирным шрифтом.
        msg_type: 'info' - зелёная рамка, 'error' - красная рамка
        """
        win = tk.Toplevel(self.root)
        win.title(title)
        win_width = 600
        win_height = 200
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()
        x = (screen_width - win_width) // 2
        y = (screen_height - win_height) // 2
        win.geometry(f"{win_width}x{win_height}+{x}+{y}")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        border_color = "red" if msg_type == 'error' else "green"

        frame = tk.Frame(win, bd=5, relief='solid',
                         highlightbackground=border_color,
                         highlightcolor=border_color,
                         highlightthickness=3)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        lbl = tk.Label(frame, text=message, font=("Arial", 20, "bold"),
                       justify='center', wraplength=550)
        lbl.pack(expand=True, padx=10, pady=10)

        win.after(timeout, lambda: (win.destroy(), self.entry.focus_set()))

    # ------------------------------------------------------------------
    # Очистка информационного блока детали
    # ------------------------------------------------------------------
    def clear_detal_info(self):
        self.info_line1.config(text="")
        self.info_line2.config(text="")
        self.info_line3.config(text="")
        self.info_line4.config(text="")
        self.info_line5.config(text="")

    def set_detal_info(self, detal_info, width_mm, height_mm, perimeter_m, area_m2):
        line1 = f"{detal_info.get('zaknum', '')}  {detal_info.get('konname', '')}  {detal_info.get('matcolor', '')}"
        length = detal_info.get('length', 0)
        length_minus_6 = length - 6 if length is not None else 0
        line2 = f"{detal_info.get('matname', '')}  {length_minus_6}"
        line3 = detal_info.get('primproizv', '')
        gri = detal_info.get('gri', '')
        if gri == 'r':
            type_text = "Тип: РАМА"
        elif gri == 's':
            type_text = "Тип: СТВОРКА"
        elif gri == '3':
            type_text = "Тип: ИМПОСТ"
        else:
            type_text = f"Тип: {gri}"

        if width_mm is not None and height_mm is not None:
            w_mm = width_mm - 6
            h_mm = height_mm - 6
            w_m = w_mm / 1000.0
            h_m = h_mm / 1000.0
            dims_text = (f"Ширина: {w_m:.3f} м  Высота: {h_m:.3f} м  "
                         f"Периметр: {perimeter_m:.3f} м  Площадь: {area_m2:.3f} м²")
        else:
            dims_text = "Габариты: данные отсутствуют (нет ORIENT='Н' или 'П')"

        self.info_line1.config(text=line1)
        self.info_line2.config(text=line2)
        self.info_line3.config(text=line3)
        self.info_line4.config(text=type_text)
        self.info_line5.config(text=dims_text)

    # ------------------------------------------------------------------
    # Обновление панели статистики
    # ------------------------------------------------------------------
    def update_stats_panel(self):
        try:
            count, total_perimeter, total_area = get_stats_for_shift(self.cursor, self.current_ip)
            self.stats_line1.config(text=f"Количество изделий: {count}")
            self.stats_line2.config(text=f"Суммарный периметр: {total_perimeter:.3f} м")
            self.stats_line3.config(text=f"Суммарная площадь: {total_area:.3f} м²")
            self.entry.focus_set()
        except Exception as e:
            msg = f"Ошибка получения статистики: {e}"
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')

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
                self.show_auto_message(msg, "Информация", 5000, msg_type='info')
            else:
                self.entry.focus_set()
        except Exception as e:
            self.conn.rollback()
            msg = f"Не удалось закрыть предыдущие сессии сотрудников: {e}"
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')

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
            self.entry.focus_set()
        except Exception as e:
            msg = f"Не удалось обновить список сотрудников: {e}"
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')

    # ------------------------------------------------------------------
    # Обработчик сканирования
    # ------------------------------------------------------------------
    def on_scan(self, event):
        code = self.entry.get().strip()
        self.entry.delete(0, tk.END)
        if not code:
            self.entry.focus_set()
            return
        if code.startswith('2200'):
            self.process_worker(code)
        else:
            self.process_detal(code)
        self.entry.focus_set()

    # ------------------------------------------------------------------
    # Обработка сотрудника
    # ------------------------------------------------------------------
    def process_worker(self, rab):
        try:
            fio = get_worker_fio(self.cursor, rab)
            if fio is None:
                msg = f"Работник с кодом {rab} не найден в базе."
                log_message(f"ОШИБКА: {msg}")
                self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')
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
                        self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')
                        return
                    self.conn.commit()
                    msg = f"Работник {fio} завершил работу на {self.workplace_name}."
                    log_message(msg)
                    self.show_auto_message(msg, "Информация", 5000, msg_type='info')
                else:
                    info = get_workplace_info(self.cursor, self.current_ip)
                    if info is None:
                        self.conn.rollback()
                        msg = "Текущее рабочее место не найдено."
                        log_message(f"ОШИБКА: {msg}")
                        self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')
                        return
                    name, max_cur, cur_cur, gr = info

                    if cur_cur >= max_cur:
                        msg = f"Место '{name}' заполнено (занято {cur_cur} из {max_cur}). Перемещение невозможно."
                        log_message(f"ПРЕДУПРЕЖДЕНИЕ: {msg}")
                        self.show_auto_message(msg, "Предупреждение", 5000, msg_type='error')
                        self.conn.rollback()
                        return

                    old_place_name = get_workplace_name(self.cursor, old_ip)

                    close_session(self.cursor, rab)
                    if not decrement_workplace_cur(self.cursor, old_ip):
                        self.conn.rollback()
                        msg = f"Счётчик RAB_CUR на старом месте уже равен 0. Операция отменена."
                        log_message(f"ОШИБКА: {msg}")
                        self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')
                        return
                    create_session(self.cursor, rab, self.current_ip)
                    increment_workplace_cur(self.cursor, self.current_ip)
                    self.conn.commit()
                    msg = f"Работник {fio} перемещён с '{old_place_name}' на '{name}'."
                    log_message(msg)
                    self.show_auto_message(msg, "Информация", 5000, msg_type='info')
            else:
                info = get_workplace_info(self.cursor, self.current_ip)
                if info is None:
                    self.conn.rollback()
                    msg = "Текущее рабочее место не найдено."
                    log_message(f"ОШИБКА: {msg}")
                    self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')
                    return
                name, max_cur, cur_cur, gr = info

                if cur_cur >= max_cur:
                    msg = f"Место '{name}' заполнено (занято {cur_cur} из {max_cur}). Регистрация невозможна."
                    log_message(f"ПРЕДУПРЕЖДЕНИЕ: {msg}")
                    self.show_auto_message(msg, "Предупреждение", 5000, msg_type='error')
                    self.conn.rollback()
                    return

                create_session(self.cursor, rab, self.current_ip)
                increment_workplace_cur(self.cursor, self.current_ip)
                self.conn.commit()
                msg = f"Работник {fio} зарегистрирован на {name}."
                log_message(msg)
                self.show_auto_message(msg, "Информация", 5000, msg_type='info')

            self.update_workers_list()

        except Exception as e:
            self.conn.rollback()
            msg = f"Ошибка обработки сотрудника: {e}"
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')

    # ------------------------------------------------------------------
    # Обработка детали (основная логика)
    # ------------------------------------------------------------------
    def process_detal(self, detal):
        try:
            detal_info = get_detal_full_info(self.cursor, detal)
            if detal_info is None:
                msg = f"Деталь с кодом {detal} не найдена в CCALC."
                log_message(f"ОШИБКА: {msg}")
                self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')
                self.conn.rollback()
                self.clear_detal_info()
                return

            bci = detal_info['bci']
            gri = detal_info['gri']

            if gri == '3':
                msg = "Для регистрации отсканируйте деталь рамы или створки. Импост не регистрируется."
                log_message(f"ПРЕДУПРЕЖДЕНИЕ: {msg}")
                self.show_auto_message(msg, "Предупреждение", 5000, msg_type='error')
                self.conn.rollback()
                self.clear_detal_info()
                return

            if gri not in self.allowed_gri:
                group_names = {'r': 'РАМЫ', 's': 'СТВОРКИ'}
                allowed_names = ', '.join([group_names.get(g, g) for g in self.allowed_gri])
                if not self.allowed_gri:
                    allowed_names = "не разрешена регистрация деталей"
                msg = f"На данном рабочем месте разрешена регистрация только {allowed_names}, а вы сканируете {group_names.get(gri, gri)}."
                log_message(f"ПРЕДУПРЕЖДЕНИЕ: {msg}")
                self.show_auto_message(msg, "Предупреждение", 5000, msg_type='error')
                self.conn.rollback()
                self.clear_detal_info()
                return

            open_session = get_open_detal_session_at_ip(self.cursor, detal, self.current_ip)
            if open_session:
                msg = (f"Деталь {detal} (ZAKNUM:{detal_info['zaknum']}) уже зарегистрирована на {self.workplace_name}.")
                log_message(f"ПРЕДУПРЕЖДЕНИЕ: {msg}")
                self.show_auto_message(msg, "Предупреждение", 5000, msg_type='error')
                self.conn.rollback()
                width_mm, height_mm, perimeter, area = get_dimensions_and_perimeter_area_for_bci_gri(self.cursor, bci, gri)
                self.set_detal_info(detal_info, width_mm, height_mm, perimeter, area)
                return

            if has_open_detal_sessions_for_bci_gri(self.cursor, bci, gri, self.current_ip, exclude_detal=detal):
                msg = (f"Изделие (BCI={bci}, тип={gri}) уже зарегистрировано на {self.workplace_name} "
                       f"(другая деталь). Регистрация новой детали невозможна.")
                log_message(f"ПРЕДУПРЕЖДЕНИЕ: {msg}")
                self.show_auto_message(msg, "Предупреждение", 5000, msg_type='error')
                self.conn.rollback()
                return

            width_mm, height_mm, perimeter, area = get_dimensions_and_perimeter_area_for_bci_gri(self.cursor, bci, gri)
            if width_mm is None or height_mm is None:
                msg = "Не удалось определить габариты изделия (нет ORIENT='Н' или 'П')."
                log_message(f"ОШИБКА: {msg}")
                self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')
                self.conn.rollback()
                return

            create_detal_session(self.cursor, detal, self.current_ip, self.gr, bci, gri, perimeter, area)
            self.conn.commit()
            log_message(f"Деталь {detal} зарегистрирована на IP {self.current_ip} (BCI={bci}, GRI={gri})")
            self.set_detal_info(detal_info, width_mm, height_mm, perimeter, area)

            self.update_workers_list()
            self.update_stats_panel()

        except Exception as e:
            self.conn.rollback()
            msg = f"Ошибка обработки детали: {e}"
            log_message(f"ОШИБКА: {msg}")
            self.show_auto_message(msg, "Ошибка", 5000, msg_type='error')

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
#!/usr/bin/env -S python3
# -*- coding: utf-8 -*-
"""
Программа анализа таблиц WORKER_TIME и DETAL_TIME.
Вкладки:
  1. Работники
  2. Детали
  3. Заказы
  4. Табель
  5. Проверка  — поиск и исправление неправильно закрытых сессий
"""

import sys
import os
import csv
import pymysql
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, date, timedelta

# ----------------------------------------------------------------------
# Параметры подключения к базе данных
# ----------------------------------------------------------------------
DB_HOST = '192.168.0.10'
DB_USER = 'workshop'
DB_PASSWORD = 'w0rK5h0p'
DB_NAME = 'workshop'
DB_CHARSET = 'utf8'

ALLOWED_WORKPLACE_IPS = (201, 202, 203, 204, 205, 206)
EXCLUDED_WORKPLACES = ('ТЕСТ0',)
EXCLUDED_WORKPLACE_IPS = (250,)
ORDERS_DAYS_BACK = 90
TAXEL_DAYS_DEFAULT = 30


# ----------------------------------------------------------------------
def connect_db():
    return pymysql.connect(
        host=DB_HOST, user=DB_USER, password=DB_PASSWORD,
        database=DB_NAME, charset=DB_CHARSET, autocommit=False
    )


# ----------------------------------------------------------------------
def get_workplaces(cursor):
    cursor.execute("SELECT IP, NAME FROM WORKPLACE ORDER BY IP")
    rows = cursor.fetchall()
    result = []
    for ip, name in rows:
        if ip not in ALLOWED_WORKPLACE_IPS:
            continue
        if ip in EXCLUDED_WORKPLACE_IPS:
            continue
        if name and name.strip() in EXCLUDED_WORKPLACES:
            continue
        result.append((ip, name))
    return result


# ----------------------------------------------------------------------
def get_workers_data(cursor, target_date):
    workplaces = get_workplaces(cursor)
    cursor.execute("SELECT RAB, FIO FROM WORKER ORDER BY FIO")
    workers = cursor.fetchall()
    allowed_ips = tuple(ip for ip, _ in workplaces) or (0,)
    placeholders = ','.join(['%s'] * len(allowed_ips))
    cursor.execute(
        f"SELECT wt.RAB, wt.IP, MIN(wt.TIME_START_W) "
        f"FROM WORKER_TIME wt "
        f"WHERE DATE(wt.TIME_START_W) = %s AND wt.IP IN ({placeholders}) "
        f"GROUP BY wt.RAB, wt.IP",
        (target_date, *allowed_ips)
    )
    cells = {}
    for rab, ip, start_time in cursor.fetchall():
        cells[(rab, ip)] = start_time
    return workplaces, workers, cells


# ----------------------------------------------------------------------
def get_detals_data(cursor, target_date):
    workplaces = get_workplaces(cursor)
    allowed_ips = tuple(ip for ip, _ in workplaces) or (0,)
    placeholders = ','.join(['%s'] * len(allowed_ips))
    cursor.execute(
        f"SELECT DISTINCT cc.ZAKNUM, cc.KONNAME, cc.NUMIZDELIE "
        f"FROM DETAL_TIME dt JOIN CCALC cc ON dt.DETAL = cc.DETAL "
        f"WHERE DATE(dt.TIME_START_D) = %s AND dt.IP IN ({placeholders}) "
        f"ORDER BY cc.ZAKNUM, cc.KONNAME, cc.NUMIZDELIE",
        (target_date, *allowed_ips)
    )
    detals = cursor.fetchall()
    cursor.execute(
        f"SELECT cc.ZAKNUM, cc.KONNAME, cc.NUMIZDELIE, dt.IP, MIN(dt.TIME_START_D) "
        f"FROM DETAL_TIME dt JOIN CCALC cc ON dt.DETAL = cc.DETAL "
        f"WHERE DATE(dt.TIME_START_D) = %s AND dt.IP IN ({placeholders}) "
        f"GROUP BY cc.ZAKNUM, cc.KONNAME, cc.NUMIZDELIE, dt.IP",
        (target_date, *allowed_ips)
    )
    cells = {}
    for zaknum, konname, numizdelie, ip, start_time in cursor.fetchall():
        cells[(zaknum, konname, numizdelie, ip)] = start_time
    return workplaces, detals, cells


# ----------------------------------------------------------------------
def get_orders_data(cursor, zaknum_filter=None, days_back=ORDERS_DAYS_BACK):
    workplaces = get_workplaces(cursor)
    allowed_ips = tuple(ip for ip, _ in workplaces) or (0,)
    placeholders = ','.join(['%s'] * len(allowed_ips))
    base_query = (
        f"SELECT cc.ZAKNUM, cc.KONNAME, cc.NUMIZDELIE, dt.IP, MAX(dt.TIME_START_D) "
        f"FROM DETAL_TIME dt JOIN CCALC cc ON dt.DETAL = cc.DETAL "
        f"WHERE dt.TIME_START_D >= DATE_SUB(CURDATE(), INTERVAL %s DAY) "
        f"AND dt.IP IN ({placeholders}) "
    )
    params = [days_back, *allowed_ips]
    if zaknum_filter:
        base_query += " AND cc.ZAKNUM LIKE %s "
        params.append(f"%{zaknum_filter}%")
    base_query += " GROUP BY cc.ZAKNUM, cc.KONNAME, cc.NUMIZDELIE, dt.IP "
    base_query += " ORDER BY cc.ZAKNUM, cc.KONNAME, cc.NUMIZDELIE, dt.IP "
    cursor.execute(base_query, params)
    cells = {}
    orders_set = set()
    for zaknum, konname, numizdelie, ip, start_time in cursor.fetchall():
        cells[(zaknum, konname, numizdelie, ip)] = start_time
        orders_set.add((zaknum, konname, numizdelie))
    return workplaces, sorted(orders_set), cells


# ----------------------------------------------------------------------
def get_timesheet_data(cursor, date_from, date_to):
    days = (date_to - date_from).days + 1
    dates_list = [date_from + timedelta(days=i) for i in range(days)]

    cursor.execute("SELECT RAB, FIO FROM WORKER ORDER BY FIO")
    workers = cursor.fetchall()

    cursor.execute(
        "SELECT wt.RAB, DATE(wt.TIME_START_W), "
        "SUM(TIMESTAMPDIFF(SECOND, wt.TIME_START_W, IFNULL(wt.TIME_END_W, NOW()))) "
        "FROM WORKER_TIME wt "
        "WHERE DATE(wt.TIME_START_W) BETWEEN %s AND %s "
        "GROUP BY wt.RAB, DATE(wt.TIME_START_W)",
        (date_from, date_to)
    )
    cells = {}
    for rab, day, sec in cursor.fetchall():
        if sec is None:
            sec = 0
        hours = round(float(sec) / 3600.0, 1)
        if hours > 0:
            cells[(rab, day)] = hours
    return dates_list, workers, cells


# ----------------------------------------------------------------------
# НОВОЕ: поиск и исправление проблемных сессий
# ----------------------------------------------------------------------
def get_problematic_sessions(cursor):
    """
    Возвращает список сессий, у которых TIME_END_W позже 18:00 дня открытия.
    Такие сессии появились из-за того, что их закрывали текущим моментом (NOW())
    в последующие дни.
    """
    cursor.execute(
        "SELECT wt.RAB, w.FIO, wt.IP, wp.NAME, wt.TIME_START_W, wt.TIME_END_W, "
        "       TIMESTAMPDIFF(MINUTE, wt.TIME_START_W, wt.TIME_END_W) AS minutes "
        "FROM WORKER_TIME wt "
        "LEFT JOIN WORKER w ON wt.RAB = w.RAB "
        "LEFT JOIN WORKPLACE wp ON wt.IP = wp.IP "
        "WHERE wt.TIME_END_W IS NOT NULL "
        "  AND wt.TIME_END_W > TIMESTAMP(DATE(wt.TIME_START_W), '18:00:00') "
        "ORDER BY wt.TIME_START_W DESC"
    )
    return cursor.fetchall()


def fix_problematic_sessions(cursor, conn):
    """
    Исправляет все проблемные сессии, у которых дата открытия РАНЬШЕ сегодняшней:
    ставит TIME_END_W = 18:00 дня открытия.
    Возвращает число исправленных записей.
    """
    cursor.execute(
        "UPDATE WORKER_TIME "
        "SET TIME_END_W = TIMESTAMP(DATE(TIME_START_W), '18:00:00') "
        "WHERE TIME_END_W IS NOT NULL "
        "  AND TIME_END_W > TIMESTAMP(DATE(TIME_START_W), '18:00:00') "
        "  AND DATE(TIME_START_W) < CURDATE()"
    )
    affected = cursor.rowcount
    conn.commit()
    return affected


# ----------------------------------------------------------------------
class AnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализ WORKER_TIME и DETAL_TIME")
        self.root.geometry("1600x900")
        self.root.state('zoomed')

        self.conn = None
        self.cursor = None
        try:
            self.conn = connect_db()
            self.cursor = self.conn.cursor()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться к БД:\n{e}")
            sys.exit(1)

        self.taxel_date_to = date.today()
        self.taxel_date_from = self.taxel_date_to - timedelta(days=TAXEL_DAYS_DEFAULT - 1)

        # --- Верхняя панель ---
        top = tk.Frame(root)
        top.pack(fill=tk.X, padx=10, pady=5)

        tk.Button(top, text="◀ Вчера", font=("Arial", 12),
                  command=self.prev_day).pack(side=tk.LEFT, padx=2)

        tk.Label(top, text="Дата:", font=("Arial", 12)).pack(side=tk.LEFT, padx=(15, 3))
        self.date_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))
        self.date_entry = tk.Entry(top, textvariable=self.date_var,
                                   font=("Arial", 12), width=12, justify='center')
        self.date_entry.pack(side=tk.LEFT, padx=3)
        self.date_entry.bind("<Return>", lambda e: self.refresh_all())

        tk.Button(top, text="Сегодня", font=("Arial", 12),
                  command=self.today).pack(side=tk.LEFT, padx=2)
        tk.Button(top, text="Завтра ▶", font=("Arial", 12),
                  command=self.next_day).pack(side=tk.LEFT, padx=2)

        try:
            from tkcalendar import DateEntry
            self.cal = DateEntry(top, width=12, font=("Arial", 12),
                                 date_pattern='yyyy-mm-dd',
                                 background='darkblue', foreground='white',
                                 borderwidth=2)
            self.cal.pack(side=tk.LEFT, padx=(15, 5))
            self.cal.bind("<<DateEntrySelected>>", self.on_calendar_select)
        except ImportError:
            self.cal = None

        tk.Button(top, text="Обновить", font=("Arial", 12),
                  command=self.refresh_all).pack(side=tk.LEFT, padx=10)
        tk.Button(top, text="Экспорт в CSV", font=("Arial", 12),
                  command=self.export_current_tab).pack(side=tk.LEFT, padx=10)

        self.status = tk.Label(top, text="", font=("Arial", 12), fg="blue")
        self.status.pack(side=tk.RIGHT, padx=10)

        # --- Notebook ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Вкладка 1: Работники
        self.tab_workers = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_workers, text="Работники")
        self.tree_workers = self.create_tree(self.tab_workers)

        # Вкладка 2: Детали
        self.tab_detals = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_detals, text="Детали")
        self.tree_detals = self.create_tree(self.tab_detals)

        # Вкладка 3: Заказы
        self.tab_orders = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_orders, text="Заказы")
        search_frame = tk.Frame(self.tab_orders)
        search_frame.pack(fill=tk.X, padx=5, pady=5)
        tk.Label(search_frame, text="Поиск по номеру заказа:",
                 font=("Arial", 12)).pack(side=tk.LEFT, padx=(5, 3))
        self.orders_search_var = tk.StringVar()
        self.orders_search_entry = tk.Entry(search_frame,
                                            textvariable=self.orders_search_var,
                                            font=("Arial", 12), width=25)
        self.orders_search_entry.pack(side=tk.LEFT, padx=5)
        self.orders_search_entry.bind("<Return>", lambda e: self.refresh_orders())
        self.orders_search_var.trace_add("write", self.on_orders_search_change)
        tk.Button(search_frame, text="Очистить", font=("Arial", 11),
                  command=self.clear_orders_search).pack(side=tk.LEFT, padx=5)
        self.tree_orders = self.create_tree(self.tab_orders)

        # Вкладка 4: Табель
        self.tab_taxel = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_taxel, text="Табель")
        taxel_top = tk.Frame(self.tab_taxel)
        taxel_top.pack(fill=tk.X, padx=5, pady=5)
        tk.Label(taxel_top, text="Период:", font=("Arial", 12)).pack(side=tk.LEFT, padx=(5, 3))
        self.taxel_from_var = tk.StringVar(value=self.taxel_date_from.strftime("%Y-%m-%d"))
        self.taxel_to_var = tk.StringVar(value=self.taxel_date_to.strftime("%Y-%m-%d"))
        tk.Entry(taxel_top, textvariable=self.taxel_from_var,
                 font=("Arial", 12), width=12, justify='center').pack(side=tk.LEFT, padx=3)
        tk.Label(taxel_top, text="—", font=("Arial", 12)).pack(side=tk.LEFT)
        tk.Entry(taxel_top, textvariable=self.taxel_to_var,
                 font=("Arial", 12), width=12, justify='center').pack(side=tk.LEFT, padx=3)
        tk.Button(taxel_top, text="30 дней", font=("Arial", 11),
                  command=lambda: self.set_taxel_period(30)).pack(side=tk.LEFT, padx=5)
        tk.Button(taxel_top, text="Месяц", font=("Arial", 11),
                  command=lambda: self.set_taxel_period(31)).pack(side=tk.LEFT, padx=2)
        tk.Button(taxel_top, text="Квартал", font=("Arial", 11),
                  command=lambda: self.set_taxel_period(92)).pack(side=tk.LEFT, padx=2)
        tk.Button(taxel_top, text="60 дней", font=("Arial", 11),
                  command=lambda: self.set_taxel_period(60)).pack(side=tk.LEFT, padx=5)
        tk.Button(taxel_top, text="Обновить табель", font=("Arial", 11),
                  command=self.refresh_timesheet).pack(side=tk.LEFT, padx=10)
        self.tree_taxel = self.create_tree(self.tab_taxel)

        # Вкладка 5: Проверка
        self.tab_check = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_check, text="Проверка")

        check_top = tk.Frame(self.tab_check)
        check_top.pack(fill=tk.X, padx=5, pady=5)

        tk.Button(check_top, text="🔍 Найти проблемные сессии", font=("Arial", 12),
                  command=self.check_sessions).pack(side=tk.LEFT, padx=5)
        tk.Button(check_top, text="🛠 Исправить все (18:00 дня открытия)", font=("Arial", 12),
                  bg="#f4c842",
                  command=self.fix_sessions).pack(side=tk.LEFT, padx=15)

        tk.Label(check_top,
                 text="Проблема: TIME_END_W позже 18:00 дня открытия сессии.\n"
                      "Исправление не затрагивает сессии, начатые сегодня.",
                 font=("Arial", 10), fg="gray", justify='left').pack(side=tk.LEFT, padx=20)

        self.check_info = tk.Label(self.tab_check, text="", font=("Arial", 12), fg="blue")
        self.check_info.pack(fill=tk.X, padx=10, pady=(0, 5))

        self.tree_check = self.create_tree(self.tab_check)

        # Данные
        self.last_workers_data = None
        self.last_detals_data = None
        self._orders_search_job = None

        self.auto_refresh()

    # ------------------------------------------------------------------
    def prev_day(self):
        try:
            d = datetime.strptime(self.date_var.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            d = date.today()
        d -= timedelta(days=1)
        self.date_var.set(d.strftime("%Y-%m-%d"))
        if self.cal is not None:
            self.cal.set_date(d)
        self.refresh_all()

    def next_day(self):
        try:
            d = datetime.strptime(self.date_var.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            d = date.today()
        d += timedelta(days=1)
        self.date_var.set(d.strftime("%Y-%m-%d"))
        if self.cal is not None:
            self.cal.set_date(d)
        self.refresh_all()

    def today(self):
        d = date.today()
        self.date_var.set(d.strftime("%Y-%m-%d"))
        if self.cal is not None:
            self.cal.set_date(d)
        self.refresh_all()

    def on_calendar_select(self, event):
        if self.cal is None:
            return
        d = self.cal.get_date()
        self.date_var.set(d.strftime("%Y-%m-%d"))
        self.refresh_all()

    # ------------------------------------------------------------------
    def create_tree(self, parent):
        frame = tk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(frame, show="headings")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 12), rowheight=24)
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
        return tree

    # ------------------------------------------------------------------
    def refresh_all(self):
        try:
            target_date = datetime.strptime(self.date_var.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ГГГГ-ММ-ДД.")
            return
        self.refresh_workers(target_date)
        self.refresh_detals(target_date)
        self.refresh_orders()
        self.refresh_timesheet()
        self.status.config(text=f"Обновлено: {datetime.now().strftime('%H:%M:%S')} ({target_date})")

    def refresh_workers(self, target_date):
        workplaces, workers, cells = get_workers_data(self.cursor, target_date)
        self.last_workers_data = (workplaces, workers, cells)
        columns = ["fio"] + [f"ip_{ip}" for ip, _ in workplaces]
        tree = self.tree_workers
        tree.delete(*tree.get_children())
        tree["columns"] = columns
        tree.heading("fio", text="ФИО")
        tree.column("fio", width=250, anchor="w")
        for ip, name in workplaces:
            tree.heading(f"ip_{ip}", text=f"{name}\n({ip})")
            tree.column(f"ip_{ip}", width=110, anchor="center")
        for rab, fio in workers:
            row = [fio]
            has_any = False
            for ip, _ in workplaces:
                t = cells.get((rab, ip))
                if t:
                    row.append(t.strftime("%H:%M"))
                    has_any = True
                else:
                    row.append("")
            if has_any:
                tree.insert("", "end", values=row)

    def refresh_detals(self, target_date):
        workplaces, detals, cells = get_detals_data(self.cursor, target_date)
        self.last_detals_data = (workplaces, detals, cells)
        columns = ["info"] + [f"ip_{ip}" for ip, _ in workplaces]
        tree = self.tree_detals
        tree.delete(*tree.get_children())
        tree["columns"] = columns
        tree.heading("info", text="Заказ | Конструкция | Изделие")
        tree.column("info", width=420, anchor="w")
        for ip, name in workplaces:
            tree.heading(f"ip_{ip}", text=f"{name}\n({ip})")
            tree.column(f"ip_{ip}", width=110, anchor="center")
        for zaknum, konname, numizdelie in detals:
            row = [f"{zaknum} | {konname} | {numizdelie}"]
            for ip, _ in workplaces:
                t = cells.get((zaknum, konname, numizdelie, ip))
                row.append(t.strftime("%H:%M") if t else "0")
            tree.insert("", "end", values=row)

    def refresh_orders(self):
        try:
            zaknum_filter = self.orders_search_var.get().strip() or None
            workplaces, orders, cells = get_orders_data(
                self.cursor, zaknum_filter=zaknum_filter, days_back=ORDERS_DAYS_BACK)
        except Exception as e:
            self.status.config(text=f"Ошибка заказов: {e}")
            return
        columns = ["info"] + [f"ip_{ip}" for ip, _ in workplaces]
        tree = self.tree_orders
        tree.delete(*tree.get_children())
        tree["columns"] = columns
        tree.heading("info", text="Заказ | Конструкция | Изделие")
        tree.column("info", width=420, anchor="w")
        for ip, name in workplaces:
            tree.heading(f"ip_{ip}", text=f"{name}\n({ip})")
            tree.column(f"ip_{ip}", width=140, anchor="center")
        for zaknum, konname, numizdelie in orders:
            row = [f"{zaknum} | {konname} | {numizdelie}"]
            for ip, _ in workplaces:
                t = cells.get((zaknum, konname, numizdelie, ip))
                row.append(t.strftime("%d.%m %H:%M") if t else "0")
            tree.insert("", "end", values=row)

    def set_taxel_period(self, days):
        self.taxel_date_to = date.today()
        self.taxel_date_from = self.taxel_date_to - timedelta(days=days - 1)
        self.taxel_from_var.set(self.taxel_date_from.strftime("%Y-%m-%d"))
        self.taxel_to_var.set(self.taxel_date_to.strftime("%Y-%m-%d"))
        self.refresh_timesheet()

    def refresh_timesheet(self):
        try:
            date_from = datetime.strptime(self.taxel_from_var.get().strip(), "%Y-%m-%d").date()
            date_to = datetime.strptime(self.taxel_to_var.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат даты табеля.")
            return
        if date_from > date_to:
            date_from, date_to = date_to, date_from
            self.taxel_from_var.set(date_from.strftime("%Y-%m-%d"))
            self.taxel_to_var.set(date_to.strftime("%Y-%m-%d"))
        dates_list, workers, cells = get_timesheet_data(self.cursor, date_from, date_to)
        columns = ["fio"] + [d.strftime("%Y-%m-%d") for d in dates_list] + ["total"]
        tree = self.tree_taxel
        tree.delete(*tree.get_children())
        tree["columns"] = columns
        tree.heading("fio", text="ФИО")
        tree.column("fio", width=250, anchor="w")
        for d in dates_list:
            col = d.strftime("%Y-%m-%d")
            tree.heading(col, text=d.strftime("%d.%m"))
            tree.column(col, width=60, anchor="center")
        tree.heading("total", text="Всего")
        tree.column("total", width=70, anchor="center")
        for rab, fio in workers:
            row = [fio]
            total = 0.0
            has_any = False
            for d in dates_list:
                h = cells.get((rab, d))
                if h is not None:
                    row.append(f"{h:.1f}")
                    total += h
                    has_any = True
                else:
                    row.append("")
            row.append(f"{total:.1f}" if total > 0 else "")
            if has_any:
                tree.insert("", "end", values=row)

    # ------------------------------------------------------------------
    # Проверка и исправление проблемных сессий
    # ------------------------------------------------------------------
    def check_sessions(self):
        try:
            rows = get_problematic_sessions(self.cursor)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось получить список:\n{e}")
            return

        tree = self.tree_check
        tree.delete(*tree.get_children())
        columns = ("rab", "fio", "ip", "place", "start", "end", "hours")
        tree["columns"] = columns
        tree.heading("rab", text="RAB")
        tree.heading("fio", text="ФИО")
        tree.heading("ip", text="IP")
        tree.heading("place", text="Рабочее место")
        tree.heading("start", text="Начало")
        tree.heading("end", text="Конец (неверный)")
        tree.heading("hours", text="Часов")
        tree.column("rab", width=90, anchor="center")
        tree.column("fio", width=200, anchor="w")
        tree.column("ip", width=60, anchor="center")
        tree.column("place", width=180, anchor="w")
        tree.column("start", width=140, anchor="center")
        tree.column("end", width=140, anchor="center")
        tree.column("hours", width=70, anchor="center")

        count = 0
        for rab, fio, ip, place, ts, te, minutes in rows:
            hours = round(float(minutes) / 60.0, 1) if minutes else 0.0
            tree.insert("", "end", values=(
                rab, fio, ip, place or "-",
                ts.strftime("%Y-%m-%d %H:%M") if ts else "",
                te.strftime("%Y-%m-%d %H:%M") if te else "",
                f"{hours:.1f}"
            ))
            count += 1

        self.check_info.config(
            text=f"Найдено проблемных сессий: {count}. "
                 f"(Закрытие позже 18:00 дня открытия, включая дни, которые уже прошли.)"
        )

    def fix_sessions(self):
        rows = get_problematic_sessions(self.cursor)
        # Оставляем только те, что начались раньше сегодняшнего дня
        today = date.today()
        fixable = [r for r in rows if r[4] and r[4].date() < today]
        if not fixable:
            messagebox.showinfo("Нечего исправлять",
                                "Нет проблемных сессий, начатых до сегодняшнего дня.")
            return

        if not messagebox.askyesno(
            "Подтверждение",
            f"Будет исправлено {len(fixable)} сессий.\n"
            f"Время окончания будет установлено в 18:00 дня открытия.\n"
            f"Продолжить?"
        ):
            return

        try:
            affected = fix_problematic_sessions(self.cursor, self.conn)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось исправить:\n{e}")
            return

        messagebox.showinfo("Готово",
                            f"Исправлено сессий: {affected}.\n"
                            f"Обновите вкладку «Табель», чтобы увидеть результат.")
        # Обновляем список проблемных
        self.check_sessions()
        # Обновляем табель
        self.refresh_timesheet()

    # ------------------------------------------------------------------
    def on_orders_search_change(self, *args):
        if self._orders_search_job is not None:
            try:
                self.root.after_cancel(self._orders_search_job)
            except Exception:
                pass
        self._orders_search_job = self.root.after(500, self.refresh_orders)

    def clear_orders_search(self):
        self.orders_search_var.set("")
        self.refresh_orders()

    # ------------------------------------------------------------------
    def export_current_tab(self):
        current = self.notebook.index(self.notebook.select())
        if current == 0:
            name, tree = "workers", self.tree_workers
        elif current == 1:
            name, tree = "detals", self.tree_detals
        elif current == 2:
            name, tree = "orders", self.tree_orders
        elif current == 3:
            name, tree = "taxel", self.tree_taxel
        else:
            name, tree = "check", self.tree_check

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"analysis_{name}_{self.date_var.get()}.csv"
        )
        if not filepath:
            return
        try:
            with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                headers = [tree.heading(c)["text"].replace("\n", " ") for c in tree["columns"]]
                writer.writerow(headers)
                for item in tree.get_children():
                    writer.writerow(tree.item(item)["values"])
            messagebox.showinfo("Готово", f"Данные сохранены в:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

    # ------------------------------------------------------------------
    def auto_refresh(self):
        try:
            self.refresh_all()
        except Exception as e:
            self.status.config(text=f"Ошибка: {e}")
        self.root.after(30000, self.auto_refresh)

    def on_close(self):
        try:
            if self.conn:
                self.conn.close()
        finally:
            self.root.destroy()


# ----------------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = AnalyzerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()
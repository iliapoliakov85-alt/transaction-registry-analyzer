"""
Анализатор реестра транзакций v8.1 (PySide6)
Исправлено: календарь (белый фон, читаемые даты), дефолтный диапазон = реальные min/max из данных.
"""
from __future__ import annotations
import sys
import os
import time
import json
import urllib.request
import io
import shutil
import tempfile
import zipfile
import codecs
import struct
import zlib
import hashlib
import subprocess
import pandas as pd
from collections import OrderedDict
from typing import Optional, List, Dict, Any, Tuple
from datetime import datetime, date

try:
    import duckdb
except ImportError:
    print("❌ pip install duckdb"); sys.exit(1)
try:
    import openpyxl
except ImportError:
    print("❌ pip install openpyxl"); sys.exit(1)

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QToolBar, QSplitter,
    QSizePolicy, QStyle, QProxyStyle, QTableView, QStackedWidget,
    QAbstractItemView, QStyledItemDelegate,
    QListWidget, QListWidgetItem, QCheckBox, QScrollArea,
    QFrame, QProgressBar, QFileDialog, QMessageBox, QStatusBar, QToolTip,
    QDialog, QDialogButtonBox, QGraphicsDropShadowEffect, QGraphicsOpacityEffect,
    QHeaderView, QDateEdit, QCalendarWidget, QSlider, QTableWidget, QTableWidgetItem, QMenu,
    QStyleOptionViewItem
)
from PySide6.QtCore import (
    Qt, QSize, QRect, QPropertyAnimation, QEasingCurve, QEvent, QLocale,
    QObject, Signal, QThread, QSettings, QTimer, QModelIndex, QUrl, QDir,
    QAbstractTableModel, QDate
)
from PySide6.QtGui import (
    QFont, QColor, QPalette, QIcon, QPixmap, QPainter, QAction,
    QKeySequence, QPen, QFontDatabase, QTextCharFormat, QDesktopServices
)
from PySide6.QtSvg import QSvgRenderer

DB_FILE = "analysis_db.duckdb"
APP_VERSION = "8.4.1"
# Адрес манифеста обновлений берётся из переменной окружения либо из
# update_config.json рядом с программой. Манифест можно хранить в GitHub.
UPDATE_MANIFEST_URL = os.getenv("REGISTRY_ANALYZER_UPDATE_URL", "")

def get_update_manifest_url():
    if UPDATE_MANIFEST_URL:
        return UPDATE_MANIFEST_URL
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "update_config.json")
    try:
        with open(config_path, "r", encoding="utf-8") as config_file:
            return str(json.load(config_file).get("manifest_url", "")).strip()
    except (OSError, ValueError, TypeError):
        return ""

# ─── ИКОНКИ (SVG) ─────────────────────────────────────────────
ICON_OPEN = """<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20"><path fill="CURRENT" d="M3 5a2 2 0 012-2h3.172a2 2 0 011.414.586l1.828 1.828A2 2 0 0012.828 6H15a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2V5z"/></svg>"""
ICON_EXPORT = """<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20"><path fill="CURRENT" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm3.293-7.707a1 1 0 011.414 0L9 10.586V3a1 1 0 112 0v7.586l1.293-1.293a1 1 0 111.414 1.414l-3 3a1 1 0 01-1.414 0l-3-3a1 1 0 010-1.414z"/></svg>"""
ICON_APPLY = """<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20"><path fill="CURRENT" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z"/></svg>"""
ICON_SEARCH = """<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><path fill="CURRENT" d="M11.742 10.344a6.5 6.5 0 10-1.397 1.398h-.001l3.85 3.85a1 1 0 001.415-1.414l-3.85-3.85a1 1 0 00-.115-.1zM12 6.5a5.5 5.5 0 11-11 0 5.5 5.5 0 0111 0z"/></svg>"""
ICON_SIDEBAR = """<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20"><path fill="CURRENT" d="M3 4a1 1 0 011-1h12a1 1 0 011 1v12a1 1 0 01-1 1H4a1 1 0 01-1-1V4zm1 0v12h4V4H4z"/></svg>"""

# ─── КОНСТАНТЫ ────────────────────────────────────────────────
CARRIER_COLUMN = "Идентификатор перевозчика"
SUM_COLUMNS = {
    "Применённый тариф": "Применённый тариф",
    "Регулируемый тариф": "Регулируемый тариф",
    "Льгота": "Льгота",
    "Оплаченная сумма": "Оплаченная сумма",
}
DATETIME_COLUMN = "Дата и время совершения поездки"

DATE_COLUMNS = [
    "Дата и время совершения поездки",
    "Дата и время обработки на сервере"
]

# CSS для календаря — принудительно белый фон, тёмный текст
DATE_EDIT_STYLE = """
    QDateEdit {
        border: 1px solid #D1D5DB;
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 12px;
        min-width: 95px;
        background-color: #FFFFFF;
        color: #111827;
    }
    QDateEdit:focus {
        border: 1px solid #3B82F6;
    }
    QDateEdit::drop-down {
        border: none;
        width: 22px;
        background-color: #F3F4F6;
        border-top-right-radius: 4px;
        border-bottom-right-radius: 4px;
    }
    QDateEdit::down-arrow {
        image: none;
        border-left: 4px solid transparent;
        border-right: 4px solid transparent;
        border-top: 5px solid #6B7280;
        width: 0;
        height: 0;
    }
    QCalendarWidget {
        background-color: #FFFFFF;
        color: #111827;
        border: 1px solid #D1D5DB;
    }
    QCalendarWidget QWidget#qt_calendar_navigationbar {
        background-color: #F9FAFB;
        border-bottom: 1px solid #E5E7EB;
        padding: 4px;
    }
    QCalendarWidget QToolButton {
        color: #111827;
        background-color: transparent;
        border: none;
        padding: 4px 8px;
        font-size: 12px;
        font-weight: bold;
    }
    QCalendarWidget QToolButton:hover {
        background-color: #E5E7EB;
        border-radius: 4px;
    }
    QCalendarWidget QToolButton#qt_calendar_prevmonth,
    QCalendarWidget QToolButton#qt_calendar_nextmonth {
        width: 24px;
        height: 24px;
    }
    QCalendarWidget QSpinBox {
        background-color: #FFFFFF;
        color: #111827;
        border: 1px solid #D1D5DB;
        border-radius: 3px;
        padding: 2px;
    }
    QCalendarWidget QAbstractItemView {
        background-color: #FFFFFF;
        color: #111827;
        selection-background-color: #3B82F6;
        selection-color: #FFFFFF;
        font-size: 12px;
        outline: none;
    }
    QCalendarWidget QAbstractItemView:enabled {
        background-color: #FFFFFF;
        color: #111827;
    }
    QCalendarWidget QAbstractItemView:disabled {
        color: #D1D5DB;
    }
"""

CALENDAR_POPUP_STYLE = """
QCalendarWidget { background:#FFFFFF; color:#172033; border:1px solid #CBD5E1; }
QCalendarWidget QWidget#qt_calendar_navigationbar { background:#F8FAFC; border-bottom:1px solid #E2E8F0; }
QCalendarWidget QToolButton { color:#172033; background:transparent; border:none; font-weight:700; }
QCalendarWidget QToolButton:hover { background:#EAF2FF; border-radius:6px; }
QCalendarWidget QSpinBox { color:#172033; background:#FFFFFF; border:1px solid #CBD5E1; border-radius:5px; }
QCalendarWidget QAbstractItemView { color:#172033; background:#FFFFFF; selection-background-color:#1769E0; selection-color:#FFFFFF; alternate-background-color:#FFFFFF; outline:0; }
QCalendarWidget QAbstractItemView::item { color:#172033; background:#FFFFFF; }
QCalendarWidget QAbstractItemView::item:hover { background:#EAF2FF; color:#172033; }
QCalendarWidget QAbstractItemView::item:selected { background:#1769E0; color:#FFFFFF; }
QCalendarWidget QTableView { color:#172033; background:#FFFFFF; alternate-background-color:#FFFFFF; selection-background-color:#1769E0; selection-color:#FFFFFF; }
QCalendarWidget QTableView::item { color:#172033; background:#FFFFFF; }
QCalendarWidget QTableView::item:!selected { color:#172033; background:#FFFFFF; }
QCalendarWidget QTableView::item:hover { color:#172033; background:#EAF2FF; }
QCalendarWidget QTableView::item:selected { color:#FFFFFF; background:#1769E0; }
QCalendarWidget QWidget#qt_calendar_calendarview { color:#172033; background:#FFFFFF; }
QCalendarWidget QWidget#qt_calendar_calendarview::item { color:#172033; background:#FFFFFF; }
QCalendarWidget QWidget#qt_calendar_calendarview::item:selected { color:#FFFFFF; background:#1769E0; }
QCalendarWidget QWidget#qt_calendar_calendarview::item:disabled { color:#94A3B8; background:#FFFFFF; }
"""

class ReadableCalendarDelegate(QStyledItemDelegate):
    """Принудительно рисует номера дней: системная тема macOS их не переопределит."""
    def paint(self, painter, option, index):
        option = QStyleOptionViewItem(option)
        selected = bool(option.state & QStyle.State_Selected)
        foreground = QColor("#FFFFFF" if selected else "#172033")
        option.palette.setColor(QPalette.Text, foreground)
        option.palette.setColor(QPalette.WindowText, foreground)
        option.palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
        option.palette.setColor(QPalette.Base, QColor("#FFFFFF"))
        super().paint(painter, option, index)
        # Некоторые версии Qt/macOS заново применяют белый цвет именно к
        # части дат после 9-го числа. Дублируем цифру поверх системной
        # отрисовки, поэтому она остаётся читаемой при любой теме.
        value = index.data(Qt.DisplayRole)
        if value not in (None, ""):
            painter.save()
            painter.setPen(QColor("#FFFFFF" if selected else "#172033"))
            painter.drawText(option.rect, Qt.AlignCenter, str(value))
            painter.restore()

def setup_russian_calendar(date_edit):
    """Единый читаемый русский календарь для всех дат в приложении."""
    locale = QLocale(QLocale.Russian, QLocale.Russia)
    date_edit.setLocale(locale)
    calendar = date_edit.calendarWidget()
    calendar.setLocale(locale); calendar.setFirstDayOfWeek(Qt.Monday); calendar.setGridVisible(False)
    calendar.setStyleSheet(CALENDAR_POPUP_STYLE)
    # В Qt/macOS цифры месяца выводятся внутренним QTableView; делегат
    # рисует их явно и не зависит от палитры активной темы приложения.
    for view in calendar.findChildren(QTableView):
        view.setItemDelegate(ReadableCalendarDelegate(view))
        view.setStyleSheet("background:#FFFFFF; color:#172033; selection-background-color:#1769E0; selection-color:#FFFFFF;")

class Theme:
    class Colors:
        blue = "#3B82F6"; green = "#10B981"; purple = "#8B5CF6"
        orange = "#F59E0B"; gray = "#9CA3AF"; bg = "#F9FAFB"
        card_bg = "#FFFFFF"; border = "#E5E7EB"
        text_primary = "#111827"; text_secondary = "#6B7280"
    class Spacing:
        xs = 4; s = 8; m = 12; l = 16; xl = 24; xxl = 32
    class Typography:
        class Size:
            caption = 11; body = 13; title = 14; headline = 22; large = 28
        class Weight:
            regular = QFont.Normal; medium = QFont.Medium; semibold = QFont.DemiBold; bold = QFont.Bold
    class Animation:
        fast = 120; normal = 150
    class Toolbar:
        icon_size = 18
    class Sidebar:
        min_width = 280; max_width = 450; normal_width = 320
    class Table:
        row_height = 36; header_height = 28; MAX_DISPLAY_ROWS = 5000
    class Chart:
        default_height = 170

class DuckDBRepository:
    def __init__(self):
        self._conn = None; self._columns = []; self._all_columns = []
        self._table_name = "raw_data"; self._db_path = None
        self._id_column = "row_id"; self._sum_columns = {}
        self._date_ranges = {}

    def _table_exists(self):
        if not self._conn: return False
        try:
            return self._conn.execute("SELECT COUNT(*) FROM information_schema.tables WHERE table_name='raw_data'").fetchone()[0] > 0
        except:
            return False

    def _timestamp_expr(self, column):
        """Поддерживает и ISO-даты, и распространённый формат CSV дд.мм.гггг."""
        col = f'CAST("{column}" AS VARCHAR)'
        return (
            f"COALESCE(TRY_CAST({col} AS TIMESTAMP), "
            f"TRY_STRPTIME({col}, '%d.%m.%Y %H:%M:%S'), "
            f"TRY_STRPTIME({col}, '%d.%m.%Y'))"
        )

    def _detect_columns(self):
        if not self._conn or not self._table_exists(): 
            self._columns = []; self._all_columns = []; self._date_ranges = {}
            return
        try:
            result = self._conn.execute(f"SELECT * FROM {self._table_name} LIMIT 0")
            col_names = [d[0] for d in result.description]
            # `_columns` deliberately hides technical keys from the main UI.
            # Reconciliation, however, must be able to find RRN among all
            # physical columns of the imported registry.
            self._all_columns = col_names
            self._id_column = "row_id" if "row_id" in col_names else "RRN"
            self._columns = [c for c in col_names if c not in ("row_id", "RRN")]
            
            self._sum_columns = {}
            for alias, exact_name in SUM_COLUMNS.items():
                if exact_name in self._columns:
                    self._sum_columns[alias] = exact_name
                else:
                    found = None
                    keywords = {
                        "Оплаченная сумма": ["оплаченн", "оплач", "paid", "сумма оплат"],
                        "Применённый тариф": ["тариф", "применённ", "примененн", "tariff"],
                        "Регулируемый тариф": ["регулируем", "регулир"],
                        "Льгота": ["льгот", "льг", "скидк", "скид", "benefit", "compensation"],
                    }
                    for kw in keywords.get(alias, []):
                        for col in self._columns:
                            if kw in col.lower():
                                found = col; break
                        if found: break
                    if found:
                        self._sum_columns[alias] = found

            # Определяем реальные min/max даты из данных
            self._date_ranges = {}
            for col in DATE_COLUMNS:
                if col in self._columns:
                    try:
                        timestamp_expr = self._timestamp_expr(col)
                        r = self._conn.execute(f"""
                            SELECT 
                                MIN({timestamp_expr}),
                                MAX({timestamp_expr})
                            FROM {self._table_name}
                            WHERE "{col}" IS NOT NULL
                              AND {timestamp_expr} IS NOT NULL
                        """).fetchone()
                        if r and r[0] and r[1]:
                            min_dt = r[0]
                            max_dt = r[1]
                            # Конвертируем в QDate
                            if hasattr(min_dt, 'date'):
                                min_qd = QDate(min_dt.year, min_dt.month, min_dt.day)
                            else:
                                min_qd = QDate.currentDate().addYears(-1)
                            if hasattr(max_dt, 'date'):
                                max_qd = QDate(max_dt.year, max_dt.month, max_dt.day)
                            else:
                                max_qd = QDate.currentDate()
                            self._date_ranges[col] = (min_qd, max_qd)
                            print(f"[DEBUG] {col}: {min_qd.toString('dd.MM.yyyy')} — {max_qd.toString('dd.MM.yyyy')}")
                    except Exception as e:
                        print(f"[DEBUG] Date range error for {col}: {e}")
        except Exception as e:
            print(f"Column detection error: {e}")

    def get_date_range(self, column):
        return self._date_ranges.get(column, None)

    def connect(self, db_path):
        if self._conn:
            try: self._conn.close()
            except: pass
        if not os.path.exists(db_path): raise FileNotFoundError(f"DB not found: {db_path}")
        self._conn = duckdb.connect(db_path)
        self._conn.execute("PRAGMA memory_limit='2GB'"); self._conn.execute("PRAGMA threads=4")
        self._db_path = db_path; self._detect_columns()

    def create_from_csv(self, db_path, csv_path, progress_callback=None):
        if self._conn:
            try: self._conn.close()
            except: pass
        if os.path.exists(db_path):
            try: os.remove(db_path)
            except PermissionError: raise PermissionError(f"Файл БД '{db_path}' занят другой программой.")
        
        self._conn = duckdb.connect(db_path)
        self._conn.execute("PRAGMA memory_limit='2GB'"); self._conn.execute("PRAGMA threads=4")
        self._db_path = db_path

        if progress_callback: progress_callback(5, "Анализ CSV...")
        if not os.path.exists(csv_path): raise FileNotFoundError(f"CSV not found: {csv_path}")
        
        safe_path = csv_path.replace("'", "''")
        
        if progress_callback: progress_callback(10, "Чтение заголовков...")
        col_q = f"SELECT * FROM read_csv('{safe_path}', sep=';', header=True, skip=1) LIMIT 0"
        res = self._conn.execute(col_q)
        all_cols = [d[0] for d in res.description]
        type_overrides = {c: "VARCHAR" for c in all_cols}
        
        if progress_callback: progress_callback(30, "Загрузка данных...")
        self._conn.execute(f"DROP TABLE IF EXISTS {self._table_name}")
        self._conn.execute(f"""
            CREATE TABLE {self._table_name} AS
            SELECT ROW_NUMBER() OVER () as row_id, *
            FROM read_csv(
                '{safe_path}', sep=';', header=True, skip=1,
                ignore_errors=True,
                columns={type_overrides},
                nullstr=['', '-', 'NULL', 'null'],
                strict_mode=False, quote=''
            )
        """)
        
        self._id_column = "row_id"
        if progress_callback: progress_callback(60, "Индексы...")
        self._conn.execute(f"CREATE INDEX idx_row_id ON {self._table_name}(row_id)")
        if progress_callback: progress_callback(75, "ANALYZE...")
        self._conn.execute(f"ANALYZE {self._table_name}")
        if progress_callback: progress_callback(85, "Колонки...")
        self._detect_columns()
        if progress_callback: progress_callback(90, "Подсчёт...")
        row_count = self._conn.execute(f"SELECT COUNT(*) FROM {self._table_name}").fetchone()[0]
        if progress_callback: progress_callback(95, "Оптимизация...")
        self._conn.execute("CHECKPOINT")
        if progress_callback: progress_callback(100, "Готово")
        return row_count

    def get_columns(self): return self._columns
    def get_row_count(self, where_clause="1=1"):
        if not self._conn or not self._table_exists(): return 0
        try: return self._conn.execute(f"SELECT COUNT(*) FROM {self._table_name} WHERE {where_clause}").fetchone()[0]
        except: return 0

    def get_aggregated(self, where_clause="1=1", limit=5000):
        if not self._conn or not self._table_exists(): return None
        sum_exprs = []
        for alias, db_col in self._sum_columns.items():
            if db_col:
                clean_expr = f"REGEXP_REPLACE(REPLACE(REPLACE(\"{db_col}\", ' ', ''), ',', '.'), '[^0-9.-]', '', 'g')"
                sum_exprs.append(f'SUM(COALESCE(TRY_CAST({clean_expr} AS DOUBLE), 0)) as "{alias}"')
            else:
                sum_exprs.append(f'0.0 as "{alias}"')
        sum_exprs_str = ", ".join(sum_exprs)
        
        query = f"""
            SELECT COALESCE("{CARRIER_COLUMN}", '(Пусто)') as "{CARRIER_COLUMN}",
                   COUNT(*) as "Количество",
                   {sum_exprs_str}
            FROM {self._table_name}
            WHERE {where_clause}
            GROUP BY "{CARRIER_COLUMN}"
            ORDER BY "Оплаченная сумма" DESC
            LIMIT {limit}
        """
        try: return self._conn.execute(query).df()
        except Exception as e:
            print(f"Aggregation error: {e}"); return None

    def get_daily_paid_totals(self, where_clause="1=1"):
        """Данные только для графика динамики; основной расчёт и сводка не меняются."""
        if not self._conn or DATETIME_COLUMN not in self._columns:
            return []
        paid_column = self._sum_columns.get("Оплаченная сумма")
        if not paid_column:
            return []
        date_expr = self._timestamp_expr(DATETIME_COLUMN)
        clean_expr = f"REGEXP_REPLACE(REPLACE(REPLACE(\"{paid_column}\", ' ', ''), ',', '.'), '[^0-9.-]', '', 'g')"
        try:
            rows = self._conn.execute(f"""
                SELECT CAST({date_expr} AS DATE) AS day,
                       SUM(COALESCE(TRY_CAST({clean_expr} AS DOUBLE), 0)) AS total
                FROM {self._table_name}
                WHERE {where_clause} AND {date_expr} IS NOT NULL
                GROUP BY day ORDER BY day
            """).fetchall()
            return [(str(day), float(total or 0)) for day, total in rows]
        except Exception as e:
            print(f"Chart error: {e}")
            return []

    def get_fleet_carriers(self, where_clause="1=1"):
        """Список перевозчиков для отдельного окна аналитики парка."""
        if not self._conn or CARRIER_COLUMN not in self._columns:
            return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        try:
            return [(str(name), int(count)) for name, count in self._conn.execute(f"""
                SELECT {carrier} AS carrier, COUNT(*) AS total
                FROM {self._table_name} WHERE {where_clause}
                GROUP BY carrier ORDER BY total DESC
            """).fetchall()]
        except Exception as e:
            print(f"Fleet carriers error: {e}"); return []

    def get_fleet_summary(self, where_clause="1=1", carriers=None):
        """Сводка ТС, маршрутов и транзакций без изменения исходных данных."""
        required = {CARRIER_COLUMN, "Транспортное средство", "Маршрут", "Тип ТС"}
        if not self._conn or not required.issubset(self._columns):
            return {}
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), '')"
        route = "NULLIF(TRIM(CAST(\"Маршрут\" AS VARCHAR)), '')"
        type_expr = "COALESCE(NULLIF(TRIM(CAST(\"Тип ТС\" AS VARCHAR)), ''), 'Не указан')"
        driver = "NULLIF(TRIM(CAST(\"ФИО водителя\" AS VARCHAR)), '')" if "ФИО водителя" in self._columns else None
        extra = ""
        if carriers:
            escaped = "', '".join(str(value).replace("'", "''") for value in carriers)
            extra = f" AND {carrier} IN ('{escaped}')"
        try:
            totals = self._conn.execute(f"""
                SELECT {carrier} AS carrier, COUNT(DISTINCT {vehicle}) AS vehicles,
                       COUNT(*) AS transactions, COUNT(DISTINCT {route}) AS routes,
                       {f'COUNT(DISTINCT {driver})' if driver else '0'} AS drivers
                FROM {self._table_name} WHERE {where_clause}{extra}
                GROUP BY carrier
            """).fetchall()
            types = self._conn.execute(f"""
                SELECT {carrier} AS carrier, {type_expr} AS vehicle_type,
                       COUNT(DISTINCT {vehicle}) AS vehicles, COUNT(*) AS transactions
                FROM {self._table_name} WHERE {where_clause}{extra}
                GROUP BY carrier, vehicle_type
            """).fetchall()
            result = {str(name): {"vehicles": int(v or 0), "transactions": int(t or 0), "routes": int(r or 0), "drivers": int(d or 0), "work_hours": 0.0, "types": []}
                      for name, v, t, r, d in totals}
            for name, vehicle_type, vehicles, transactions in types:
                result[str(name)]["types"].append({"type": str(vehicle_type), "vehicles": int(vehicles or 0), "transactions": int(transactions or 0)})
            for item in result.values():
                item["vehicles_by_type"] = sum(row["vehicles"] for row in item["types"])
            if DATETIME_COLUMN in self._columns:
                timestamp = self._timestamp_expr(DATETIME_COLUMN)
                work_rows = self._conn.execute(f"""
                    SELECT carrier, SUM(work_hours) AS work_hours FROM (
                        SELECT {carrier} AS carrier, {vehicle} AS vehicle,
                               CAST({timestamp} AS DATE) AS day,
                               EXTRACT(EPOCH FROM (MAX({timestamp}) - MIN({timestamp}))) / 3600.0 AS work_hours
                        FROM {self._table_name}
                        WHERE {where_clause}{extra} AND {vehicle} IS NOT NULL AND {timestamp} IS NOT NULL
                        GROUP BY carrier, vehicle, day
                    ) daily_work
                    GROUP BY carrier
                """).fetchall()
                for name, hours in work_rows:
                    if str(name) in result: result[str(name)]["work_hours"] = float(hours or 0)
            return result
        except Exception as e:
            print(f"Fleet summary error: {e}"); return {}

    def get_fleet_vehicles(self, where_clause, carrier_name):
        if not self._conn or "Транспортное средство" not in self._columns:
            return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), '')"
        escaped = str(carrier_name).replace("'", "''")
        try:
            return [(str(vehicle), int(count)) for vehicle, count in self._conn.execute(f"""
                SELECT {vehicle} AS vehicle, COUNT(*) AS transactions
                FROM {self._table_name}
                WHERE {where_clause} AND {carrier} = '{escaped}' AND {vehicle} IS NOT NULL
                GROUP BY vehicle ORDER BY transactions DESC
            """).fetchall()]
        except Exception as e:
            print(f"Fleet vehicles error: {e}"); return []

    def get_duplicate_fleet_vehicles(self, where_clause, carrier_name):
        """Номера ТС, которые попали более чем в один тип ТС.

        Это объясняет расхождение «уникальных ТС» и суммы ТС по типам, не
        меняя ни одной исходной записи.
        """
        required = {CARRIER_COLUMN, "Транспортное средство", "Тип ТС"}
        if not self._conn or not required.issubset(self._columns): return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), '')"
        vehicle_type = "COALESCE(NULLIF(TRIM(CAST(\"Тип ТС\" AS VARCHAR)), ''), 'Не указан')"
        escaped = str(carrier_name).replace("'", "''")
        try:
            rows = self._conn.execute(f"""
                SELECT {vehicle} AS vehicle,
                       STRING_AGG(DISTINCT {vehicle_type}, ', ' ORDER BY {vehicle_type}) AS types,
                       COUNT(DISTINCT {vehicle_type}) AS type_count,
                       COUNT(*) AS transactions
                FROM {self._table_name}
                WHERE {where_clause} AND {carrier} = '{escaped}' AND {vehicle} IS NOT NULL
                GROUP BY vehicle
                HAVING COUNT(DISTINCT {vehicle_type}) > 1
                ORDER BY type_count DESC, transactions DESC, vehicle
            """).fetchall()
            return [(str(vehicle), str(types), int(type_count or 0), int(transactions or 0))
                    for vehicle, types, type_count, transactions in rows]
        except Exception as e:
            print(f"Duplicate fleet vehicles error: {e}")
            return []

    def get_fraud_cards(self, where_clause="1=1"):
        """Маски обслуженных карт только для безналичной формы оплаты."""
        card_column = "Номер обслуженной карты"
        payment_column = "Форма оплаты проезда"
        if not self._conn or card_column not in self._columns or payment_column not in self._columns:
            return []
        paid_column = self._sum_columns.get("Оплаченная сумма")
        card = f"NULLIF(TRIM(CAST(\"{card_column}\" AS VARCHAR)), '')"
        payment = f"LOWER(TRIM(CAST(\"{payment_column}\" AS VARCHAR)))"
        paid = "0"
        if paid_column:
            clean = f"REGEXP_REPLACE(REPLACE(REPLACE(\"{paid_column}\", ' ', ''), ',', '.'), '[^0-9.-]', '', 'g')"
            paid = f"COALESCE(TRY_CAST({clean} AS DOUBLE), 0)"
        try:
            rows = self._conn.execute(f"""
                SELECT {card} AS card_mask, COUNT(*) AS transactions, SUM({paid}) AS amount
                FROM {self._table_name}
                WHERE {where_clause} AND {card} IS NOT NULL AND {payment} LIKE '%безналич%'
                GROUP BY card_mask
                ORDER BY transactions DESC, amount DESC, card_mask
            """).fetchall()
            return [{"mask": str(mask), "transactions": int(count or 0), "amount": float(amount or 0)}
                    for mask, count, amount in rows]
        except Exception as e:
            print(f"Fraud cards error: {e}")
            return []

    def get_fraud_card_rows(self, where_clause, card_mask):
        """Исходные строки одной маски для проверки расчёта по двойному клику."""
        card_column = "Номер обслуженной карты"
        payment_column = "Форма оплаты проезда"
        if not self._conn or card_column not in self._columns or payment_column not in self._columns:
            return pd.DataFrame()
        card = f"NULLIF(TRIM(CAST(\"{card_column}\" AS VARCHAR)), '')"
        payment = f"LOWER(TRIM(CAST(\"{payment_column}\" AS VARCHAR)))"
        escaped = str(card_mask).replace("'", "''")
        try:
            df = self._conn.execute(f"""
                SELECT * FROM {self._table_name}
                WHERE {where_clause} AND {card} = '{escaped}' AND {payment} LIKE '%безналич%'
                ORDER BY {self._id_column}
            """).df()
            return df.drop("row_id", axis=1) if "row_id" in df.columns else df
        except Exception as e:
            print(f"Fraud card rows error: {e}")
            return pd.DataFrame()

    def _find_column(self, *candidates):
        """Поиск колонки без зависимости от регистра и лишних пробелов."""
        def normalize(value): return "".join(ch for ch in str(value).casefold() if ch.isalnum())
        wanted = {normalize(item) for item in candidates}
        return next((column for column in self._all_columns if normalize(column) in wanted), None)

    def investor_source_columns(self):
        return {
            "rrn": self._find_column("RRN"),
            "transaction_id": self._find_column("ID транзакции", "transaction_id"),
            "realizator": self._find_column("Реализатор"),
            "payment": self._find_column("Форма оплаты проезда"),
            "category": self._find_column("Категория пассажира"),
            "carrier": self._find_column(CARRIER_COLUMN, "Перевозчик"),
            "trip_datetime": self._find_column(DATETIME_COLUMN, "Дата поездки"),
            "server_datetime": self._find_column("Дата и время обработки на сервере", "Дата обработки на сервере"),
            "status": self._find_column("Статус"),
        }

    def _reconciliation_period_conditions(self, source, asop_dates=None, vtb_dates=None):
        """Возвращает SQL-ограничения независимых периодов АСОП и ВТБ."""
        def range_clause(expr, date_range):
            if not date_range: return ""
            start, end = date_range
            try:
                if start: datetime.strptime(start, "%Y-%m-%d")
                if end: datetime.strptime(end, "%Y-%m-%d")
            except (TypeError, ValueError):
                return ""
            parts = []
            if start: parts.append(f"{expr} >= DATE '{start}'")
            if end: parts.append(f"{expr} <= DATE '{end}'")
            return "AND " + " AND ".join(parts) if parts else ""
        asop_expr = f"CAST({self._timestamp_expr(source['trip_datetime'])} AS DATE)" if source.get("trip_datetime") else "NULL"
        vtb_expr = "COALESCE(TRY_CAST(CAST(\"Дата оплаты на терминале\" AS VARCHAR) AS DATE), CAST(TRY_STRPTIME(CAST(\"Дата оплаты на терминале\" AS VARCHAR), '%d.%m.%Y') AS DATE))"
        return range_clause(asop_expr, asop_dates), range_clause(vtb_expr, vtb_dates)

    def get_investor_reconciliation(self, investor_db_path, asop_dates=None, vtb_dates=None):
        """Сверка АСОП/безналичных операций строго по ID транзакции."""
        source = self.investor_source_columns()
        required = ("transaction_id", "realizator", "payment", "category", "carrier")
        if not self._conn or not all(source.get(key) for key in required) or not os.path.exists(investor_db_path):
            return []
        paid_col = self._sum_columns.get("Оплаченная сумма")
        if not paid_col: return []
        escaped_path = investor_db_path.replace("'", "''")
        paid_clean = f"REGEXP_REPLACE(REPLACE(REPLACE(\"{paid_col}\", ' ', ''), ',', '.'), '[^0-9.-]', '', 'g')"
        paid = f"COALESCE(TRY_CAST({paid_clean} AS DOUBLE), 0)"
        asop_period, vtb_period = self._reconciliation_period_conditions(source, asop_dates, vtb_dates)
        try:
            try: self._conn.execute("DETACH investor")
            except: pass
            self._conn.execute(f"ATTACH '{escaped_path}' AS investor (READ_ONLY)")
            query = f"""
                WITH main_source AS (
                    SELECT COALESCE(NULLIF(TRIM(CAST(\"{source['carrier']}\" AS VARCHAR)), ''), 'Не указан') AS carrier,
                           NULLIF(TRIM(CAST(\"{source['transaction_id']}\" AS VARCHAR)), '') AS transaction_id,
                           {paid} AS amount
                    FROM {self._table_name}
                    WHERE TRIM(COALESCE(CAST(\"{source['realizator']}\" AS VARCHAR), '')) = ''
                      AND LOWER(TRIM(COALESCE(CAST(\"{source['payment']}\" AS VARCHAR), ''))) LIKE '%безналич%'
                      AND LOWER(TRIM(COALESCE(CAST(\"{source['category']}\" AS VARCHAR), ''))) = 'без льготы'
                      {asop_period}
                ), main_data AS (
                    SELECT carrier, transaction_id, COUNT(*) AS main_count, SUM(amount) AS main_amount
                    FROM main_source WHERE transaction_id IS NOT NULL
                    GROUP BY carrier, transaction_id
                ), investor_data AS (
                    SELECT NULLIF(TRIM(CAST(\"Организация\" AS VARCHAR)), '') AS organization,
                           NULLIF(TRIM(CAST(\"transaction_id\" AS VARCHAR)), '') AS transaction_id,
                           COUNT(*) AS investor_count,
                           SUM(COALESCE(TRY_CAST(REPLACE(TRIM(CAST(\"Сумма операции, руб.\" AS VARCHAR)), ',', '.') AS DOUBLE), 0)) AS investor_amount
                    FROM investor.investor_data
                    WHERE NULLIF(TRIM(CAST(\"transaction_id\" AS VARCHAR)), '') IS NOT NULL
                      {vtb_period}
                    GROUP BY organization, transaction_id
                ), comparison AS (
                    SELECT CASE WHEN m.transaction_id IS NULL THEN 'Нет в АСОП' ELSE m.carrier END AS carrier,
                           CASE WHEN i.transaction_id IS NULL THEN 'Нет в отчёте ВТБ'
                                ELSE COALESCE(i.organization, 'Организация не указана') END AS organization,
                           m.transaction_id AS main_transaction_id, i.transaction_id AS investor_transaction_id,
                           COALESCE(m.main_count, 0) AS main_count, COALESCE(i.investor_count, 0) AS investor_count,
                           COALESCE(m.main_amount, 0) AS main_amount, COALESCE(i.investor_amount, 0) AS investor_amount,
                           CASE WHEN m.transaction_id IS NULL THEN 'Только у инвестора'
                                WHEN i.transaction_id IS NULL THEN 'Только в АСОП'
                                WHEN ABS(COALESCE(m.main_amount, 0) - COALESCE(i.investor_amount, 0)) > 0.01 THEN 'Расхождение суммы'
                                ELSE 'Совпало' END AS status
                    FROM main_data m FULL OUTER JOIN investor_data i
                      ON m.transaction_id = i.transaction_id
                ), metrics AS (
                    SELECT
                        SUM(CASE WHEN main_transaction_id IS NOT NULL AND investor_transaction_id IS NOT NULL THEN 1 ELSE 0 END) AS matched_ids,
                        SUM(CASE WHEN main_count > 1 THEN 1 ELSE 0 END) AS repeated_asop_ids,
                        SUM(CASE WHEN investor_count > 1 THEN 1 ELSE 0 END) AS repeated_vtb_ids
                    FROM comparison
                )
                SELECT carrier, organization,
                       SUM(main_count) AS asop_operations, SUM(investor_count) AS investor_operations,
                       SUM(main_amount) AS asop_amount, SUM(investor_amount) AS investor_amount,
                       SUM(CASE WHEN status = 'Совпало' THEN 1 ELSE 0 END) AS matched,
                       SUM(CASE WHEN status = 'Расхождение суммы' THEN 1 ELSE 0 END) AS amount_mismatches,
                       SUM(CASE WHEN status = 'Только в АСОП' THEN 1 ELSE 0 END) AS only_asop,
                       SUM(CASE WHEN status = 'Только у инвестора' THEN 1 ELSE 0 END) AS only_investor,
                       metrics.matched_ids, metrics.repeated_asop_ids, metrics.repeated_vtb_ids
                FROM comparison CROSS JOIN metrics
                WHERE status IN ('Только в АСОП', 'Только у инвестора')
                GROUP BY carrier, organization, metrics.matched_ids, metrics.repeated_asop_ids, metrics.repeated_vtb_ids
                ORDER BY asop_amount DESC, investor_amount DESC
            """
            rows = self._conn.execute(query).fetchall()
            return [{"carrier": str(c), "organization": str(o), "asop_operations": int(ac or 0), "investor_operations": int(ic or 0), "asop_amount": float(aa or 0), "investor_amount": float(ia or 0), "matched": int(m or 0), "amount_mismatches": int(mm or 0), "only_asop": int(oa or 0), "only_investor": int(oi or 0), "matched_ids": int(mi or 0), "repeated_asop_ids": int(ra or 0), "repeated_vtb_ids": int(rv or 0)} for c, o, ac, ic, aa, ia, m, mm, oa, oi, mi, ra, rv in rows]
        except Exception as e:
            print(f"Investor reconciliation error: {e}")
            return []
        finally:
            # Не держим файл архива заблокированным: готовую сверку можно
            # открывать повторно, а другой архив — подготовить без перезапуска.
            try: self._conn.execute("DETACH investor")
            except: pass

    def get_investor_mismatches(self, investor_db_path, limit=10000, asop_dates=None, vtb_dates=None, offset=0):
        """Только ID, присутствующие строго в одном из двух отчётов."""
        source = self.investor_source_columns()
        paid_col = self._sum_columns.get("Оплаченная сумма")
        required = ("transaction_id", "realizator", "payment", "category", "carrier")
        if not self._conn or not all(source.get(key) for key in required) or not paid_col or not os.path.exists(investor_db_path): return []
        paid_clean = f"REGEXP_REPLACE(REPLACE(REPLACE(\"{paid_col}\", ' ', ''), ',', '.'), '[^0-9.-]', '', 'g')"
        paid = f"COALESCE(TRY_CAST({paid_clean} AS DOUBLE), 0)"
        asop_period, vtb_period = self._reconciliation_period_conditions(source, asop_dates, vtb_dates)
        trip_date = f"MIN(NULLIF(TRIM(CAST(\"{source['trip_datetime']}\" AS VARCHAR)), ''))" if source.get("trip_datetime") else "NULL"
        server_date = f"MIN(NULLIF(TRIM(CAST(\"{source['server_datetime']}\" AS VARCHAR)), ''))" if source.get("server_datetime") else "NULL"
        offset = max(0, int(offset))
        try:
            try: self._conn.execute("DETACH investor")
            except: pass
            self._conn.execute(f"ATTACH '{investor_db_path.replace(chr(39), chr(39) * 2)}' AS investor (READ_ONLY)")
            rows = self._conn.execute(f"""
                WITH main_data AS (
                    SELECT COALESCE(NULLIF(TRIM(CAST(\"{source['carrier']}\" AS VARCHAR)), ''), 'Не указан') AS carrier,
                           NULLIF(TRIM(CAST(\"{source['transaction_id']}\" AS VARCHAR)), '') AS transaction_id,
                           SUM({paid}) AS main_amount,
                           {trip_date} AS asop_trip_date,
                           {server_date} AS asop_server_datetime
                    FROM {self._table_name}
                    WHERE TRIM(COALESCE(CAST(\"{source['realizator']}\" AS VARCHAR), '')) = ''
                      AND LOWER(TRIM(COALESCE(CAST(\"{source['payment']}\" AS VARCHAR), ''))) LIKE '%безналич%'
                      AND LOWER(TRIM(COALESCE(CAST(\"{source['category']}\" AS VARCHAR), ''))) = 'без льготы'
                      {asop_period}
                    GROUP BY carrier, transaction_id
                ), investor_data AS (
                    SELECT NULLIF(TRIM(CAST(\"Организация\" AS VARCHAR)), '') AS organization,
                           NULLIF(TRIM(CAST(\"transaction_id\" AS VARCHAR)), '') AS transaction_id,
                           SUM(COALESCE(TRY_CAST(REPLACE(TRIM(CAST(\"Сумма операции, руб.\" AS VARCHAR)), ',', '.') AS DOUBLE), 0)) AS investor_amount,
                           MIN(NULLIF(TRIM(CAST(\"Дата оплаты на терминале\" AS VARCHAR)), '')) AS investor_payment_date
                    FROM investor.investor_data
                    WHERE NULLIF(TRIM(CAST(\"transaction_id\" AS VARCHAR)), '') IS NOT NULL
                      {vtb_period}
                    GROUP BY organization, transaction_id
                ), asop_unique AS (
                    SELECT m.carrier, 'Нет в отчёте ВТБ' AS organization,
                           m.transaction_id, m.main_amount, 0 AS investor_amount,
                           m.asop_trip_date, m.asop_server_datetime, NULL AS investor_payment_date,
                           'Уникально в АСОП' AS status
                    FROM main_data m
                    WHERE NOT EXISTS (SELECT 1 FROM investor_data i WHERE i.transaction_id = m.transaction_id)
                    ORDER BY m.transaction_id LIMIT {max(1, int(limit) // 2)} OFFSET {offset}
                ), investor_unique AS (
                    SELECT 'Нет в АСОП' AS carrier, COALESCE(i.organization, 'Организация не указана') AS organization,
                           i.transaction_id, 0 AS main_amount, i.investor_amount,
                           NULL AS asop_trip_date, NULL AS asop_server_datetime, i.investor_payment_date,
                           'Уникально у ВТБ' AS status
                    FROM investor_data i
                    WHERE NOT EXISTS (SELECT 1 FROM main_data m WHERE m.transaction_id = i.transaction_id)
                    ORDER BY i.transaction_id LIMIT {max(1, int(limit) - max(1, int(limit) // 2))} OFFSET {offset}
                )
                SELECT * FROM asop_unique
                UNION ALL SELECT * FROM investor_unique
                ORDER BY status, transaction_id
            """).fetchall()
            return [{"carrier": str(c), "organization": str(o), "transaction_id": str(tid or ""), "main_amount": float(main or 0), "investor_amount": float(inv or 0), "asop_trip_date": str(asop_date or ""), "asop_server_datetime": str(server_date or ""), "investor_payment_date": str(investor_date or ""), "status": str(status)} for c, o, tid, main, inv, asop_date, server_date, investor_date, status in rows]
        except Exception as e:
            print(f"Investor mismatch details error: {e}")
            return []
        finally:
            try: self._conn.execute("DETACH investor")
            except: pass

    def get_asop_status_comparison(self, investor_db_path, asop_dates=None, vtb_dates=None):
        """Сверка уникальных ID с ВТБ в разрезе каждого статуса АСОП."""
        source = self.investor_source_columns()
        required = ("transaction_id", "realizator", "payment", "category", "status")
        if not self._conn or not all(source.get(key) for key in required) or not os.path.isfile(investor_db_path): return []
        asop_period, vtb_period = self._reconciliation_period_conditions(source, asop_dates, vtb_dates)
        try:
            try: self._conn.execute("DETACH investor")
            except: pass
            escaped = investor_db_path.replace("'", "''")
            self._conn.execute(f"ATTACH '{escaped}' AS investor (READ_ONLY)")
            rows = self._conn.execute(f'''
                WITH asop AS (
                    SELECT COALESCE(NULLIF(TRIM(CAST("{source['status']}" AS VARCHAR)), ''), '(Пусто)') AS status,
                           NULLIF(TRIM(CAST("{source['transaction_id']}" AS VARCHAR)), '') AS transaction_id
                    FROM {self._table_name}
                    WHERE TRIM(COALESCE(CAST("{source['realizator']}" AS VARCHAR), '')) = ''
                      AND LOWER(TRIM(COALESCE(CAST("{source['payment']}" AS VARCHAR), ''))) LIKE '%безналич%'
                      AND LOWER(TRIM(COALESCE(CAST("{source['category']}" AS VARCHAR), ''))) = 'без льготы'
                      {asop_period}
                ), vtb AS (
                    SELECT DISTINCT NULLIF(TRIM(CAST("transaction_id" AS VARCHAR)), '') AS transaction_id
                    FROM investor.investor_data
                    WHERE NULLIF(TRIM(CAST("transaction_id" AS VARCHAR)), '') IS NOT NULL {vtb_period}
                )
                SELECT a.status,
                       COUNT(DISTINCT CASE WHEN v.transaction_id IS NOT NULL THEN a.transaction_id END) AS matched_ids,
                       COUNT(DISTINCT CASE WHEN v.transaction_id IS NULL THEN a.transaction_id END) AS unique_ids
                FROM asop a LEFT JOIN vtb v ON v.transaction_id = a.transaction_id
                WHERE a.transaction_id IS NOT NULL
                GROUP BY a.status
                ORDER BY matched_ids + unique_ids DESC, a.status
            ''').fetchall()
            return [(str(status), int(matched or 0), int(unique or 0)) for status, matched, unique in rows]
        except Exception as e:
            print(f"ASOP status comparison error: {e}")
            return []
        finally:
            try: self._conn.execute("DETACH investor")
            except: pass

    def get_asop_daily_id_comparison(self, investor_db_path, asop_dates=None, vtb_dates=None):
        """Уникальные ID АСОП по дате поездки с результатом сверки ВТБ."""
        source = self.investor_source_columns()
        required = ("transaction_id", "realizator", "payment", "category", "trip_datetime")
        if not self._conn or not all(source.get(key) for key in required) or not os.path.isfile(investor_db_path): return []
        asop_period, vtb_period = self._reconciliation_period_conditions(source, asop_dates, vtb_dates)
        trip_day = f"CAST({self._timestamp_expr(source['trip_datetime'])} AS DATE)"
        try:
            try: self._conn.execute("DETACH investor")
            except: pass
            self._conn.execute(f"ATTACH '{investor_db_path.replace(chr(39), chr(39) * 2)}' AS investor (READ_ONLY)")
            rows = self._conn.execute(f'''
                WITH asop AS (
                    SELECT {trip_day} AS trip_day, NULLIF(TRIM(CAST("{source['transaction_id']}" AS VARCHAR)), '') AS transaction_id
                    FROM {self._table_name}
                    WHERE TRIM(COALESCE(CAST("{source['realizator']}" AS VARCHAR), '')) = ''
                      AND LOWER(TRIM(COALESCE(CAST("{source['payment']}" AS VARCHAR), ''))) LIKE '%безналич%'
                      AND LOWER(TRIM(COALESCE(CAST("{source['category']}" AS VARCHAR), ''))) = 'без льготы'
                      {asop_period}
                ), vtb AS (
                    SELECT DISTINCT NULLIF(TRIM(CAST("transaction_id" AS VARCHAR)), '') AS transaction_id
                    FROM investor.investor_data
                    WHERE NULLIF(TRIM(CAST("transaction_id" AS VARCHAR)), '') IS NOT NULL {vtb_period}
                )
                SELECT trip_day, COUNT(DISTINCT a.transaction_id) AS all_ids,
                       COUNT(DISTINCT CASE WHEN v.transaction_id IS NOT NULL THEN a.transaction_id END) AS matched_ids,
                       COUNT(DISTINCT CASE WHEN v.transaction_id IS NULL THEN a.transaction_id END) AS unique_ids
                FROM asop a LEFT JOIN vtb v ON v.transaction_id = a.transaction_id
                WHERE trip_day IS NOT NULL AND a.transaction_id IS NOT NULL
                GROUP BY trip_day ORDER BY trip_day
            ''').fetchall()
            return [(str(day), int(all_ids or 0), int(matched or 0), int(unique or 0)) for day, all_ids, matched, unique in rows]
        except Exception as e:
            print(f"ASOP daily ID comparison error: {e}")
            return []
        finally:
            try: self._conn.execute("DETACH investor")
            except: pass

    def get_investor_report_rows(self, investor_db_path, organization, search_text="", limit=5000, asop_dates=None, vtb_dates=None):
        """Только уникальные у ВТБ строки выбранной организации со всеми полями отчёта."""
        source = self.investor_source_columns()
        required = ("transaction_id", "realizator", "payment", "category")
        if not self._conn or not all(source.get(key) for key in required) or not os.path.isfile(investor_db_path): return pd.DataFrame()
        asop_period, vtb_period = self._reconciliation_period_conditions(source, asop_dates, vtb_dates)
        organization_clause = 'i."Организация" = ?'
        params = [organization]
        if organization == "Организация не указана":
            organization_clause = '(i."Организация" IS NULL OR TRIM(CAST(i."Организация" AS VARCHAR)) = \'\')'
            params = []
        search_clause = ""
        if search_text.strip():
            search_clause = ' AND CAST(i."transaction_id" AS VARCHAR) ILIKE ?'
            params.append(f"%{search_text.strip()}%")
        safe_limit = max(1, min(int(limit), 50000))
        try:
            try: self._conn.execute("DETACH investor")
            except: pass
            escaped = investor_db_path.replace("'", "''")
            self._conn.execute(f"ATTACH '{escaped}' AS investor (READ_ONLY)")
            return self._conn.execute(f'''WITH main_ids AS (
                    SELECT DISTINCT NULLIF(TRIM(CAST("{source['transaction_id']}" AS VARCHAR)), '') AS transaction_id
                    FROM {self._table_name}
                    WHERE TRIM(COALESCE(CAST("{source['realizator']}" AS VARCHAR), '')) = ''
                      AND LOWER(TRIM(COALESCE(CAST("{source['payment']}" AS VARCHAR), ''))) LIKE '%безналич%'
                      AND LOWER(TRIM(COALESCE(CAST("{source['category']}" AS VARCHAR), ''))) = 'без льготы'
                      {asop_period}
                )
                SELECT i."Организация", i."Дата оплаты на терминале", i."Время оплаты на терминале",
                       i."Дата перечисления на расчетный счет", i."Сумма операции, руб.",
                       i."Вознаграждение Инвестора, руб.", i."transaction_id", i."RRN"
                FROM investor.investor_data i
                WHERE {organization_clause} {search_clause}
                  {vtb_period}
                  AND NOT EXISTS (SELECT 1 FROM main_ids m WHERE m.transaction_id = NULLIF(TRIM(CAST(i."transaction_id" AS VARCHAR)), ''))
                ORDER BY i."Дата оплаты на терминале" DESC, i."Время оплаты на терминале" DESC
                LIMIT {safe_limit}''', params).df()
        except Exception as e:
            print(f"Investor unique report rows error: {e}")
            return pd.DataFrame()
        finally:
            try: self._conn.execute("DETACH investor")
            except: pass

    def get_asop_reconciliation_rows(self, carrier, investor_db_path, search_text="", limit=5000, asop_dates=None, vtb_dates=None):
        """Только уникальные в АСОП строки перевозчика со всеми полями реестра."""
        source = self.investor_source_columns()
        required = ("transaction_id", "realizator", "payment", "category", "carrier")
        if not self._conn or not all(source.get(key) for key in required): return pd.DataFrame()
        conditions = [
            f"TRIM(COALESCE(CAST(\"{source['realizator']}\" AS VARCHAR), '')) = ''",
            f"LOWER(TRIM(COALESCE(CAST(\"{source['payment']}\" AS VARCHAR), ''))) LIKE '%безналич%'",
            f"LOWER(TRIM(COALESCE(CAST(\"{source['category']}\" AS VARCHAR), ''))) = 'без льготы'",
            f"COALESCE(NULLIF(TRIM(CAST(\"{source['carrier']}\" AS VARCHAR)), ''), 'Не указан') = ?",
        ]
        asop_period, vtb_period = self._reconciliation_period_conditions(source, asop_dates, vtb_dates)
        if asop_period: conditions.append(asop_period.removeprefix("AND "))
        params = [carrier]
        if search_text.strip():
            conditions.append(f"CAST(\"{source['transaction_id']}\" AS VARCHAR) ILIKE ?")
            params.append(f"%{search_text.strip()}%")
        try:
            safe_limit = max(1, min(int(limit), 50000))
            try: self._conn.execute("DETACH investor")
            except: pass
            escaped = investor_db_path.replace("'", "''")
            self._conn.execute(f"ATTACH '{escaped}' AS investor (READ_ONLY)")
            return self._conn.execute(
                f'''SELECT * FROM {self._table_name}
                    WHERE {' AND '.join(conditions)}
                      AND NOT EXISTS (SELECT 1 FROM investor.investor_data i
                                      WHERE NULLIF(TRIM(CAST(i."transaction_id" AS VARCHAR)), '') = NULLIF(TRIM(CAST("{source['transaction_id']}" AS VARCHAR)), '')
                                        {vtb_period})
                    LIMIT {safe_limit}''', params
            ).df().drop(columns=["row_id"], errors="ignore")
        except Exception as e:
            print(f"ASOP unique report rows error: {e}")
            return pd.DataFrame()
        finally:
            try: self._conn.execute("DETACH investor")
            except: pass

    def get_asop_unique_daily(self, carrier, investor_db_path, asop_dates=None, vtb_dates=None):
        """Количество уникальных для АСОП операций перевозчика по датам поездки."""
        source = self.investor_source_columns()
        required = ("transaction_id", "realizator", "payment", "category", "carrier", "trip_datetime")
        if not self._conn or not all(source.get(key) for key in required): return []
        conditions = [
            f"TRIM(COALESCE(CAST(\"{source['realizator']}\" AS VARCHAR), '')) = ''",
            f"LOWER(TRIM(COALESCE(CAST(\"{source['payment']}\" AS VARCHAR), ''))) LIKE '%безналич%'",
            f"LOWER(TRIM(COALESCE(CAST(\"{source['category']}\" AS VARCHAR), ''))) = 'без льготы'",
            f"COALESCE(NULLIF(TRIM(CAST(\"{source['carrier']}\" AS VARCHAR)), ''), 'Не указан') = ?",
        ]
        asop_period, vtb_period = self._reconciliation_period_conditions(source, asop_dates, vtb_dates)
        if asop_period: conditions.append(asop_period.removeprefix("AND "))
        date_expr = f"CAST({self._timestamp_expr(source['trip_datetime'])} AS DATE)"
        try:
            try: self._conn.execute("DETACH investor")
            except: pass
            self._conn.execute(f"ATTACH '{investor_db_path.replace(chr(39), chr(39) * 2)}' AS investor (READ_ONLY)")
            return [(str(day), int(count)) for day, count in self._conn.execute(f'''SELECT {date_expr} AS day, COUNT(*) AS transactions
                FROM {self._table_name}
                WHERE {' AND '.join(conditions)}
                  AND NOT EXISTS (SELECT 1 FROM investor.investor_data i
                                  WHERE NULLIF(TRIM(CAST(i."transaction_id" AS VARCHAR)), '') = NULLIF(TRIM(CAST("{source['transaction_id']}" AS VARCHAR)), '') {vtb_period})
                GROUP BY day ORDER BY day''', [carrier]).fetchall()]
        except Exception as e:
            print(f"ASOP unique daily error: {e}"); return []
        finally:
            try: self._conn.execute("DETACH investor")
            except: pass

    def get_type_vehicles(self, where_clause, carrier_name, type_name):
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), '')"
        type_expr = "COALESCE(NULLIF(TRIM(CAST(\"Тип ТС\" AS VARCHAR)), ''), 'Не указан')"
        carrier_escaped = str(carrier_name).replace("'", "''"); type_escaped = str(type_name).replace("'", "''")
        try:
            return [(str(name), int(count)) for name, count in self._conn.execute(f"""
                SELECT {vehicle}, COUNT(*) FROM {self._table_name}
                WHERE {where_clause} AND {carrier} = '{carrier_escaped}' AND {type_expr} = '{type_escaped}' AND {vehicle} IS NOT NULL
                GROUP BY 1 ORDER BY 2 DESC
            """).fetchall()]
        except Exception as e:
            print(f"Type vehicles error: {e}"); return []

    def get_fleet_drivers(self, where_clause, carrier_name):
        if "ФИО водителя" not in self._columns: return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        driver = "NULLIF(TRIM(CAST(\"ФИО водителя\" AS VARCHAR)), '')"
        escaped = str(carrier_name).replace("'", "''")
        try:
            return [(str(name), int(count)) for name, count in self._conn.execute(f"""
                SELECT {driver}, COUNT(*) FROM {self._table_name}
                WHERE {where_clause} AND {carrier} = '{escaped}' AND {driver} IS NOT NULL
                GROUP BY 1 ORDER BY 2 DESC
            """).fetchall()]
        except Exception as e:
            print(f"Fleet drivers error: {e}"); return []

    def get_driver_vehicles(self, where_clause, carrier_name, driver_name):
        if "ФИО водителя" not in self._columns: return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        driver = "NULLIF(TRIM(CAST(\"ФИО водителя\" AS VARCHAR)), '')"
        vehicle = "COALESCE(NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), ''), 'Не указан')"
        carrier_escaped = str(carrier_name).replace("'", "''"); driver_escaped = str(driver_name).replace("'", "''")
        try:
            return [(str(name), int(count)) for name, count in self._conn.execute(f"""
                SELECT {vehicle}, COUNT(*) FROM {self._table_name}
                WHERE {where_clause} AND {carrier} = '{carrier_escaped}' AND {driver} = '{driver_escaped}'
                GROUP BY 1 ORDER BY 2 DESC
            """).fetchall()]
        except Exception as e:
            print(f"Driver vehicles error: {e}"); return []

    def get_fleet_totals(self, where_clause="1=1", carriers=None):
        if not self._conn:
            return {"vehicles": 0, "vehicles_by_type": 0, "routes": 0, "transactions": 0, "work_hours": 0.0}
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), '')"
        route = "NULLIF(TRIM(CAST(\"Маршрут\" AS VARCHAR)), '')"
        extra = ""
        if carriers:
            escaped = "', '".join(str(value).replace("'", "''") for value in carriers)
            extra = f" AND {carrier} IN ('{escaped}')"
        try:
            vehicles, routes, transactions = self._conn.execute(f"""
                SELECT COUNT(DISTINCT {vehicle}), COUNT(DISTINCT {route}), COUNT(*)
                FROM {self._table_name} WHERE {where_clause}{extra}
            """).fetchone()
            by_type = self._conn.execute(f"""
                SELECT COALESCE(SUM(vehicles), 0) FROM (
                    SELECT COUNT(DISTINCT {vehicle}) AS vehicles
                    FROM {self._table_name} WHERE {where_clause}{extra} GROUP BY "Тип ТС"
                )
            """).fetchone()[0]
            work_hours = self.get_fleet_work_hours(where_clause, carriers)
            return {"vehicles": int(vehicles or 0), "vehicles_by_type": int(by_type or 0), "routes": int(routes or 0), "transactions": int(transactions or 0), "work_hours": work_hours}
        except Exception as e:
            print(f"Fleet totals error: {e}"); return {"vehicles": 0, "vehicles_by_type": 0, "routes": 0, "transactions": 0, "work_hours": 0.0}

    def get_fleet_work_hours(self, where_clause="1=1", carriers=None):
        """Оценка рабочего времени: первый–последний рейс каждого ТС за день."""
        required = {CARRIER_COLUMN, "Транспортное средство", DATETIME_COLUMN}
        if not self._conn or not required.issubset(self._columns): return 0.0
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), '')"
        timestamp = self._timestamp_expr(DATETIME_COLUMN)
        extra = ""
        if carriers:
            escaped = "', '".join(str(value).replace("'", "''") for value in carriers)
            extra = f" AND {carrier} IN ('{escaped}')"
        try:
            hours = self._conn.execute(f"""
                SELECT COALESCE(SUM(work_hours), 0) FROM (
                    SELECT {carrier} AS carrier, {vehicle} AS vehicle, CAST({timestamp} AS DATE) AS day,
                           EXTRACT(EPOCH FROM (MAX({timestamp}) - MIN({timestamp}))) / 3600.0 AS work_hours
                    FROM {self._table_name}
                    WHERE {where_clause}{extra} AND {vehicle} IS NOT NULL AND {timestamp} IS NOT NULL
                    GROUP BY carrier, vehicle, day
                ) daily_work
            """).fetchone()[0]
            return float(hours or 0)
        except Exception as e:
            print(f"Fleet work hours error: {e}")
            return 0.0

    def get_fleet_vehicle_work_hours(self, where_clause="1=1", carrier_name=None, carriers=None):
        """Расшифровка оценки рабочего времени до уровня конкретного ТС."""
        required = {CARRIER_COLUMN, "Транспортное средство", DATETIME_COLUMN}
        if not self._conn or not required.issubset(self._columns): return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), '')"
        timestamp = self._timestamp_expr(DATETIME_COLUMN)
        extra = ""
        if carrier_name is not None:
            extra = f" AND {carrier} = '{str(carrier_name).replace(chr(39), chr(39) * 2)}'"
        elif carriers:
            escaped = "', '".join(str(value).replace("'", "''") for value in carriers)
            extra = f" AND {carrier} IN ('{escaped}')"
        try:
            rows = self._conn.execute(f"""
                SELECT carrier, vehicle, SUM(work_hours) AS work_hours, COUNT(*) AS work_days,
                       SUM(transactions) AS transactions
                FROM (
                    SELECT {carrier} AS carrier, {vehicle} AS vehicle, CAST({timestamp} AS DATE) AS day,
                           EXTRACT(EPOCH FROM (MAX({timestamp}) - MIN({timestamp}))) / 3600.0 AS work_hours,
                           COUNT(*) AS transactions
                    FROM {self._table_name}
                    WHERE {where_clause}{extra} AND {vehicle} IS NOT NULL AND {timestamp} IS NOT NULL
                    GROUP BY carrier, vehicle, day
                ) daily_work
                GROUP BY carrier, vehicle
                ORDER BY work_hours DESC, transactions DESC, vehicle
            """).fetchall()
            return [{"carrier": str(c), "vehicle": str(v), "work_hours": float(h or 0),
                     "work_days": int(days or 0), "transactions": int(transactions or 0)}
                    for c, v, h, days, transactions in rows]
        except Exception as e:
            print(f"Fleet vehicle work hours error: {e}")
            return []

    def get_global_duplicate_fleet_vehicles(self, where_clause="1=1", carriers=None):
        """Одинаковые номера, попавшие в разные связки перевозчик × тип ТС."""
        required = {CARRIER_COLUMN, "Транспортное средство", "Тип ТС"}
        if not self._conn or not required.issubset(self._columns): return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), '')"
        vehicle_type = "COALESCE(NULLIF(TRIM(CAST(\"Тип ТС\" AS VARCHAR)), ''), 'Не указан')"
        extra = ""
        if carriers:
            escaped = "', '".join(str(value).replace("'", "''") for value in carriers)
            extra = f" AND {carrier} IN ('{escaped}')"
        try:
            rows = self._conn.execute(f"""
                SELECT {vehicle} AS vehicle,
                       STRING_AGG(DISTINCT {carrier} || ' — ' || {vehicle_type}, '\n' ORDER BY {carrier} || ' — ' || {vehicle_type}) AS locations,
                       COUNT(DISTINCT {carrier} || '|' || {vehicle_type}) AS placements,
                       COUNT(*) AS transactions
                FROM {self._table_name}
                WHERE {where_clause}{extra} AND {vehicle} IS NOT NULL
                GROUP BY vehicle
                HAVING COUNT(DISTINCT {carrier} || '|' || {vehicle_type}) > 1
                ORDER BY placements DESC, transactions DESC, vehicle
            """).fetchall()
            return [(str(v), str(locations), int(placements or 0), int(transactions or 0)) for v, locations, placements, transactions in rows]
        except Exception as e:
            print(f"Global duplicate fleet vehicles error: {e}")
            return []

    def get_duplicate_vehicle_details(self, where_clause, vehicle_name, carriers=None):
        """Детали одного повторяющегося номера: где и кем он использовался."""
        required = {CARRIER_COLUMN, "Транспортное средство", "Тип ТС"}
        if not self._conn or not required.issubset(self._columns): return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), '')"
        vehicle_type = "COALESCE(NULLIF(TRIM(CAST(\"Тип ТС\" AS VARCHAR)), ''), 'Не указан')"
        route = "COALESCE(NULLIF(TRIM(CAST(\"Маршрут\" AS VARCHAR)), ''), 'Не указан')" if "Маршрут" in self._columns else "'Не указан'"
        driver = "COALESCE(NULLIF(TRIM(CAST(\"ФИО водителя\" AS VARCHAR)), ''), 'Не указан')" if "ФИО водителя" in self._columns else "'Не указан'"
        extra = ""
        if carriers:
            escaped = "', '".join(str(value).replace("'", "''") for value in carriers)
            extra = f" AND {carrier} IN ('{escaped}')"
        escaped_vehicle = str(vehicle_name).replace("'", "''")
        try:
            rows = self._conn.execute(f"""
                SELECT {carrier} AS carrier, {vehicle_type} AS vehicle_type,
                       {route} AS route, {driver} AS driver, COUNT(*) AS transactions
                FROM {self._table_name}
                WHERE {where_clause}{extra} AND {vehicle} = '{escaped_vehicle}'
                GROUP BY carrier, vehicle_type, route, driver
                ORDER BY carrier, vehicle_type, transactions DESC, route, driver
            """).fetchall()
            return [{"carrier": str(c), "type": str(t), "route": str(r), "driver": str(d), "transactions": int(count or 0)}
                    for c, t, r, d, count in rows]
        except Exception as e:
            print(f"Duplicate vehicle details error: {e}")
            return []

    def get_vehicle_routes(self, where_clause, carrier_name, vehicle_name):
        if not self._conn or "Маршрут" not in self._columns:
            return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), '')"
        route = "COALESCE(NULLIF(TRIM(CAST(\"Маршрут\" AS VARCHAR)), ''), 'Не указан')"
        carrier_escaped = str(carrier_name).replace("'", "''"); vehicle_escaped = str(vehicle_name).replace("'", "''")
        try:
            return [(str(route), int(count)) for route, count in self._conn.execute(f"""
                SELECT {route} AS route, COUNT(*) AS transactions
                FROM {self._table_name}
                WHERE {where_clause} AND {carrier} = '{carrier_escaped}' AND {vehicle} = '{vehicle_escaped}'
                GROUP BY route ORDER BY transactions DESC
            """).fetchall()]
        except Exception as e:
            print(f"Vehicle routes error: {e}"); return []

    def get_fleet_routes(self, where_clause, carrier_name):
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        route = "COALESCE(NULLIF(TRIM(CAST(\"Маршрут\" AS VARCHAR)), ''), 'Не указан')"
        escaped = str(carrier_name).replace("'", "''")
        try:
            return [(str(name), int(count)) for name, count in self._conn.execute(f"""
                SELECT {route} AS route, COUNT(*) AS transactions FROM {self._table_name}
                WHERE {where_clause} AND {carrier} = '{escaped}' GROUP BY route ORDER BY transactions DESC
            """).fetchall()]
        except Exception as e:
            print(f"Fleet routes error: {e}"); return []

    def get_route_vehicles(self, where_clause, carrier_name, route_name):
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "COALESCE(NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), ''), 'Не указан')"
        route = "COALESCE(NULLIF(TRIM(CAST(\"Маршрут\" AS VARCHAR)), ''), 'Не указан')"
        carrier_escaped = str(carrier_name).replace("'", "''"); route_escaped = str(route_name).replace("'", "''")
        try:
            return [(str(name), int(count)) for name, count in self._conn.execute(f"""
                SELECT {vehicle} AS vehicle, COUNT(*) AS transactions FROM {self._table_name}
                WHERE {where_clause} AND {carrier} = '{carrier_escaped}' AND {route} = '{route_escaped}'
                GROUP BY vehicle ORDER BY transactions DESC
            """).fetchall()]
        except Exception as e:
            print(f"Route vehicles error: {e}"); return []

    def get_daily_transactions(self, where_clause="1=1", carrier_name=None, vehicle_name=None, route_name=None, driver_name=None, type_name=None):
        if not self._conn or DATETIME_COLUMN not in self._columns:
            return []
        conditions = [where_clause]
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "COALESCE(NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), ''), 'Не указан')"
        route = "COALESCE(NULLIF(TRIM(CAST(\"Маршрут\" AS VARCHAR)), ''), 'Не указан')"
        driver = "COALESCE(NULLIF(TRIM(CAST(\"ФИО водителя\" AS VARCHAR)), ''), 'Не указан')"
        type_expr = "COALESCE(NULLIF(TRIM(CAST(\"Тип ТС\" AS VARCHAR)), ''), 'Не указан')"
        for expr, value in ((carrier, carrier_name), (vehicle, vehicle_name), (route, route_name), (driver, driver_name), (type_expr, type_name)):
            if value is not None: conditions.append(f"{expr} = '{str(value).replace(chr(39), chr(39) * 2)}'")
        date_expr = self._timestamp_expr(DATETIME_COLUMN)
        try:
            return [(str(day), int(count)) for day, count in self._conn.execute(f"""
                SELECT CAST({date_expr} AS DATE) AS day, COUNT(*) AS transactions
                FROM {self._table_name} WHERE {' AND '.join(conditions)} AND {date_expr} IS NOT NULL
                GROUP BY day ORDER BY day
            """).fetchall()]
        except Exception as e:
            print(f"Fleet chart error: {e}"); return []

    def get_operational_analysis(self, where_clause="1=1"):
        """Сравнение связок маршрут → ТС → водитель.

        Это отдельная аналитическая метрика и она не меняет расчёты сводки,
        карточек или экспорта. Для каждой операции выбирается ровно один
        денежный источник по правилам реализатора.
        """
        required = {CARRIER_COLUMN, "Маршрут", "Транспортное средство", "ФИО водителя"}
        if not self._conn or not required.issubset(self._columns): return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        route = "COALESCE(NULLIF(TRIM(CAST(\"Маршрут\" AS VARCHAR)), ''), 'Не указан')"
        vehicle = "COALESCE(NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), ''), 'Не указан')"
        vehicle_type = "COALESCE(NULLIF(TRIM(CAST(\"Тип ТС\" AS VARCHAR)), ''), 'Не указан')" if "Тип ТС" in self._columns else "'Не указан'"
        driver = "COALESCE(NULLIF(TRIM(CAST(\"ФИО водителя\" AS VARCHAR)), ''), 'Не указан')"
        def amount(alias):
            col = self._sum_columns.get(alias)
            if not col: return "0"
            clean = f"REGEXP_REPLACE(REPLACE(REPLACE(\"{col}\", ' ', ''), ',', '.'), '[^0-9.-]', '', 'g')"
            return f"COALESCE(TRY_CAST({clean} AS DOUBLE), 0)"
        applied = amount("Применённый тариф")
        regulated = amount("Регулируемый тариф")
        benefit = amount("Льгота")
        paid = amount("Оплаченная сумма")
        realizator = "LOWER(TRIM(COALESCE(CAST(\"Реализатор\" AS VARCHAR), '')))" if "Реализатор" in self._columns else "''"
        payment = "LOWER(TRIM(COALESCE(CAST(\"Форма оплаты проезда\" AS VARCHAR), '')))" if "Форма оплаты проезда" in self._columns else "''"
        passenger_category = "LOWER(TRIM(COALESCE(CAST(\"Категория пассажира\" AS VARCHAR), '')))" if "Категория пассажира" in self._columns else "''"
        # NULLIF защищает от двойного учёта: выбирается первое ненулевое
        # допустимое поле, а не сумма всех зелёных ячеек из схемы отчёта.
        calculated = f"""
            CASE
                WHEN {realizator} = 'юппк' THEN {applied}
                WHEN {realizator} = 'кат' AND {payment} LIKE '%льгот%' THEN {regulated}
                WHEN {realizator} = 'кат' THEN {paid}
                WHEN {realizator} = 'кт' AND {payment} LIKE '%льгот%'
                    THEN COALESCE(NULLIF({applied}, 0), NULLIF({regulated}, 0), NULLIF({benefit}, 0), 0)
                WHEN {realizator} = 'кт' THEN {paid}
                /* пустой реализатор = АСОП; применённый тариф в приоритете */
                ELSE COALESCE(NULLIF({applied}, 0), NULLIF({paid}, 0), 0)
            END
        """
        try:
            rows = self._conn.execute(f"""
                SELECT {carrier} AS carrier, {route} AS route, {vehicle} AS vehicle, {vehicle_type} AS vehicle_type, {driver} AS driver,
                       COUNT(*) AS transactions,
                       SUM({calculated}) AS calculated_amount,
                       AVG({calculated}) AS average_operation,
                       100.0 * SUM(CASE WHEN {passenger_category} = 'с льготой' THEN 1 ELSE 0 END) / COUNT(*) AS benefit_share
                FROM {self._table_name} WHERE {where_clause}
                GROUP BY carrier, route, vehicle, vehicle_type, driver
                ORDER BY calculated_amount DESC, transactions DESC LIMIT 10000
            """).fetchall()
            return [{"carrier": str(c), "route": str(r), "vehicle": str(v), "vehicle_type": str(vt), "driver": str(d), "transactions": int(t or 0), "calculated_amount": float(total or 0), "average_operation": float(average or 0), "benefit_share": float(share or 0)} for c, r, v, vt, d, t, total, average, share in rows]
        except Exception as e:
            print(f"Operational analysis error: {e}"); return []

    def get_carrier_check_leaders(self, where_clause, carrier_name, limit=None):
        """Топы для раскрытой карточки.

        Средний чек намеренно считается только по операциям категории «Без льготы».
        Доля льготных — это доля строк категории «С льготой» среди всех операций
        соответствующего ТС, маршрута или водителя.
        """
        required = {CARRIER_COLUMN, "Транспортное средство", "Маршрут", "ФИО водителя", "Категория пассажира"}
        if not self._conn or not required.issubset(self._columns):
            return {"vehicles": [], "routes": [], "drivers": []}
        paid_column = self._sum_columns.get("Оплаченная сумма")
        if not paid_column:
            return {"vehicles": [], "routes": [], "drivers": []}
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        category = "LOWER(TRIM(CAST(\"Категория пассажира\" AS VARCHAR)))"
        paid_clean = f"REGEXP_REPLACE(REPLACE(REPLACE(\"{paid_column}\", ' ', ''), ',', '.'), '[^0-9.-]', '', 'g')"
        paid = f"COALESCE(TRY_CAST({paid_clean} AS DOUBLE), 0)"
        escaped = str(carrier_name).replace("'", "''")
        entity_columns = {
            "vehicles": "Транспортное средство",
            "routes": "Маршрут",
            "drivers": "ФИО водителя",
        }
        result = {key: [] for key in entity_columns}
        suffix = f" LIMIT {int(limit)}" if limit else ""
        try:
            for key, column in entity_columns.items():
                entity = f"COALESCE(NULLIF(TRIM(CAST(\"{column}\" AS VARCHAR)), ''), 'Не указан')"
                rows = self._conn.execute(f"""
                    SELECT {entity} AS entity,
                           COUNT(*) AS transactions,
                           AVG(CASE WHEN {category} = 'без льготы' THEN {paid} END) AS average_check,
                           100.0 * SUM(CASE WHEN {category} = 'с льготой' THEN 1 ELSE 0 END) / COUNT(*) AS benefit_share
                    FROM {self._table_name}
                    WHERE {where_clause} AND {carrier} = '{escaped}'
                    GROUP BY entity
                    HAVING SUM(CASE WHEN {category} = 'без льготы' THEN 1 ELSE 0 END) > 0
                    ORDER BY average_check DESC NULLS LAST, transactions DESC{suffix}
                """).fetchall()
                result[key] = [
                    {"name": str(name), "transactions": int(transactions or 0),
                     "average_check": float(average_check or 0), "benefit_share": float(benefit_share or 0)}
                    for name, transactions, average_check, benefit_share in rows
                ]
        except Exception as e:
            print(f"Carrier check leaders error: {e}")
        return result

    def get_vehicle_driver_checks(self, where_clause, carrier_name, vehicle_name):
        """Водители выбранного ТС с тем же правилом среднего чека."""
        # Нужен более узкий запрос, а не фильтрация общего топа: водитель может
        # не попасть в топ всего предприятия.
        required = {CARRIER_COLUMN, "Транспортное средство", "ФИО водителя", "Категория пассажира"}
        paid_column = self._sum_columns.get("Оплаченная сумма")
        if not self._conn or not required.issubset(self._columns) or not paid_column:
            return []
        carrier = f'COALESCE(NULLIF(TRIM(CAST("{CARRIER_COLUMN}" AS VARCHAR)), \'\'), \'Не указан\')'
        vehicle = "COALESCE(NULLIF(TRIM(CAST(\"Транспортное средство\" AS VARCHAR)), ''), 'Не указан')"
        driver = "COALESCE(NULLIF(TRIM(CAST(\"ФИО водителя\" AS VARCHAR)), ''), 'Не указан')"
        category = "LOWER(TRIM(CAST(\"Категория пассажира\" AS VARCHAR)))"
        paid_clean = f"REGEXP_REPLACE(REPLACE(REPLACE(\"{paid_column}\", ' ', ''), ',', '.'), '[^0-9.-]', '', 'g')"
        paid = f"COALESCE(TRY_CAST({paid_clean} AS DOUBLE), 0)"
        carrier_escaped = str(carrier_name).replace("'", "''")
        vehicle_escaped = str(vehicle_name).replace("'", "''")
        try:
            rows = self._conn.execute(f"""
                SELECT {driver} AS entity, COUNT(*) AS transactions,
                       AVG(CASE WHEN {category} = 'без льготы' THEN {paid} END) AS average_check,
                       100.0 * SUM(CASE WHEN {category} = 'с льготой' THEN 1 ELSE 0 END) / COUNT(*) AS benefit_share
                FROM {self._table_name}
                WHERE {where_clause} AND {carrier} = '{carrier_escaped}' AND {vehicle} = '{vehicle_escaped}'
                GROUP BY entity
                HAVING SUM(CASE WHEN {category} = 'без льготы' THEN 1 ELSE 0 END) > 0
                ORDER BY average_check DESC NULLS LAST, transactions DESC
            """).fetchall()
            return [{"name": str(name), "transactions": int(transactions or 0),
                     "average_check": float(average_check or 0), "benefit_share": float(benefit_share or 0)}
                    for name, transactions, average_check, benefit_share in rows]
        except Exception as e:
            print(f"Vehicle driver checks error: {e}")
            return []

    def get_value_counts(self, column, where_clause="1=1", text="", limit=None, offset=0):
        if not self._conn or not self._table_exists(): return []
        try:
            col = f'"{column}"'
            if column in DATE_COLUMNS:
                timestamp_expr = self._timestamp_expr(column)
                query = f"""
                    SELECT strftime({timestamp_expr}, '%d.%m.%Y') as date_str, 
                           COUNT(*) as cnt 
                    FROM {self._table_name} 
                    WHERE {where_clause} AND {col} IS NOT NULL
                      AND {timestamp_expr} IS NOT NULL
                """
                if text:
                    query += f" AND strftime({timestamp_expr}, '%d.%m.%Y') ILIKE '%{text}%'"
                query += f" GROUP BY date_str ORDER BY MIN({timestamp_expr}) DESC"
            else:
                query = f"SELECT {col}, COUNT(*) as cnt FROM {self._table_name} WHERE {where_clause} AND {col} IS NOT NULL"
                if text:
                    query += f" AND CAST({col} AS VARCHAR) ILIKE '%{text}%'"
                query += f" GROUP BY {col} ORDER BY cnt DESC"
            if limit is not None:
                query += f" LIMIT {limit} OFFSET {offset}"
            result = self._conn.execute(query).fetchall()
            return [(str(row[0]), row[1]) for row in result]
        except Exception as e:
            print(f"Value counts error: {e}")
            return []

    def get_statistics(self, where_clause="1=1"):
        result = {"transactions": 0}
        for alias in SUM_COLUMNS.keys(): result[alias] = 0.0
        if not self._conn or not self._table_exists(): return result
        try:
            parts = ["COUNT(*) as transactions"]
            for alias, db_col in self._sum_columns.items():
                if db_col:
                    clean_expr = f"REGEXP_REPLACE(REPLACE(REPLACE(\"{db_col}\", ' ', ''), ',', '.'), '[^0-9.-]', '', 'g')"
                    parts.append(f'SUM(COALESCE(TRY_CAST({clean_expr} AS DOUBLE), 0)) as "{alias}"')
                else:
                    parts.append(f'0.0 as "{alias}"')
            query = f"SELECT {', '.join(parts)} FROM {self._table_name} WHERE {where_clause}"
            r = self._conn.execute(query).fetchone()
            if r:
                result["transactions"] = r[0] or 0
                for i, alias in enumerate(SUM_COLUMNS.keys()):
                    if i + 1 < len(r): result[alias] = float(r[i+1] or 0)
        except Exception as e:
            if "cancelled" not in str(e).lower(): print(f"Stats error: {e}")
        return result

    def build_where_clause(self, filters, date_filters=None, exclude_column=None):
        if (not filters and not date_filters) or not self._conn: return "1=1"
        conditions = []
        if filters:
            for column, values in filters.items():
                if exclude_column is not None and column == exclude_column: continue
                if not values or column not in self._columns: continue
                regular = [v for v in values if v != "(Пусто)"]
                include_null = "(Пусто)" in values
                parts = []
                if regular:
                    escaped = [v.replace("'", "''") for v in regular]
                    parts.append(f'CAST("{column}" AS VARCHAR) IN (\'' + "', '".join(escaped) + '\')')
                if include_null:
                    parts.append(f'("{column}" IS NULL OR TRIM(CAST("{column}" AS VARCHAR)) = \'\')')
                if parts: conditions.append(f"({' OR '.join(parts)})")
        if date_filters:
            for col, date_range in date_filters.items():
                if exclude_column is not None and col == exclude_column: continue
                if col in DATE_COLUMNS and date_range:
                    date_from = date_range.get('from')
                    date_to = date_range.get('to')
                    if date_from and date_to:
                        timestamp_expr = self._timestamp_expr(col)
                        conditions.append(
                            f'({timestamp_expr} BETWEEN '
                            f"TIMESTAMP '{date_from} 00:00:00' AND "
                            f"TIMESTAMP '{date_to} 23:59:59')"
                        )
        return " AND ".join(conditions) if conditions else "1=1"

    def export_to_excel(self, where_clause, output_path):
        if not self._conn: return False
        try:
            q = f"SELECT * FROM {self._table_name} WHERE {where_clause} ORDER BY {self._id_column}"
            df = self._conn.execute(q).df()
            if "row_id" in df.columns: df = df.drop("row_id", axis=1)
            df.to_excel(output_path, index=False, engine="openpyxl"); return True
        except Exception as e:
            print(f"Export error: {e}"); return False

    def export_aggregated_to_excel(self, where_clause, output_path):
        if not self._conn: return False
        try:
            df = self.get_aggregated(where_clause, limit=1000000)
            if df is not None and not df.empty:
                df.to_excel(output_path, index=False, engine="openpyxl"); return True
            return False
        except Exception as e:
            print(f"Export aggregated error: {e}"); return False

    def close(self):
        if self._conn:
            try: self._conn.close()
            except: pass
            self._conn = None

class ImportWorker(QThread):
    progress = Signal(int, str); finished = Signal(int, float); error = Signal(str)
    def __init__(self, repo, csv, db, parent=None):
        super().__init__(parent); self.repo, self.csv, self.db = repo, csv, db
    def run(self):
        try:
            t0 = time.time()
            self.progress.emit(5, "Подготовка...")
            n = self.repo.create_from_csv(self.db, self.csv, lambda p, m: self.progress.emit(p, m))
            self.finished.emit(n, time.time() - t0)
        except Exception as e:
            import traceback; traceback.print_exc()
            self.error.emit(f"{type(e).__name__}: {str(e)}")

class InvestorArchiveWorker(QThread):
    """Собирает split-ZIP и сохраняет отчёт инвестора в отдельной DuckDB БД."""
    progress = Signal(int, str); finished = Signal(int, str, float); error = Signal(str)
    def __init__(self, archive_path, output_db, parent=None):
        super().__init__(parent); self.archive_path = archive_path; self.output_db = output_db
    def _extract_split_zip_csv(self, parts, output_path):
        """Извлекает единственный deflate CSV из ZIP, разделённого на .z01… части.

        Стандартная библиотека Python не умеет читать split-ZIP. Здесь не
        собирается новый архив: поток сжатых байтов читается последовательно
        из частей и сразу распаковывается в UTF-8 CSV.
        """
        with open(parts[0], "rb") as first:
            prefix = first.read(4)
            if prefix == b"PK\x07\x08": prefix = first.read(4)
            if prefix != b"PK\x03\x04": raise ValueError("Первая часть не содержит ZIP-запись CSV.")
            header_rest = first.read(26)
            if len(header_rest) != 26: raise ValueError("Повреждён заголовок ZIP-записи.")
            _, flags, compression, _, _, _, compressed_size, _, name_size, extra_size = struct.unpack("<HHHHHIIIHH", header_rest)
            if compression != 8: raise ValueError("Ожидался ZIP CSV со сжатием Deflate.")
            first.read(name_size + extra_size)
            initial = first.read()
        remaining = int(compressed_size)
        decompressor = zlib.decompressobj(-zlib.MAX_WBITS)
        decoder = codecs.getincrementaldecoder("cp1251")("replace")
        with open(output_path, "w", encoding="utf-8", newline="") as output:
            def consume(chunk):
                nonlocal remaining
                if remaining <= 0: return
                data = chunk[:remaining]; remaining -= len(data)
                text = decoder.decode(decompressor.decompress(data))
                if text: output.write(text)
            consume(initial)
            for part in parts[1:]:
                if remaining <= 0: break
                with open(part, "rb") as source:
                    while remaining > 0:
                        chunk = source.read(min(1024 * 1024, remaining))
                        if not chunk: break
                        consume(chunk)
            tail = decoder.decode(decompressor.flush(), final=True)
            if tail: output.write(tail)
        if remaining != 0: raise ValueError("Архив обрывается до конца CSV-файла.")
    def run(self):
        try:
            t0 = time.time(); archive = os.path.abspath(self.archive_path); base, ext = os.path.splitext(archive)
            self.progress.emit(5, "Проверка архива инвестора…")
            with tempfile.TemporaryDirectory(prefix="registry_investor_") as temp_dir:
                parts = sorted([f"{base}.z{i:02d}" for i in range(1, 100) if os.path.exists(f"{base}.z{i:02d}")])
                utf8_csv = os.path.join(temp_dir, "investor.csv")
                if parts:
                    self.progress.emit(15, "Собираем части многотомного архива…")
                    self.progress.emit(30, "Распаковываем CSV отчёта инвестора…")
                    self._extract_split_zip_csv(parts + [archive], utf8_csv)
                else:
                    self.progress.emit(30, "Распаковываем CSV отчёта инвестора…")
                    with zipfile.ZipFile(archive) as zf:
                        members = [name for name in zf.namelist() if name.lower().endswith(".csv")]
                        if not members: raise ValueError("В архиве не найден CSV-файл.")
                        with zf.open(members[0]) as binary, io.TextIOWrapper(binary, encoding="cp1251", errors="replace", newline="") as reader, open(utf8_csv, "w", encoding="utf-8", newline="") as writer:
                            shutil.copyfileobj(reader, writer, 1024 * 1024)
                self.progress.emit(60, "Формируем базу сверки…")
                db = duckdb.connect(self.output_db)
                safe_csv = utf8_csv.replace("'", "''")
                db.execute("DROP TABLE IF EXISTS investor_data")
                db.execute(f"""CREATE TABLE investor_data AS
                    SELECT * FROM read_csv('{safe_csv}', delim=';', header=true, all_varchar=true,
                                           ignore_errors=true, nullstr=['', 'NULL', 'null'])""")
                db.execute('CREATE INDEX IF NOT EXISTS idx_investor_rrn ON investor_data("RRN")')
                db.execute('CREATE INDEX IF NOT EXISTS idx_investor_transaction ON investor_data("transaction_id")')
                count = db.execute("SELECT COUNT(*) FROM investor_data").fetchone()[0]
                db.execute("CHECKPOINT"); db.close()
            self.progress.emit(100, "Отчёт инвестора готов")
            self.finished.emit(int(count or 0), self.output_db, time.time() - t0)
        except Exception as e:
            import traceback; traceback.print_exc(); self.error.emit(f"{type(e).__name__}: {str(e)}")

class StatsWorker(QThread):
    finished = Signal(dict)
    def __init__(self, repo, where, parent=None):
        super().__init__(parent); self.repo, self.where = repo, where
    def run(self):
        try:
            # Отдельный cursor DuckDB безопасен для фонового запроса, но использует ту же БД.
            # Повторное connect() здесь блокируется основным подключением на macOS.
            worker_repo = DuckDBRepository()
            worker_repo._conn = self.repo._conn.cursor()
            worker_repo._columns = list(self.repo._columns)
            worker_repo._sum_columns = dict(self.repo._sum_columns)
            worker_repo._table_name = self.repo._table_name
            result = worker_repo.get_statistics(self.where)
            worker_repo.close()
            self.finished.emit(result)
        except: self.finished.emit({"transactions": 0})

class UpdateWorker(QThread):
    completed = Signal(dict)
    error = Signal(str)
    def run(self):
        manifest_url = get_update_manifest_url()
        if not manifest_url:
            self.error.emit("Обновления ещё не подключены. Нужен адрес GitHub-манифеста.")
            return
        try:
            with urllib.request.urlopen(manifest_url, timeout=8) as response:
                self.completed.emit(json.loads(response.read().decode("utf-8")))
        except Exception as e:
            self.error.emit(str(e))

class ScriptUpdateWorker(QThread):
    """Безопасно заменяет только исходный .py и сохраняет прошлую версию."""
    completed = Signal(str)
    error = Signal(str)
    def __init__(self, script_url, expected_hash="", parent=None):
        super().__init__(parent); self.script_url = script_url; self.expected_hash = expected_hash.strip().lower()
        # При запуске из .app sys.argv[0] указывает на нативный launcher,
        # а обновлять нужно сам сценарий приложения в Resources.
        self.target_path = os.path.abspath(__file__)
    def run(self):
        temporary_path = self.target_path + ".download"
        backup_path = self.target_path + ".backup"
        try:
            request = urllib.request.Request(self.script_url, headers={"User-Agent": "TransactionAnalyzer-Updater"})
            with urllib.request.urlopen(request, timeout=30) as response:
                payload = response.read()
            if not payload:
                raise ValueError("Сервер вернул пустой файл обновления.")
            digest = hashlib.sha256(payload).hexdigest()
            if self.expected_hash and digest != self.expected_hash:
                raise ValueError("Контрольная сумма обновления не совпадает.")
            source = payload.decode("utf-8-sig")
            compile(source, self.target_path, "exec")
            with open(temporary_path, "wb") as temporary_file:
                temporary_file.write(payload)
            shutil.copy2(self.target_path, backup_path)
            os.replace(temporary_path, self.target_path)
            self.completed.emit(backup_path)
        except Exception as e:
            try:
                if os.path.exists(temporary_path): os.remove(temporary_path)
            except OSError:
                pass
            self.error.emit(str(e))

def make_icon(svg, palette):
    color = palette.color(QPalette.ButtonText).name()
    data = svg.replace("CURRENT", color)
    pm = QPixmap(Theme.Toolbar.icon_size, Theme.Toolbar.icon_size); pm.fill(Qt.transparent)
    r = QSvgRenderer(data.encode("utf-8")); p = QPainter(pm); r.render(p); p.end()
    return QIcon(pm)

class LoadingOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent); self.setVisible(False); self.setStyleSheet("background-color: rgba(15, 23, 42, 72);")
        lay = QVBoxLayout(self); lay.setAlignment(Qt.AlignCenter)
        c = QWidget(); c.setFixedWidth(360); c.setObjectName("loadingCard")
        c.setStyleSheet("QWidget#loadingCard { background-color: #FFFFFF; border: 1px solid #D8DEE9; border-radius: 16px; }")
        cl = QVBoxLayout(c); cl.setContentsMargins(20, 20, 20, 20); cl.setSpacing(12)
        self.progress = QProgressBar(); self.progress.setRange(0, 100); self.progress.setTextVisible(True); self.progress.setFormat("%p%"); self.progress.setFixedHeight(24)
        self.progress.setStyleSheet("QProgressBar { color: #1D1D1F; background: #EEF2F7; border: none; border-radius: 8px; text-align: center; font-weight: 700; } QProgressBar::chunk { background: #0A84FF; border-radius: 8px; }")
        cl.addWidget(self.progress)
        self.label = QLabel("Загрузка..."); self.label.setAlignment(Qt.AlignCenter)
        cl.addWidget(self.label); lay.addWidget(c)
    def show_loading(self, msg="Загрузка...", custom_stages=False):
        self._custom_stages = custom_stages; self.progress.setValue(0); self.label.setText(msg); self.setVisible(True); self.raise_(); self._upd()
    def update_progress(self, p, m):
        self.progress.setValue(p)
        if getattr(self, "_custom_stages", False):
            self.label.setText(m)
            return
        if p < 25:
            stage = "1 из 4 · Читаем исходный файл"
        elif p < 70:
            stage = "2 из 4 · Формируем таблицы"
        elif p < 95:
            stage = "3 из 4 · Формируем списки и фильтры"
        elif p < 100:
            stage = "4 из 4 · Считаем данные"
        else:
            stage = "Готово · Данные подготовлены"
        self.label.setText(stage)
    def hide_loading(self): self.setVisible(False)
    def _upd(self):
        if self.parent(): self.setGeometry(self.parent().rect())
    def resizeEvent(self, e): self._upd(); super().resizeEvent(e)
    def showEvent(self, e): self._upd(); super().showEvent(e)

class AggregatedTableModel(QAbstractTableModel):
    def __init__(self, repo=None, parent=None):
        super().__init__(parent); self._repo = repo; self._data = []
        self._headers = [CARRIER_COLUMN, "Количество"] + list(SUM_COLUMNS.keys())
        self._where = "1=1"; self._limit = 5000
    def set_where_clause(self, w): self._where = w; self.reload()
    def reload(self):
        if not self._repo: self._data = []
        else:
            df = self._repo.get_aggregated(self._where, self._limit)
            self._data = df.to_dict('records') if df is not None and not df.empty else []
        self.beginResetModel(); self.endResetModel()
    def rowCount(self, parent=None): return len(self._data)
    def columnCount(self, parent=None): return len(self._headers)
    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid(): return None
        row, col = index.row(), index.column()
        if row < 0 or row >= len(self._data): return None
        record = self._data[row]; key = self._headers[col]; val = record.get(key, "")
        if role == Qt.DisplayRole:
            if key == "Количество": return f"{int(val):,}".replace(",", " ")
            if isinstance(val, (int, float)): return f"{val:,.2f}".replace(",", " ") if val != 0 else "0.00"
            return str(val)
        return None
    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole and section < len(self._headers):
            return self._headers[section]
        if orientation == Qt.Horizontal and role == Qt.TextAlignmentRole:
            return Qt.AlignLeft | Qt.AlignVCenter if section == 0 else Qt.AlignRight | Qt.AlignVCenter
        return None
    def sort(self, column, order=Qt.AscendingOrder):
        """Сортирует уже показанную сводку, не меняя запросы и расчёты."""
        if column < 0 or column >= len(self._headers):
            return
        key = self._headers[column]
        def sort_key(record):
            value = record.get(key, "")
            if key == CARRIER_COLUMN:
                return str(value or "").casefold()
            try:
                return float(value or 0)
            except (TypeError, ValueError):
                return 0.0
        self.layoutAboutToBeChanged.emit()
        self._data.sort(key=sort_key, reverse=(order == Qt.DescendingOrder))
        self.layoutChanged.emit()

class FilterSection(QWidget):
    filter_changed = Signal()
    def __init__(self, title, repo=None, parent=None):
        super().__init__(parent)
        self.title = title; self.is_expanded = False
        self._repo = repo; self._col = ""; self._where = "1=1"; self._search = ""
        self._checked_values = set()
        self.is_date_column = title in DATE_COLUMNS
        self._default_min = QDate(2020, 1, 1)
        self._default_max = QDate.currentDate()
        
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        self.setStyleSheet(f"QWidget {{ background-color: {Theme.Colors.card_bg}; border-radius: 8px; }}")
        shadow = QGraphicsDropShadowEffect(self); shadow.setBlurRadius(8); shadow.setOffset(0, 2); shadow.setColor(QColor(0, 0, 0, 30))
        self.setGraphicsEffect(shadow)
        lay = QVBoxLayout(self); lay.setContentsMargins(0, 0, 0, 0); lay.setSpacing(0)
        
        hdr = QWidget(); hdr.setFixedHeight(36); hdr.setStyleSheet("background: transparent;")
        hl = QHBoxLayout(hdr); hl.setContentsMargins(8, 0, 8, 0); hl.setSpacing(6)
        self.disclosure_btn = QPushButton("▸"); self.disclosure_btn.setFixedSize(14, 14); self.disclosure_btn.setFlat(True)
        self.disclosure_btn.setCursor(Qt.PointingHandCursor); self.disclosure_btn.setStyleSheet("QPushButton { font-size: 10px; padding: 0; color: #6B7280; }")
        self.disclosure_btn.clicked.connect(self.toggle); hl.addWidget(self.disclosure_btn)
        tl = QLabel(title); tl.setFont(QFont(QFontDatabase.systemFont(QFontDatabase.GeneralFont).family(), Theme.Typography.Size.body, Theme.Typography.Weight.medium))
        tl.setStyleSheet(f"color: {Theme.Colors.text_primary};"); hl.addWidget(tl); hl.addStretch()
        self.badge = QLabel("0"); self.badge.setAlignment(Qt.AlignCenter); self.badge.setFixedHeight(20); self.badge.setMinimumWidth(24)
        self.badge.setStyleSheet("QLabel { background-color: #E5E7EB; color: #6B7280; border-radius: 10px; padding: 0 6px; font-size: 11px; font-weight: 600; }")
        hl.addWidget(self.badge); lay.addWidget(hdr)
        
        self.content = QWidget(); self.content.setVisible(False)
        cl = QVBoxLayout(self.content); cl.setContentsMargins(8, 4, 8, 8); cl.setSpacing(4)
        
        self.search = QLineEdit(); self.search.setPlaceholderText("Поиск значений…"); self.search.setClearButtonEnabled(True); self.search.setFixedHeight(30)
        self.search.setStyleSheet("QLineEdit { color: #1D1D1F; background: #F2F4F7; border: 1px solid #D8DEE9; border-radius: 8px; padding: 3px 9px; font-size: 12px; } QLineEdit:focus { background: #FFFFFF; border-color: #0A84FF; } QLineEdit::placeholder { color: #7A8494; }")
        self._st = QTimer(); self._st.setSingleShot(True); self._st.setInterval(200)
        self._st.timeout.connect(self._do_search)
        self.search.textChanged.connect(lambda: self._st.start()); cl.addWidget(self.search)
        
        if self.is_date_column:
            self.search.setVisible(False)
            
            date_frame = QWidget()
            date_frame.setStyleSheet("background: transparent;")
            date_layout = QHBoxLayout(date_frame)
            date_layout.setContentsMargins(0, 4, 0, 4)
            date_layout.setSpacing(6)
            
            lbl_from = QLabel("От:")
            lbl_from.setStyleSheet("font-size: 12px; color: #374151; font-weight: 500;")
            date_layout.addWidget(lbl_from)
            
            self.date_from = QDateEdit()
            self.date_from.setCalendarPopup(True)
            self.date_from.setDisplayFormat("dd.MM.yyyy")
            self.date_from.setDate(self._default_min)
            self.date_from.setStyleSheet(DATE_EDIT_STYLE)
            setup_russian_calendar(self.date_from)
            self.date_from.dateChanged.connect(self._on_date_range_changed)
            date_layout.addWidget(self.date_from)
            
            lbl_to = QLabel("До:")
            lbl_to.setStyleSheet("font-size: 12px; color: #374151; font-weight: 500;")
            date_layout.addWidget(lbl_to)
            
            self.date_to = QDateEdit()
            self.date_to.setCalendarPopup(True)
            self.date_to.setDisplayFormat("dd.MM.yyyy")
            self.date_to.setDate(self._default_max)
            self.date_to.setStyleSheet(DATE_EDIT_STYLE)
            setup_russian_calendar(self.date_to)
            self.date_to.dateChanged.connect(self._on_date_range_changed)
            date_layout.addWidget(self.date_to)
            
            date_layout.addStretch()
            cl.insertWidget(1, date_frame)
        else:
            self.null_cb = QCheckBox("(Пусто)")
            self.null_cb.setStyleSheet("QCheckBox { padding-left: 8px; font-size: 12px; color: #1D1D1F; margin: 2px 0; } QCheckBox::indicator { width: 14px; height: 14px; border: 1px solid #C7C7CC; border-radius: 4px; background: #FFFFFF; } QCheckBox::indicator:checked { background: #0A84FF; border-color: #0A84FF; }")
            self.null_cb.stateChanged.connect(self._on_null_changed)
            cl.addWidget(self.null_cb)
            
            self.list_widget = QListWidget()
            self.list_widget.setSelectionMode(QAbstractItemView.MultiSelection)
            self.list_widget.setFrameShape(QFrame.NoFrame)
            self.list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.list_widget.setTextElideMode(Qt.ElideRight)
            self.list_widget.setMaximumHeight(220)
            self.list_widget.setStyleSheet("QListWidget { font-size: 12px; color: #1D1D1F; background: transparent; border: none; } QListWidget::item { color: #1D1D1F; padding: 5px 4px; } QListWidget::item:hover { background: #F2F6FC; border-radius: 5px; } QListWidget::indicator { width: 14px; height: 14px; border: 1px solid #C7C7CC; border-radius: 4px; background: #FFFFFF; } QListWidget::indicator:checked { background: #0A84FF; border-color: #0A84FF; }")
            self.list_widget.itemChanged.connect(self._on_item_changed)
            cl.addWidget(self.list_widget)
            
        lay.addWidget(self.content)
        self.setFixedHeight(36)

    def set_date_range_from_data(self, min_date, max_date):
        """Устанавливает реальные min/max из БД как дефолтные значения."""
        self._default_min = min_date
        self._default_max = max_date
        self.date_from.blockSignals(True)
        self.date_to.blockSignals(True)
        self.date_from.setDate(min_date)
        self.date_to.setDate(max_date)
        self.date_from.setMinimumDate(min_date.addYears(-1))
        self.date_to.setMaximumDate(max_date.addYears(1))
        self.date_from.blockSignals(False)
        self.date_to.blockSignals(False)
        self._upd_badge()

    def _do_search(self):
        self._search = self.search.text()
        self._load_data()

    def _on_null_changed(self, state):
        self._upd_badge(); self.filter_changed.emit()

    def _on_date_range_changed(self):
        self._upd_badge(); self.filter_changed.emit()

    def _reset_dates(self):
        self.date_from.blockSignals(True)
        self.date_to.blockSignals(True)
        self.date_from.setDate(self._default_min)
        self.date_to.setDate(self._default_max)
        self.date_from.blockSignals(False)
        self.date_to.blockSignals(False)
        self._upd_badge()
        self.filter_changed.emit()

    def set_column(self, col):
        self._col = col
        self._load_data()

    def _load_data(self):
        if self.is_date_column:
            self._upd_badge()
            return
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        if self._repo:
            data = self._repo.get_value_counts(self._col, self._where, self._search, limit=500, offset=0)
            for val, count in data:
                item = QListWidgetItem(f"{val}  ({count:,})")
                item.setData(Qt.UserRole, val)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked if val in self._checked_values else Qt.Unchecked)
                self.list_widget.addItem(item)
        self.list_widget.blockSignals(False)
        self._upd_badge()

    def _on_item_changed(self, item):
        value = item.data(Qt.UserRole)
        if item.checkState() == Qt.Checked:
            self._checked_values.add(value)
        else:
            self._checked_values.discard(value)
        self._upd_badge(); self.filter_changed.emit()

    def toggle(self):
        self.is_expanded = not self.is_expanded
        self.disclosure_btn.setText("▾" if self.is_expanded else "▸")
        self.content.setVisible(self.is_expanded)
        if self.is_expanded:
            self.setMinimumHeight(0); self.setMaximumHeight(16777215)
            self.adjustSize()
        else:
            self.setFixedHeight(36)
        self.updateGeometry()
        parent = self.parentWidget()
        if parent and parent.layout():
            parent.layout().invalidate()
            parent.layout().activate()
            parent.adjustSize()

    def _is_default_range(self):
        return (self.date_from.date() == self._default_min and 
                self.date_to.date() == self._default_max)

    def _upd_badge(self):
        if self.is_date_column:
            n = 0 if self._is_default_range() else 1
        else:
            n = len(self.get_selected()) + (1 if self.null_cb.isChecked() else 0)
        self.badge.setText(str(n))
        if n > 0: self.badge.setStyleSheet("QLabel { background-color: #3B82F6; color: white; border-radius: 10px; padding: 0 6px; font-size: 11px; font-weight: 600; }")
        else: self.badge.setStyleSheet("QLabel { background-color: #E5E7EB; color: #6B7280; border-radius: 10px; padding: 0 6px; font-size: 11px; font-weight: 600; }")

    def get_selected(self):
        if self.is_date_column: return []
        return list(self._checked_values)

    def include_null(self): 
        if self.is_date_column: return False
        return self.null_cb.isChecked()

    def get_date_range(self):
        if self.is_date_column and hasattr(self, 'date_from') and hasattr(self, 'date_to'):
            return {
                'from': self.date_from.date().toString("yyyy-MM-dd"),
                'to': self.date_to.date().toString("yyyy-MM-dd")
            }
        return None

    def update_counts(self, where_clause):
        if self.is_date_column:
            self._where = where_clause
            self._upd_badge()
            return
        old_selected = set(self._checked_values)
        old_null = self.null_cb.isChecked()
        self._where = where_clause
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        if self._repo:
            data = self._repo.get_value_counts(self._col, self._where, self._search, limit=500, offset=0)
            for val, count in data:
                item = QListWidgetItem(f"{val}  ({count:,})")
                item.setData(Qt.UserRole, val)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked if val in old_selected else Qt.Unchecked)
                self.list_widget.addItem(item)
        self.list_widget.blockSignals(False)
        self.null_cb.blockSignals(True)
        self.null_cb.setChecked(old_null)
        self.null_cb.blockSignals(False)
        self._upd_badge()

    def deselect_value(self, value):
        if self.is_date_column: return
        self._checked_values.discard(value)
        self.list_widget.blockSignals(True)
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(Qt.UserRole) == value:
                item.setCheckState(Qt.Unchecked); break
        self.list_widget.blockSignals(False)
        self._upd_badge()

    def deselect_all(self):
        if self.is_date_column:
            self._reset_dates()
            return
        self._checked_values.clear()
        self.list_widget.blockSignals(True)
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(Qt.Unchecked)
        self.null_cb.blockSignals(True); self.null_cb.setChecked(False); self.null_cb.blockSignals(False)
        self.list_widget.blockSignals(False)
        self._upd_badge()

class SelectedTagsWidget(QWidget):
    tag_removed = Signal(str, str); clear_all_requested = Signal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("selectedTagsCard"); self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(f"QWidget#selectedTagsCard {{ background-color: {Theme.Colors.card_bg}; border-radius: 10px; border: 1px solid {Theme.Colors.border}; }}")
        layout = QVBoxLayout(self); layout.setContentsMargins(12, 8, 12, 8); layout.setSpacing(8)
        header = QHBoxLayout(); header.setSpacing(8)
        title = QLabel("ВЫБРАННЫЕ ФИЛЬТРЫ"); title.setFont(QFont(QFontDatabase.systemFont(QFontDatabase.GeneralFont).family(), Theme.Typography.Size.body, Theme.Typography.Weight.bold))
        title.setStyleSheet(f"color: {Theme.Colors.text_primary};"); header.addWidget(title); header.addStretch()
        self.clear_btn = QPushButton("Очистить все"); self.clear_btn.setFlat(True)
        self.clear_btn.setStyleSheet("QPushButton { color: #0A84FF; background: transparent; border: none; font-size: 12px; font-weight: 600; padding: 0; } QPushButton:hover { color: #0077ED; }")
        self.clear_btn.clicked.connect(self.clear_all_requested.emit); header.addWidget(self.clear_btn); layout.addLayout(header)
        self.tags_layout = QHBoxLayout(); self.tags_layout.setSpacing(6); self.tags_layout.setAlignment(Qt.AlignLeft)
        self.tags_widget = QWidget(); self.tags_widget.setLayout(self.tags_layout); layout.addWidget(self.tags_widget)
        self.tags = []
    def update_tags(self, filters_dict, date_filters_dict):
        while self.tags_layout.count():
            item = self.tags_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self.tags.clear()
        for col, values in filters_dict.items():
            for val in values:
                if val == "(Пусто)": continue
                tag = self._create_tag(col, val); self.tags_layout.addWidget(tag); self.tags.append((col, val, tag))
        for col, date_range in date_filters_dict.items():
            if date_range:
                tag_text = f"{date_range['from']} — {date_range['to']}"
                tag = self._create_tag(col, tag_text); self.tags_layout.addWidget(tag); self.tags.append((col, tag_text, tag))
        self.tags_layout.addStretch()
        self.setVisible(bool(self.tags))
    def _create_tag(self, col, val):
        widget = QWidget(); widget.setStyleSheet("background-color: #E5E7EB; border-radius: 12px; padding: 2px 6px;")
        layout = QHBoxLayout(widget); layout.setContentsMargins(6, 2, 6, 2); layout.setSpacing(4)
        label = QLabel(f"{val}"); label.setStyleSheet("background: transparent; font-size: 12px;")
        close_btn = QPushButton("×"); close_btn.setFixedSize(16, 16); close_btn.setFlat(True); close_btn.setStyleSheet("color: #6B7280; font-size: 14px;")
        close_btn.clicked.connect(lambda: self.tag_removed.emit(col, val))
        layout.addWidget(label); layout.addWidget(close_btn); return widget

class SidebarFilters(QWidget):
    filters_changed = Signal()
    def __init__(self, repo=None, parent=None):
        super().__init__(parent); self.setMinimumWidth(Theme.Sidebar.min_width); self.setMaximumWidth(Theme.Sidebar.max_width)
        self._repo = repo; self.setStyleSheet(f"background-color: {Theme.Colors.bg};")
        ml = QVBoxLayout(self); ml.setContentsMargins(12, 12, 12, 12); ml.setSpacing(8)
        hdr = QHBoxLayout(); title = QLabel("ФИЛЬТРЫ"); title.setFont(QFont(QFontDatabase.systemFont(QFontDatabase.GeneralFont).family(), Theme.Typography.Size.body, Theme.Typography.Weight.bold))
        title.setStyleSheet(f"color: {Theme.Colors.text_primary};"); hdr.addWidget(title); hdr.addStretch(); ml.addLayout(hdr)
        sep = QFrame(); sep.setFrameShape(QFrame.HLine); sep.setStyleSheet(f"background-color: {Theme.Colors.border};"); sep.setFixedHeight(1); ml.addWidget(sep)
        scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.NoFrame); scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        fw = QWidget(); fw.setStyleSheet("background: transparent;"); self.filters_lay = QVBoxLayout(fw); self.filters_lay.setContentsMargins(0, 0, 0, 0); self.filters_lay.setSpacing(6)
        self.filters_lay.addStretch(1)
        scroll.setWidget(fw); ml.addWidget(scroll); self.sections = {}
        self.selected_tags = SelectedTagsWidget(); self.selected_tags.tag_removed.connect(self._on_tag_removed); self.selected_tags.clear_all_requested.connect(self._on_clear_all)
        ml.addWidget(self.selected_tags)
        self.status_label = QLabel("Готово"); self.status_label.setFont(QFont(QFontDatabase.systemFont(QFontDatabase.GeneralFont).family(), Theme.Typography.Size.caption))
        self.status_label.setStyleSheet(f"color: {Theme.Colors.text_secondary};"); ml.addWidget(self.status_label)
    def add_filter_section(self, title, column=None):
        if title in self.sections: return
        col = column if column is not None else title
        s = FilterSection(title, self._repo); s.set_column(col); s.filter_changed.connect(self._on_filter_changed)
        self.filters_lay.insertWidget(self.filters_lay.count() - 1, s); self.sections[title] = (s, col)
    def clear_filters(self):
        for s, _ in self.sections.values(): s.setParent(None); s.deleteLater()
        self.sections.clear()
    def _on_filter_changed(self):
        self.update_all()
        self.filters_changed.emit()
    def _on_tag_removed(self, col, value):
        for title, (section, column) in self.sections.items():
            if column == col:
                if section.is_date_column:
                    section._reset_dates()
                else:
                    section.deselect_value(value)
                break
        self.update_all()
        self.filters_changed.emit()
    def _on_clear_all(self):
        for _, (section, _) in self.sections.items(): section.deselect_all()
        self.update_all()
        self.filters_changed.emit()
    def get_filters(self):
        f = {}; date_filters = {}
        for title, (section, column) in self.sections.items():
            if section.is_date_column:
                if not section._is_default_range():
                    date_filters[column] = section.get_date_range()
            else:
                sel = section.get_selected()
                if sel or section.include_null():
                    f[column] = sel
                    if section.include_null(): f[column].append("(Пусто)")
        return f, date_filters
    def update_all(self):
        filters, date_filters = self.get_filters()
        for title, (section, col) in self.sections.items():
            where = self._repo.build_where_clause(filters, date_filters, exclude_column=col) if self._repo else "1=1"
            section.update_counts(where)
        self.selected_tags.update_tags(filters, date_filters); self.status_label.setText("Готово")
    def set_repo(self, repo):
        self._repo = repo
        for _, (section, col) in self.sections.items():
            section._repo = repo; section.set_column(col)
            # Устанавливаем реальные даты из данных
            if section.is_date_column and repo:
                dr = repo.get_date_range(col)
                if dr:
                    section.set_date_range_from_data(dr[0], dr[1])
        self.update_all()

class SummaryCard(QWidget):
    def __init__(self, icon_color, icon_text, title, subtitle, parent=None):
        super().__init__(parent)
        self.setObjectName("summaryCard")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setStyleSheet(f"QWidget#summaryCard {{ background-color: {Theme.Colors.card_bg}; border: 1px solid {Theme.Colors.border}; border-radius: 16px; }}")
        self.setFixedHeight(112); lay = QVBoxLayout(self); lay.setContentsMargins(15, 13, 15, 13); lay.setSpacing(4)
        top = QHBoxLayout(); top.setSpacing(8)
        icon_lbl = QLabel(icon_text); icon_lbl.setFixedSize(28, 28); icon_lbl.setAlignment(Qt.AlignCenter)
        icon_lbl.setStyleSheet(f"QLabel {{ background-color: {icon_color}; border-radius: 9px; color: white; font-size: 15px; font-weight: bold; }}")
        top.addWidget(icon_lbl)
        self.title_lbl = QLabel(title); self.title_lbl.setFont(QFont(QFontDatabase.systemFont(QFontDatabase.GeneralFont).family(), Theme.Typography.Size.body, Theme.Typography.Weight.medium))
        self.title_lbl.setStyleSheet(f"color: {Theme.Colors.text_primary};"); top.addWidget(self.title_lbl); top.addStretch(); lay.addLayout(top)
        self.value_lbl = QLabel("0"); self.value_lbl.setMinimumWidth(0); self.value_lbl.setFont(QFont(QFontDatabase.systemFont(QFontDatabase.GeneralFont).family(), 19, Theme.Typography.Weight.bold))
        self.value_lbl.setStyleSheet(f"color: {Theme.Colors.text_primary};"); lay.addWidget(self.value_lbl)
        self.subtitle_lbl = QLabel(subtitle); self.subtitle_lbl.setFont(QFont(QFontDatabase.systemFont(QFontDatabase.GeneralFont).family(), Theme.Typography.Size.caption))
        self.subtitle_lbl.setStyleSheet(f"color: {Theme.Colors.text_secondary};"); lay.addWidget(self.subtitle_lbl)
    def resizeEvent(self, event):
        # Числа не обрезаются при изменении ширины окна.
        size = 16 if self.width() >= 220 else (14 if self.width() >= 180 else 12)
        font = self.value_lbl.font(); font.setPointSize(size); self.value_lbl.setFont(font)
        super().resizeEvent(event)

class DashboardHeader(QWidget):
    """Информационная шапка экрана; не участвует в расчётах."""
    reset_requested = Signal()
    charts_requested = Signal()
    fleet_requested = Signal()
    operations_requested = Signal()
    fraud_requested = Signal()
    investor_requested = Signal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dashboardHeader")
        lay = QHBoxLayout(self); lay.setContentsMargins(4, 4, 4, 12); lay.setSpacing(12)
        title_box = QVBoxLayout(); title_box.setSpacing(2)
        title = QLabel("Реестр транзакций")
        title.setStyleSheet(f"color: {Theme.Colors.text_primary}; font-size: 24px; font-weight: 700; letter-spacing: 0.2px;")
        subtitle = QLabel("Сводка по загруженной базе и выбранным фильтрам")
        subtitle.setStyleSheet(f"color: {Theme.Colors.text_secondary}; font-size: 13px;")
        title_box.addWidget(title); title_box.addWidget(subtitle); lay.addLayout(title_box); lay.addStretch()
        self.fleet_btn = QPushButton("Парк")
        self.fleet_btn.setCursor(Qt.PointingHandCursor); self.fleet_btn.setFixedHeight(34)
        self.fleet_btn.setStyleSheet("QPushButton { background: #FFFFFF; color: #1D1D1F; border: 1px solid #D8DEE9; border-radius: 9px; padding: 0 14px; font-weight: 700; } QPushButton:hover { background: #F3F7FF; }")
        self.fleet_btn.clicked.connect(self.fleet_requested.emit); lay.addWidget(self.fleet_btn)
        self.operations_btn = QPushButton("Эффективность")
        self.operations_btn.setCursor(Qt.PointingHandCursor); self.operations_btn.setFixedHeight(34)
        self.operations_btn.setStyleSheet("QPushButton { background: #FFFFFF; color: #1D1D1F; border: 1px solid #D8DEE9; border-radius: 9px; padding: 0 14px; font-weight: 700; } QPushButton:hover { background: #F3F7FF; }")
        self.operations_btn.clicked.connect(self.operations_requested.emit); lay.addWidget(self.operations_btn)
        self.fraud_btn = QPushButton("ФРОД")
        self.fraud_btn.setCursor(Qt.PointingHandCursor); self.fraud_btn.setFixedHeight(34)
        self.fraud_btn.setToolTip("Безналичные операции по маскам обслуженных карт")
        self.fraud_btn.setStyleSheet("QPushButton { background: #FFFFFF; color: #1D1D1F; border: 1px solid #D8DEE9; border-radius: 9px; padding: 0 14px; font-weight: 700; } QPushButton:hover { background: #FFF7ED; border-color:#FDBA74; }")
        self.fraud_btn.clicked.connect(self.fraud_requested.emit); lay.addWidget(self.fraud_btn)
        self.investor_btn = QPushButton("Сверка с ВТБ")
        self.investor_btn.setText("ВТБ")
        self.investor_btn.setCursor(Qt.PointingHandCursor); self.investor_btn.setFixedHeight(34)
        self.investor_btn.setToolTip("Сопоставить безналичные операции АСОП с отчётом ВТБ")
        self.investor_btn.setStyleSheet("QPushButton { background: #FFFFFF; color: #1D1D1F; border: 1px solid #D8DEE9; border-radius: 9px; padding: 0 14px; font-weight: 700; } QPushButton:hover { background: #F2EDFF; border-color:#C4B5FD; }")
        self.investor_btn.clicked.connect(self.investor_requested.emit); lay.addWidget(self.investor_btn)
        self.charts_btn = QPushButton("Графики")
        self.charts_btn.setCursor(Qt.PointingHandCursor); self.charts_btn.setFixedHeight(34)
        self.charts_btn.setStyleSheet("QPushButton { background: #0A84FF; color: white; border: none; border-radius: 9px; padding: 0 14px; font-weight: 700; } QPushButton:hover { background: #0077ED; }")
        self.charts_btn.clicked.connect(self.charts_requested.emit); lay.addWidget(self.charts_btn)
        self.reset_btn = QPushButton("Сброс")
        self.reset_btn.setToolTip("Сбросить фильтры")
        self.reset_btn.setCursor(Qt.PointingHandCursor); self.reset_btn.setFixedHeight(34)
        self.reset_btn.setStyleSheet("QPushButton { color: #1D1D1F; background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 9px; padding: 0 14px; font-weight: 600; } QPushButton:hover { color: #1D1D1F; background: #F3F7FF; border-color: #BFDBFE; }")
        self.reset_btn.clicked.connect(self.reset_requested.emit); lay.addWidget(self.reset_btn)

class SummaryBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(128)
        self.setStyleSheet("background: transparent;")
        lay = QHBoxLayout(self); lay.setContentsMargins(0, 0, 0, 0); lay.setSpacing(14)
        self.cards = {}
        for color, icon, key, title, subtitle in (
            (Theme.Colors.blue, "▤", "transactions", "Транзакции", "Всего записей"),
            (Theme.Colors.green, "₽", "Оплаченная сумма", "Оплачено", "Сумма оплат"),
            (Theme.Colors.purple, "◉", "Применённый тариф", "Тариф", "Сумма тарифов"),
            (Theme.Colors.orange, "◇", "Льгота", "Льготы", "Сумма льгот"),
        ):
            card = SummaryCard(color, icon, title, subtitle)
            lay.addWidget(card, 1); self.cards[key] = card.value_lbl
    def update_stats(self, stats):
        self.cards["transactions"].setText(f"{int(stats.get('transactions', 0)):,}".replace(",", " "))
        for alias in ["Оплаченная сумма", "Применённый тариф", "Льгота"]:
            if alias in self.cards:
                val = stats.get(alias, 0.0)
                amount = f"{val:,.2f}".replace(",", " ").replace(".", ",")
                self.cards[alias].setText(f"{amount} ₽")

class MacProxyStyle(QProxyStyle):
    def drawPrimitive(self, el, opt, p, w=None):
        super().drawPrimitive(el, opt, p, w)
        if el == QStyle.PE_PanelItemViewItem:
            if opt.state & QStyle.State_MouseOver and not (opt.state & QStyle.State_Selected):
                c = opt.palette.color(QPalette.Midlight); c.setAlpha(50); p.fillRect(opt.rect, c)
    def drawControl(self, el, opt, p, w=None):
        if el == QStyle.CE_Header:
            super().drawControl(el, opt, p, w)
            p.save(); p.setPen(QPen(opt.palette.color(QPalette.Mid), 1))
            p.drawLine(opt.rect.bottomLeft(), opt.rect.bottomRight()); p.restore(); return
        super().drawControl(el, opt, p, w)
    def pixelMetric(self, m, o=None, w=None):
        d = {QStyle.PM_HeaderMargin: 8, QStyle.PM_HeaderGripMargin: 4, QStyle.PM_SplitterWidth: 1, QStyle.PM_ToolBarItemMargin: 4, QStyle.PM_ToolBarFrameWidth: 0, QStyle.PM_ButtonMargin: 6, QStyle.PM_DefaultFrameWidth: 0, QStyle.PM_ToolBarItemSpacing: 8}
        return d.get(m, super().pixelMetric(m, o, w))
    def styleHint(self, h, o=None, w=None, rd=None):
        if h == QStyle.SH_ItemView_ShowDecorationSelected: return True
        return super().styleHint(h, o, w, rd)

class RowDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent); self._font = QFontDatabase.systemFont(QFontDatabase.GeneralFont)
    def initStyleOption(self, opt, idx):
        super().initStyleOption(opt, idx); opt.font = self._font
        d = idx.data(Qt.DisplayRole)
        if d is not None:
            try:
                cleaned = str(d).replace(' ', '').replace(',', '.')
                float(cleaned); opt.displayAlignment = Qt.AlignRight | Qt.AlignVCenter; return
            except: pass
        opt.displayAlignment = Qt.AlignLeft | Qt.AlignVCenter

class ResultsTableWidget(QWidget):
    def __init__(self, repo=None, parent=None):
        super().__init__(parent); self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        lay = QVBoxLayout(self); lay.setContentsMargins(0, 0, 0, 0); lay.setSpacing(0)
        self.tv = QTableView(); self.tv.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tv.setSelectionBehavior(QAbstractItemView.SelectRows); self.tv.setShowGrid(False); self.tv.setAlternatingRowColors(True)
        self.tv.setSortingEnabled(True)
        self.tv.verticalHeader().setVisible(False); self.tv.horizontalHeader().setStretchLastSection(False)
        self.tv.horizontalHeader().setSortIndicatorShown(True)
        header = self.tv.horizontalHeader(); header.setSectionsMovable(True); header.setSectionResizeMode(QHeaderView.Interactive)
        self.tv.setItemDelegate(RowDelegate(self.tv))
        self.model = AggregatedTableModel(repo); self.tv.setModel(self.model)
        for column, width in enumerate((280, 130, 175, 175, 140, 175)):
            self.tv.setColumnWidth(column, width)
        self.tv.verticalHeader().setDefaultSectionSize(Theme.Table.row_height); self.tv.horizontalHeader().setFixedHeight(Theme.Table.header_height)
        lay.addWidget(self.tv)
    def update_info_label(self): pass

class DistributionCanvas(QWidget):
    """Линейный график оплат по дням текущей выборки."""
    def __init__(self, parent=None):
        super().__init__(parent); self.rows = []; self.zoom = 1; self.metric_name = "Оплачено"; self.currency = True; self._points = []; self._visible_rows = []; self.setMinimumHeight(150); self.setMouseTracking(True)
    def set_rows(self, rows):
        self.rows = rows or []; self._points = []; self._visible_rows = []; self.update()
    def set_zoom(self, zoom):
        self.zoom = max(1, int(zoom)); self.update()
    def set_metric(self, name, currency=False):
        self.metric_name, self.currency = name, currency
    def paintEvent(self, event):
        painter = QPainter(self); painter.setRenderHint(QPainter.Antialiasing)
        rect = self.rect().adjusted(54, 8, -18, -28)
        if not self.rows:
            painter.setPen(QColor(Theme.Colors.text_secondary)); painter.drawText(rect, Qt.AlignCenter, "Нет данных для визуализации")
            return
        rows = self.rows[-max(2, len(self.rows) // self.zoom):]
        self._visible_rows = rows
        values = [float(value or 0) for _, value in rows]
        max_value = max(values) or 1.0
        painter.setFont(QFont(QFontDatabase.systemFont(QFontDatabase.GeneralFont).family(), 10))
        for i in range(4):
            y = rect.top() + int(rect.height() * i / 3)
            value = max_value * (1 - i / 3)
            painter.setPen(QPen(QColor("#E9EDF4"), 1)); painter.drawLine(rect.left(), y, rect.right(), y)
            painter.setPen(QColor(Theme.Colors.text_secondary)); painter.drawText(0, y - 8, 48, 16, Qt.AlignRight | Qt.AlignVCenter, f"{value / 1000:,.0f}k")
        count = len(values)
        points = []
        for i, value in enumerate(values):
            x = rect.left() if count == 1 else rect.left() + int(rect.width() * i / (count - 1))
            y = rect.bottom() - int(rect.height() * value / max_value)
            points.append((x, y))
        self._points = points
        painter.setPen(QPen(QColor("#0A84FF"), 2.5))
        for i in range(1, len(points)):
            painter.drawLine(points[i - 1][0], points[i - 1][1], points[i][0], points[i][1])
        painter.setBrush(QColor("#0A84FF")); painter.setPen(Qt.NoPen)
        for x, y in points[::max(1, len(points) // 12)]: painter.drawEllipse(x - 3, y - 3, 6, 6)
        painter.setPen(QColor(Theme.Colors.text_secondary))
        for pos in sorted(set([0, max(0, count // 2), max(0, count - 1)])):
            label = rows[pos][0][-5:] if rows[pos] else ""
            x = points[pos][0] - 32
            painter.drawText(x, rect.bottom() + 8, 64, 16, Qt.AlignCenter, label)
        painter.end()
    def mouseMoveEvent(self, event):
        if not self._points or not self._visible_rows:
            return
        pos = event.position()
        index = min(range(len(self._points)), key=lambda i: abs(self._points[i][0] - pos.x()))
        day, amount = self._visible_rows[index]
        value = f"{float(amount):,.2f}".replace(",", " ").replace(".", ",") if self.currency else f"{int(amount):,}".replace(",", " ")
        suffix = " ₽" if self.currency else ""
        QToolTip.showText(event.globalPosition().toPoint(), f"Дата: {day}\n{self.metric_name}: {value}{suffix}", self)
        super().mouseMoveEvent(event)

class ChartWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("chartCard")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(f"QWidget#chartCard {{ background: {Theme.Colors.card_bg}; border: 1px solid {Theme.Colors.border}; border-radius: 16px; }} QPushButton {{ color: {Theme.Colors.text_primary}; font-size: 15px; font-weight: 700; padding: 4px; }} QPushButton:hover {{ color: {Theme.Colors.blue}; }}")
        self.setMaximumHeight(246)
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 6, 20, 6)
        h = QHBoxLayout()
        self.btn = QPushButton("▾ Динамика оплат"); self.btn.setFlat(True)
        self.btn.clicked.connect(self.toggle)
        h.addWidget(self.btn); h.addStretch()
        self.caption = QLabel("По дням в текущей выборке")
        self.caption.setStyleSheet(f"color: {Theme.Colors.text_secondary}; font-size: 11px;")
        h.addWidget(self.caption)
        lay.addLayout(h)
        self.canvas = DistributionCanvas(); lay.addWidget(self.canvas)
        self.collapsed = False
        self._anim = QPropertyAnimation(self, b"maximumHeight")
        self._anim.setDuration(250); self._anim.setEasingCurve(QEasingCurve.InOutQuad)
        self._anim.finished.connect(self._on_fin)
        self._opacity = QGraphicsOpacityEffect(self.canvas); self.canvas.setGraphicsEffect(self._opacity)
        self._fade = QPropertyAnimation(self._opacity, b"opacity", self)
        self._fade.setDuration(220); self._fade.setEasingCurve(QEasingCurve.InOutQuad)
    def set_data(self, rows):
        self.canvas.set_rows(rows)
        self._fade.stop(); self._fade.setStartValue(0.25); self._fade.setEndValue(1.0); self._fade.start()
    def _on_fin(self):
        if self.collapsed: self.canvas.setMaximumHeight(0)
    def toggle(self):
        self.collapsed = not self.collapsed; self._anim.stop()
        if self.collapsed:
            self.btn.setText("▸ Динамика оплат")
            self._anim.setStartValue(self.maximumHeight()); self._anim.setEndValue(46)
        else:
            self.canvas.setMaximumHeight(16777215)
            self.btn.setText("▾ Динамика оплат")
            self._anim.setStartValue(self.maximumHeight()); self._anim.setEndValue(246)
        self._anim.start()
        self.updateGeometry()

class ChartsDialog(QDialog):
    """Отдельный экран графиков для текущей фильтрованной выборки."""
    def __init__(self, rows, parent=None, title="Динамика оплат", caption=None, metric_name="Оплачено", currency=True):
        super().__init__(parent); self.setWindowTitle("Графики"); self.resize(980, 620)
        lay = QVBoxLayout(self); lay.setContentsMargins(24, 20, 24, 20); lay.setSpacing(14)
        title = QLabel(title)
        title.setStyleSheet("font-size: 22px; font-weight: 700; color: #1D1D1F;"); lay.addWidget(title)
        caption = QLabel(caption or "Данные по дням. Наведите курсор на линию, чтобы увидеть дату и значение.")
        caption.setStyleSheet("color: #6B7280;"); lay.addWidget(caption)
        self.canvas = DistributionCanvas(); self.canvas.setMinimumHeight(420); self.canvas.set_metric(metric_name, currency); self.canvas.set_rows(rows); lay.addWidget(self.canvas, 1)
        controls = QHBoxLayout(); label = QLabel("Масштаб по времени")
        self.zoom_label = QLabel("1×")
        self.zoom = QSlider(Qt.Horizontal); self.zoom.setRange(1, 4); self.zoom.setValue(1); self.zoom.setTickPosition(QSlider.TicksBelow); self.zoom.setTickInterval(1)
        self.zoom.valueChanged.connect(self._set_zoom)
        controls.addWidget(label); controls.addWidget(self.zoom, 1); controls.addWidget(self.zoom_label); lay.addLayout(controls)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)
    def _set_zoom(self, value):
        self.zoom_label.setText(f"{value}×")
        self.canvas.set_zoom(value)

class FleetIcon(QWidget):
    """Лёгкая векторная иконка типа транспорта без внешних файлов."""
    COLORS = {"Автобус": "#0A84FF", "Троллейбус": "#5856D6", "Трамвай": "#FF9F0A", "Электропоезд": "#30B0C7", "Маршрут": "#34C759"}
    def __init__(self, vehicle_type, parent=None):
        super().__init__(parent); self.vehicle_type = vehicle_type; self.setFixedSize(34, 34)
    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        color = QColor(self.COLORS.get(self.vehicle_type, "#8E8E93")); p.setPen(Qt.NoPen); p.setBrush(color)
        if self.vehicle_type == "Маршрут":
            p.drawRoundedRect(6, 7, 22, 20, 7, 7); p.setPen(QPen(QColor("white"), 2)); p.drawLine(11, 23, 16, 17); p.drawLine(16, 17, 21, 19)
        else:
            p.drawRoundedRect(5, 10, 24, 16, 5, 5)
            p.setBrush(QColor("white")); p.drawRoundedRect(9, 13, 16, 6, 2, 2)
            p.setBrush(color.darker(125)); p.drawEllipse(9, 23, 5, 5); p.drawEllipse(21, 23, 5, 5)
            if self.vehicle_type == "Троллейбус":
                p.setPen(QPen(color, 2)); p.drawLine(13, 9, 10, 4); p.drawLine(21, 9, 24, 4)
            elif self.vehicle_type == "Трамвай":
                p.setPen(QPen(color, 2)); p.drawLine(17, 9, 13, 4); p.drawLine(13, 4, 21, 4); p.drawLine(21, 4, 17, 9)
            elif self.vehicle_type == "Электропоезд":
                p.setPen(QPen(QColor("white"), 1)); p.drawLine(17, 12, 17, 19)
        p.end()

class RouteVehiclesDialog(QDialog):
    def __init__(self, repo, where_clause, carrier, route, parent=None):
        super().__init__(parent); self.setWindowTitle(f"ТС на маршруте"); self.resize(760, 500)
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel(f"Маршрут {route}"); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        top = QHBoxLayout(); caption = QLabel("Уникальные ТС и число транзакций на маршруте."); caption.setStyleSheet("color:#6B7280;"); top.addWidget(caption); top.addStretch()
        chart = QPushButton("График"); chart.clicked.connect(lambda: ChartsDialog(repo.get_daily_transactions(where_clause, carrier_name=carrier, route_name=route), self, "Транзакции по маршруту", "Количество транзакций по дням для выбранного маршрута.", "Транзакции", False).exec()); top.addWidget(chart); lay.addLayout(top)
        table = QTableWidget(0, 2); table.setHorizontalHeaderLabels(["Номер ТС", "Транзакции"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        for row, (vehicle, count) in enumerate(repo.get_route_vehicles(where_clause, carrier, route)):
            table.insertRow(row); table.setItem(row, 0, QTableWidgetItem(vehicle)); table.setItem(row, 1, QTableWidgetItem(f"{count:,}".replace(",", " ")))
        lay.addWidget(table, 1); buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class VehicleRoutesDialog(QDialog):
    def __init__(self, repo, where_clause, carrier, vehicle, parent=None):
        super().__init__(parent); self.setWindowTitle(f"Маршруты ТС {vehicle}"); self.resize(760, 460)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; gridline-color:#EEF0F4; } QHeaderView::section { color:#6B7280; background:#FFFFFF; border:none; border-bottom:1px solid #D8DEE9; padding:7px; font-weight:600; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel(f"ТС {vehicle}"); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        top = QHBoxLayout(); caption = QLabel(f"Маршруты у перевозчика «{carrier}» и количество транзакций."); caption.setStyleSheet("color:#6B7280;"); top.addWidget(caption); top.addStretch()
        chart = QPushButton("График"); chart.clicked.connect(lambda: ChartsDialog(repo.get_daily_transactions(where_clause, carrier_name=carrier, vehicle_name=vehicle), self, "Транзакции по ТС", "Количество транзакций по дням для выбранного ТС.", "Транзакции", False).exec()); top.addWidget(chart); lay.addLayout(top)
        table = QTableWidget(0, 2); table.setHorizontalHeaderLabels(["Маршрут", "Транзакции"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.horizontalHeader().setStretchLastSection(False); table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        for row, (route, count) in enumerate(repo.get_vehicle_routes(where_clause, carrier, vehicle)):
            table.insertRow(row); table.setItem(row, 0, QTableWidgetItem(route)); table.setItem(row, 1, QTableWidgetItem(f"{count:,}".replace(",", " ")))
        table.cellDoubleClicked.connect(lambda row, _: RouteVehiclesDialog(repo, where_clause, carrier, table.item(row, 0).text(), self).exec())
        lay.addWidget(table, 1); buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class FleetRoutesDialog(QDialog):
    def __init__(self, repo, where_clause, carrier, parent=None):
        super().__init__(parent); self.setWindowTitle(f"Маршруты — {carrier}"); self.resize(820, 560)
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Маршруты перевозчика"); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        caption = QLabel("Дважды нажмите на маршрут, чтобы увидеть ТС на нём."); caption.setStyleSheet("color:#6B7280;"); lay.addWidget(caption)
        table = QTableWidget(0, 2); table.setHorizontalHeaderLabels(["Маршрут", "Транзакции"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        for row, (route, count) in enumerate(repo.get_fleet_routes(where_clause, carrier)):
            table.insertRow(row); table.setItem(row, 0, QTableWidgetItem(route)); table.setItem(row, 1, QTableWidgetItem(f"{count:,}".replace(",", " ")))
        table.cellDoubleClicked.connect(lambda row, _: RouteVehiclesDialog(repo, where_clause, carrier, table.item(row, 0).text(), self).exec())
        lay.addWidget(table, 1); buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class DuplicateVehiclesDialog(QDialog):
    """Показывает номера, из-за которых не сходится уникальный итог по ТС."""
    def __init__(self, repo, where_clause, carrier, parent=None):
        super().__init__(parent); self.setWindowTitle(f"Повторяющиеся номера ТС — {carrier}"); self.resize(820, 520)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; gridline-color:#EEF0F4; } QHeaderView::section { color:#6B7280; background:#FFFFFF; border:none; border-bottom:1px solid #D8DEE9; padding:8px; font-weight:600; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Номера ТС, встречающиеся в нескольких типах"); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        note = QLabel("Один и тот же номер учитывается как одно уникальное ТС, но попадает в каждый указанный тип. Поэтому итог по типам может быть больше уникального итога.")
        note.setWordWrap(True); note.setStyleSheet("color:#6B7280;"); lay.addWidget(note)
        rows = repo.get_duplicate_fleet_vehicles(where_clause, carrier)
        table = QTableWidget(0, 4); table.setHorizontalHeaderLabels(["Номер ТС", "Типы ТС", "Типов", "Транзакции"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setSortingEnabled(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        for col in (2, 3): table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        table.setSortingEnabled(False)
        for row, (vehicle, types, type_count, transactions) in enumerate(rows):
            table.insertRow(row)
            for col, value in enumerate((vehicle, types, str(type_count), f"{transactions:,}".replace(",", " "))):
                item = QTableWidgetItem(value)
                if col >= 2: item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(row, col, item)
        table.setSortingEnabled(True); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class GlobalDuplicateVehiclesDialog(QDialog):
    """Общий список совпадающих номеров без сложения их в «уникальный» итог."""
    def __init__(self, repo, where_clause, carriers=None, parent=None):
        super().__init__(parent); self.setWindowTitle("Повторяющиеся номера ТС"); self.resize(960, 600)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; gridline-color:#EEF0F4; } QHeaderView::section { color:#6B7280; background:#FFFFFF; border:none; border-bottom:1px solid #D8DEE9; padding:8px; font-weight:600; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Повторяющиеся номера ТС"); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        note = QLabel("Показаны только номера, встречающиеся в разных связках «перевозчик × тип ТС». Так, номер «А» у Крымтроллейбуса будет одной строкой с типами «Автобус» и «Троллейбус». Если номер есть у другого перевозчика, он также будет указан.")
        note.setWordWrap(True); note.setStyleSheet("color:#6B7280;"); lay.addWidget(note)
        rows = repo.get_global_duplicate_fleet_vehicles(where_clause, carriers)
        table = QTableWidget(0, 4); table.setHorizontalHeaderLabels(["Номер ТС", "Перевозчик и тип ТС", "Связок", "Транзакции"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setSortingEnabled(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        table.setSortingEnabled(False)
        for row, (vehicle, locations, placements, transactions) in enumerate(rows):
            table.insertRow(row)
            values = (vehicle, locations, str(placements), f"{transactions:,}".replace(",", " "))
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col >= 2: item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(row, col, item)
            table.setRowHeight(row, max(30, 16 * (locations.count("\n") + 1) + 12))
        table.setSortingEnabled(True); table.cellDoubleClicked.connect(lambda row, _: self._open_details(table, row, repo, where_clause, carriers)); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)
    def _open_details(self, table, row, repo, where_clause, carriers):
        item = table.item(row, 0)
        if item: DuplicateVehicleDetailsDialog(repo, where_clause, item.text(), carriers, self).exec()

class DuplicateVehicleDetailsDialog(QDialog):
    """Расшифровка одного повторяющегося номера ТС."""
    def __init__(self, repo, where_clause, vehicle, carriers=None, parent=None):
        super().__init__(parent); self.setWindowTitle(f"ТС {vehicle} — подробности"); self.resize(1120, 650)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; gridline-color:#EEF0F4; } QHeaderView::section { color:#6B7280; background:#FFFFFF; border:none; border-bottom:1px solid #D8DEE9; padding:8px; font-weight:600; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel(f"ТС {vehicle}"); title.setStyleSheet("font-size:21px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        note = QLabel("Все связи этого номера: перевозчик, тип ТС, маршрут и водитель. Данные не объединяются между разными перевозчиками или типами.")
        note.setWordWrap(True); note.setStyleSheet("color:#6B7280;"); lay.addWidget(note)
        rows = repo.get_duplicate_vehicle_details(where_clause, vehicle, carriers)
        total = sum(row["transactions"] for row in rows)
        summary = QLabel(f"{len(rows):,} связок   ·   {total:,} транзакций".replace(",", " "))
        summary.setStyleSheet("background:#FFFFFF; border:1px solid #D8DEE9; border-radius:9px; padding:8px 10px; font-weight:700;"); lay.addWidget(summary)
        headers = ["Перевозчик", "Тип ТС", "Маршрут", "Водитель", "Транзакции"]
        table = QTableWidget(0, len(headers)); table.setHorizontalHeaderLabels(headers); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setSortingEnabled(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        table.setSortingEnabled(False)
        for row_index, item in enumerate(rows):
            table.insertRow(row_index)
            values = [item["carrier"], item["type"], item["route"], item["driver"], f"{item['transactions']:,}".replace(",", " ")]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if col == 4: cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(row_index, col, cell)
        table.setSortingEnabled(True); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class FleetVehiclesDialog(QDialog):
    def __init__(self, repo, where_clause, carrier, parent=None):
        super().__init__(parent); self.repo = repo; self.where_clause = where_clause; self.carrier = carrier; self.setWindowTitle(f"ТС — {carrier}"); self.resize(760, 560)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; gridline-color:#EEF0F4; } QHeaderView::section { color:#6B7280; background:#FFFFFF; border:none; border-bottom:1px solid #D8DEE9; padding:7px; font-weight:600; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Транспортные средства"); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        top = QHBoxLayout(); caption = QLabel("Дважды нажмите на ТС, чтобы открыть маршруты и количество транзакций."); caption.setStyleSheet("color:#6B7280;"); top.addWidget(caption); top.addStretch()
        chart = QPushButton("График"); chart.clicked.connect(lambda: ChartsDialog(repo.get_daily_transactions(where_clause, carrier_name=carrier), self, "Транзакции перевозчика", "Количество транзакций по дням для выбранного перевозчика.", "Транзакции", False).exec()); top.addWidget(chart); lay.addLayout(top)
        self.table = QTableWidget(0, 2); self.table.setHorizontalHeaderLabels(["Номер ТС", "Транзакции"]); self.table.verticalHeader().setVisible(False); self.table.setEditTriggers(QAbstractItemView.NoEditTriggers); self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setStretchLastSection(False); self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch); self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        for row, (vehicle, count) in enumerate(repo.get_fleet_vehicles(where_clause, carrier)):
            self.table.insertRow(row); item = QTableWidgetItem(vehicle); item.setData(Qt.UserRole, vehicle); self.table.setItem(row, 0, item); self.table.setItem(row, 1, QTableWidgetItem(f"{count:,}".replace(",", " ")))
        self.table.cellDoubleClicked.connect(self._open_routes); lay.addWidget(self.table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)
    def _open_routes(self, row, column):
        vehicle = self.table.item(row, 0).data(Qt.UserRole)
        if vehicle: VehicleRoutesDialog(self.repo, self.where_clause, self.carrier, vehicle, self).exec()

class TypeVehiclesDialog(QDialog):
    def __init__(self, repo, where_clause, carrier, type_name, parent=None):
        super().__init__(parent); self.setWindowTitle(f"ТС типа {type_name}"); self.resize(720, 520)
        lay = QVBoxLayout(self); title = QLabel(f"{type_name} — {carrier}"); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        chart = QPushButton("График типа ТС"); chart.clicked.connect(lambda: ChartsDialog(repo.get_daily_transactions(where_clause, carrier_name=carrier, type_name=type_name), self, f"Транзакции: {type_name}", "Количество транзакций по дням для выбранного типа ТС.", "Транзакции", False).exec()); lay.addWidget(chart)
        table = QTableWidget(0, 2); table.setHorizontalHeaderLabels(["Номер ТС", "Транзакции"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        for row, (vehicle, count) in enumerate(repo.get_type_vehicles(where_clause, carrier, type_name)):
            table.insertRow(row); table.setItem(row, 0, QTableWidgetItem(vehicle)); table.setItem(row, 1, QTableWidgetItem(f"{count:,}".replace(",", " ")))
        table.cellDoubleClicked.connect(lambda row, _: VehicleRoutesDialog(repo, where_clause, carrier, table.item(row, 0).text(), self).exec()); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class DriverVehiclesDialog(QDialog):
    def __init__(self, repo, where_clause, carrier, driver, parent=None):
        super().__init__(parent); self.setWindowTitle(f"Водитель — {driver}"); self.resize(720, 520)
        lay = QVBoxLayout(self); title = QLabel(driver); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        chart = QPushButton("График водителя"); chart.clicked.connect(lambda: ChartsDialog(repo.get_daily_transactions(where_clause, carrier_name=carrier, driver_name=driver), self, "Транзакции водителя", "Количество транзакций по дням для выбранного водителя.", "Транзакции", False).exec()); lay.addWidget(chart)
        table = QTableWidget(0, 2); table.setHorizontalHeaderLabels(["ТС", "Транзакции"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        for row, (vehicle, count) in enumerate(repo.get_driver_vehicles(where_clause, carrier, driver)):
            table.insertRow(row); table.setItem(row, 0, QTableWidgetItem(vehicle)); table.setItem(row, 1, QTableWidgetItem(f"{count:,}".replace(",", " ")))
        table.cellDoubleClicked.connect(lambda row, _: VehicleRoutesDialog(repo, where_clause, carrier, table.item(row, 0).text(), self).exec()); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class FleetDriversDialog(QDialog):
    def __init__(self, repo, where_clause, carrier, parent=None):
        super().__init__(parent); self.setWindowTitle(f"Водители — {carrier}"); self.resize(760, 560)
        lay = QVBoxLayout(self); title = QLabel("Водители перевозчика"); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        caption = QLabel("Дважды нажмите на водителя, чтобы открыть его ТС и маршруты."); caption.setStyleSheet("color:#6B7280;"); lay.addWidget(caption)
        table = QTableWidget(0, 2); table.setHorizontalHeaderLabels(["Водитель", "Транзакции"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        for row, (driver, count) in enumerate(repo.get_fleet_drivers(where_clause, carrier)):
            table.insertRow(row); table.setItem(row, 0, QTableWidgetItem(driver)); table.setItem(row, 1, QTableWidgetItem(f"{count:,}".replace(",", " ")))
        table.cellDoubleClicked.connect(lambda row, _: DriverVehiclesDialog(repo, where_clause, carrier, table.item(row, 0).text(), self).exec()); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class VehicleDriversCheckDialog(QDialog):
    """Детализация среднего чека по водителям одного ТС."""
    def __init__(self, repo, where_clause, carrier, vehicle, parent=None):
        super().__init__(parent); self.setWindowTitle(f"Водители ТС {vehicle}"); self.resize(760, 560)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; gridline-color:#EEF0F4; } QHeaderView::section { color:#6B7280; background:#FFFFFF; border:none; border-bottom:1px solid #D8DEE9; padding:8px; font-weight:600; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel(f"ТС {vehicle}"); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        caption = QLabel("Средний чек — только операции «Без льготы». Доля «С льготой» считается среди всех операций водителя на этом ТС.")
        caption.setWordWrap(True); caption.setStyleSheet("color:#6B7280;"); lay.addWidget(caption)
        headers = ["Водитель", "Средний чек", "С льготой", "Транзакции"]
        table = QTableWidget(0, len(headers)); table.setHorizontalHeaderLabels(headers); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setSortingEnabled(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        for col in range(1, len(headers)): table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        table.setSortingEnabled(False)
        for row, item in enumerate(repo.get_vehicle_driver_checks(where_clause, carrier, vehicle)):
            table.insertRow(row)
            values = [item["name"], f"{item['average_check']:,.2f} ₽".replace(",", " ").replace(".", ","), f"{item['benefit_share']:.1f}%".replace(".", ","), f"{item['transactions']:,}".replace(",", " ")]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if col: cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(row, col, cell)
        table.setSortingEnabled(True); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class CarrierCheckLeadersDialog(QDialog):
    """Полный перечень трёх независимых рейтингов предприятия."""
    def __init__(self, repo, where_clause, carrier, parent=None):
        super().__init__(parent); self.setWindowTitle(f"Средний чек — {carrier}"); self.resize(1180, 690)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QLabel { color:#1D1D1F; } QPushButton { color:#1D1D1F; background:#FFFFFF; border:1px solid #D8DEE9; border-radius:8px; padding:7px 9px; text-align:left; } QPushButton:hover { background:#F0F6FF; border-color:#0A84FF; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Средний чек и льготные операции"); title.setStyleSheet("font-size:22px; font-weight:700;"); lay.addWidget(title)
        subtitle = QLabel(f"{carrier} · средний чек рассчитывается только по категории «Без льготы»; «С льготой» — доля всех операций.")
        subtitle.setWordWrap(True); subtitle.setStyleSheet("color:#6B7280;"); lay.addWidget(subtitle)
        columns = QHBoxLayout(); columns.setSpacing(12)
        data = repo.get_carrier_check_leaders(where_clause, carrier)
        for key, caption in (("vehicles", "Транспортные средства"), ("routes", "Маршруты"), ("drivers", "Водители")):
            columns.addWidget(CheckLeadersColumn(caption, data.get(key, []), repo, where_clause, carrier, vehicle_clickable=(key == "vehicles"), parent=self), 1)
        lay.addLayout(columns, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class CheckLeadersColumn(QWidget):
    """Одна колонка рейтинга: ТС, маршруты или водители."""
    def __init__(self, title, rows, repo, where_clause, carrier, vehicle_clickable=False, parent=None):
        super().__init__(parent); self.setObjectName("checkLeadersColumn"); self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet("QWidget#checkLeadersColumn { background:#FFFFFF; border:1px solid #E5E7EB; border-radius:12px; } QLabel { background:transparent; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(12, 12, 12, 12); lay.setSpacing(6)
        heading = QLabel(title); heading.setStyleSheet("font-weight:700; font-size:14px;"); lay.addWidget(heading)
        if not rows:
            empty = QLabel("Нет операций «Без льготы»"); empty.setStyleSheet("color:#8E8E93; padding:12px 0;"); lay.addWidget(empty); lay.addStretch(); return
        scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.NoFrame); scroll.setStyleSheet("QScrollArea { border:none; background:transparent; }")
        content = QWidget(); inner = QVBoxLayout(content); inner.setContentsMargins(0, 0, 0, 0); inner.setSpacing(5)
        for index, item in enumerate(rows, 1):
            text = f"{index}. {item['name']}\n  {item['average_check']:,.2f} ₽  ·  С льготой {item['benefit_share']:.1f}%".replace(",", " ").replace(".", ",")
            if vehicle_clickable:
                row = QPushButton(text); row.setToolTip("Открыть водителей и средний чек по этому ТС")
                row.clicked.connect(lambda _, vehicle=item["name"]: VehicleDriversCheckDialog(repo, where_clause, carrier, vehicle, self).exec())
                inner.addWidget(row)
            else:
                row = QLabel(text); row.setWordWrap(True); row.setStyleSheet("background:#F8FAFC; border-radius:8px; padding:7px 8px; color:#303038;")
                inner.addWidget(row)
        inner.addStretch(); scroll.setWidget(content); lay.addWidget(scroll, 1)

class FleetWorkHoursDialog(QDialog):
    """Расшифровка оценки времени работы до отдельных транспортных средств."""
    def __init__(self, repo, where_clause, carrier=None, carriers=None, parent=None):
        super().__init__(parent); self.setWindowTitle("Рабочее время ТС"); self.resize(820, 590)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; gridline-color:#EEF0F4; } QHeaderView::section { color:#6B7280; background:#FFFFFF; border:none; border-bottom:1px solid #D8DEE9; padding:8px; font-weight:600; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Оценка рабочего времени ТС" if not carrier else f"Рабочее время ТС · {carrier}")
        title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        note = QLabel("Для каждого ТС в каждом дне берётся интервал между первой и последней транзакцией. Затем интервалы суммируются. Это оценка по данным реестра, а не табель рабочего времени.")
        note.setWordWrap(True); note.setStyleSheet("color:#6B7280;"); lay.addWidget(note)
        rows = repo.get_fleet_vehicle_work_hours(where_clause, carrier_name=carrier, carriers=carriers)
        include_carrier = carrier is None
        headers = (["Перевозчик"] if include_carrier else []) + ["Номер ТС", "Рабочее время", "Активных дней", "Транзакции"]
        table = QTableWidget(0, len(headers)); table.setHorizontalHeaderLabels(headers); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setSortingEnabled(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        for col in range(1, len(headers)): table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        table.setSortingEnabled(False)
        for row, item in enumerate(rows):
            table.insertRow(row)
            values = ([item["carrier"]] if include_carrier else []) + [item["vehicle"], f"{item['work_hours']:,.1f} ч".replace(",", " ").replace(".", ","), f"{item['work_days']:,}".replace(",", " "), f"{item['transactions']:,}".replace(",", " ")]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if col >= (2 if include_carrier else 1): cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(row, col, cell)
        table.setSortingEnabled(True); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class FleetCarrierCard(QWidget):
    def __init__(self, repo, where_clause, carrier, data, parent=None):
        super().__init__(parent); self.repo = repo; self.where_clause = where_clause; self.carrier = carrier
        self.setObjectName("fleetCard"); self.setAttribute(Qt.WA_StyledBackground, True); self._expanded = False; self._details_ready = False
        self.setStyleSheet("QWidget#fleetCard { background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 14px; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(16, 14, 16, 14); lay.setSpacing(10)
        self.lay = lay
        head = QHBoxLayout(); title = QLabel(carrier); title.setStyleSheet("font-size: 15px; font-weight: 700; color: #1D1D1F;"); head.addWidget(title); head.addStretch()
        chart = QPushButton("График"); chart.setFixedHeight(28); chart.clicked.connect(lambda: ChartsDialog(repo.get_daily_transactions(where_clause, carrier_name=carrier), self, "Транзакции перевозчика", "Количество транзакций по дням для выбранного перевозчика.", "Транзакции", False).exec()); head.addWidget(chart); lay.addLayout(head)
        self.details_button = QPushButton("Подробнее ▾"); self.details_button.setFixedHeight(28); self.details_button.setToolTip("Развернуть аналитику среднего чека")
        self.details_button.clicked.connect(self._toggle_details); head.addWidget(self.details_button)
        metrics = QHBoxLayout(); metrics.setSpacing(18)
        icon = FleetIcon("Автобус"); metrics.addWidget(icon)
        vehicles_btn = QPushButton(f"{data['vehicles']:,} ТС".replace(",", " "))
        vehicles_btn.setToolTip("Открыть список уникальных транспортных средств")
        vehicles_btn.setStyleSheet("QPushButton { color:#0A84FF; background:transparent; border:none; font-weight:700; text-align:left; padding:0; } QPushButton:hover { color:#0077ED; text-decoration:underline; }")
        vehicles_btn.clicked.connect(lambda: FleetVehiclesDialog(repo, where_clause, carrier, self).exec())
        metrics.addWidget(vehicles_btn)
        if data.get("vehicles_by_type", data["vehicles"]) != data["vehicles"]:
            repeated = data["vehicles_by_type"] - data["vehicles"]
            typed = QPushButton(f"{repeated:,} повтор по типам".replace(",", " "))
            typed.setToolTip("Показать номера ТС, встречающиеся в нескольких типах")
            typed.setStyleSheet("QPushButton { color:#B45309; background:#FFF7ED; border:1px solid #FED7AA; border-radius:7px; padding:3px 6px; font-size:11px; font-weight:700; } QPushButton:hover { background:#FFEDD5; }")
            typed.clicked.connect(lambda: DuplicateVehiclesDialog(repo, where_clause, carrier, self).exec()); metrics.addWidget(typed)
        for icon_type, label, value in (("Маршрут", "Маршруты", data["routes"]),):
            icon = FleetIcon(icon_type); metrics.addWidget(icon)
            routes_btn = QPushButton(f"{value:,} маршрутов".replace(",", " ")); routes_btn.setToolTip("Открыть маршруты и ТС на них")
            routes_btn.setStyleSheet("QPushButton { color:#1D1D1F; background:transparent; border:none; font-weight:700; text-align:left; padding:0; } QPushButton:hover { color:#0077ED; text-decoration:underline; }")
        routes_btn.clicked.connect(lambda: FleetRoutesDialog(repo, where_clause, carrier, self).exec()); metrics.addWidget(routes_btn)
        drivers_btn = QPushButton(f"{data.get('drivers', 0):,} водителей".replace(",", " ")); drivers_btn.setToolTip("Открыть список водителей")
        drivers_btn.setStyleSheet("QPushButton { color:#1D1D1F; background:transparent; border:none; font-weight:700; text-align:left; padding:0; } QPushButton:hover { color:#0077ED; text-decoration:underline; }")
        drivers_btn.clicked.connect(lambda: FleetDriversDialog(repo, where_clause, carrier, self).exec()); metrics.addWidget(drivers_btn)
        transactions = QLabel(f"<b>{data['transactions']:,}</b><br><span style='color:#6B7280'>транзакций</span>".replace(",", " "))
        transactions.setStyleSheet("color: #1D1D1F; font-size: 12px;"); metrics.addStretch(); metrics.addWidget(transactions); lay.addLayout(metrics)
        avg = data["transactions"] / data["vehicles"] if data["vehicles"] else 0
        load = QLabel(f"<b>{avg:,.0f}</b><br><span style='color:#6B7280'>на 1 ТС</span>".replace(",", " "))
        load.setStyleSheet("color: #1D1D1F; font-size: 12px;"); metrics.addWidget(load)
        work = QPushButton(f"{data.get('work_hours', 0):,.1f} ч\nоценка работы".replace(",", " ").replace(".", ","))
        work.setToolTip("Открыть список ТС и их оценочное рабочее время")
        work.setStyleSheet("QPushButton { color:#1D1D1F; background:transparent; border:none; font-size:12px; font-weight:700; text-align:right; padding:0; } QPushButton:hover { color:#0077ED; text-decoration:underline; }")
        work.clicked.connect(lambda: FleetWorkHoursDialog(repo, where_clause, carrier=carrier, parent=self).exec()); metrics.addWidget(work)
        line = QFrame(); line.setFrameShape(QFrame.HLine); line.setStyleSheet("color: #EEF0F4;"); lay.addWidget(line)
        for item in sorted(data["types"], key=lambda row: row["transactions"], reverse=True):
            row = QHBoxLayout(); row.setSpacing(8); row.addWidget(FleetIcon(item["type"]))
            name = QPushButton(item["type"]); name.setToolTip("Открыть список ТС этого типа")
            name.setStyleSheet("QPushButton { color:#1D1D1F; background:transparent; border:none; font-weight:600; text-align:left; padding:0; } QPushButton:hover { color:#0077ED; text-decoration:underline; }")
            name.clicked.connect(lambda _, type_name=item["type"]: TypeVehiclesDialog(repo, where_clause, carrier, type_name, self).exec()); row.addWidget(name); row.addStretch()
            values = QLabel(f"{item['vehicles']:,} ТС   ·   {item['transactions']:,} транзакций".replace(",", " "))
            values.setStyleSheet("color:#6B7280; font-size:12px;"); row.addWidget(values); lay.addLayout(row)
        self.details = QWidget(); self.details.setObjectName("carrierCheckDetails"); self.details.setAttribute(Qt.WA_StyledBackground, True)
        self.details.setStyleSheet("QWidget#carrierCheckDetails { background:#F8FAFC; border:1px solid #E5E7EB; border-radius:12px; }")
        self.details_lay = QVBoxLayout(self.details); self.details_lay.setContentsMargins(12, 12, 12, 12); self.details_lay.setSpacing(8)
        self.details.setVisible(False); lay.addWidget(self.details)

    def mousePressEvent(self, event):
        # Нажатие на свободную область карточки — тот же жест, что «Подробнее».
        if event.button() == Qt.LeftButton:
            self._toggle_details()
            event.accept()
            return
        super().mousePressEvent(event)

    def _toggle_details(self):
        self._expanded = not self._expanded
        if self._expanded and not self._details_ready:
            self._build_details()
        self.details.setVisible(self._expanded)
        self.details_button.setText("Свернуть ▴" if self._expanded else "Подробнее ▾")

    def _build_details(self):
        self._details_ready = True
        title = QLabel("Коротко: средний чек и доля льготных операций")
        title.setStyleSheet("font-weight:700; color:#1D1D1F;"); self.details_lay.addWidget(title)
        note = QLabel("Средний чек — только «Без льготы»; «С льготой» — доля среди всех операций.")
        note.setStyleSheet("color:#6B7280; font-size:12px;"); self.details_lay.addWidget(note)
        data = self.repo.get_carrier_check_leaders(self.where_clause, self.carrier, limit=10)
        columns = QHBoxLayout(); columns.setSpacing(10)
        for key, heading in (("vehicles", "Топ‑10 ТС"), ("routes", "Топ‑10 маршрутов"), ("drivers", "Топ‑10 водителей")):
            column = CheckLeadersColumn(heading, data.get(key, []), self.repo, self.where_clause, self.carrier, vehicle_clickable=(key == "vehicles"), parent=self)
            column.setFixedHeight(305); columns.addWidget(column, 1)
        self.details_lay.addLayout(columns)
        bottom = QHBoxLayout(); bottom.addStretch()
        more = QPushButton("Ещё — весь перечень")
        more.clicked.connect(lambda: CarrierCheckLeadersDialog(self.repo, self.where_clause, self.carrier, self).exec())
        bottom.addWidget(more); self.details_lay.addLayout(bottom)

class FleetAnalyticsDialog(QDialog):
    def __init__(self, repo, base_where, parent=None):
        super().__init__(parent); self.repo = repo; self.base_where = base_where; self.setWindowTitle("Аналитика парка"); self.resize(1180, 760)
        self.setStyleSheet("QDialog { background: #F5F6F8; color: #1D1D1F; } QLabel, QCheckBox, QListWidget { color: #1D1D1F; } QLineEdit, QListWidget { background: #FFFFFF; color: #1D1D1F; border: 1px solid #D8DEE9; border-radius: 8px; padding: 6px; } QPushButton { color: #1D1D1F; background: #FFFFFF; border: 1px solid #D8DEE9; border-radius: 8px; padding: 7px 10px; font-weight: 600; } QPushButton:hover { background: #F0F6FF; border-color: #0A84FF; }")
        root = QHBoxLayout(self); root.setContentsMargins(18, 18, 18, 18); root.setSpacing(18)
        filters = QWidget(); filters.setFixedWidth(265); fl = QVBoxLayout(filters); fl.setContentsMargins(0, 0, 0, 0); fl.setSpacing(10)
        title = QLabel("ПЕРЕВОЗЧИКИ"); title.setStyleSheet("font-weight:700; color:#1D1D1F;"); fl.addWidget(title)
        self.search = QLineEdit(); self.search.setPlaceholderText("Поиск перевозчика…"); self.search.setClearButtonEnabled(True); self.search.textChanged.connect(self._filter_carriers); fl.addWidget(self.search)
        self.all_cb = QCheckBox("Все перевозчики"); self.all_cb.setChecked(True); self.all_cb.stateChanged.connect(self._toggle_all); fl.addWidget(self.all_cb)
        self._changing_all = False
        self.carriers = QListWidget(); self.carriers.setSelectionMode(QAbstractItemView.NoSelection); self.carriers.itemChanged.connect(self._carrier_changed); fl.addWidget(self.carriers, 1)
        apply = QPushButton("Применить"); apply.clicked.connect(self.refresh); fl.addWidget(apply)
        root.addWidget(filters)
        right = QWidget(); rl = QVBoxLayout(right); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(8)
        heading = QLabel("Аналитика парка"); heading.setStyleSheet("font-size:24px; font-weight:700; color:#1D1D1F;"); rl.addWidget(heading)
        self.total_label = QLabel(); self.total_label.setObjectName("fleetTotal")
        self.total_label.setStyleSheet("QLabel#fleetTotal { background:#FFFFFF; border:1px solid #D8DEE9; border-radius:12px; color:#1D1D1F; font-size:14px; font-weight:600; padding:12px 14px; }")
        rl.addWidget(self.total_label)
        self.total_work_button = QPushButton("Рабочее время по всем ТС")
        self.total_work_button.setToolTip("Открыть список ТС с оценкой рабочего времени")
        self.total_work_button.clicked.connect(self._open_total_work_hours)
        rl.addWidget(self.total_work_button)
        self.duplicates_button = QPushButton("Повторяющиеся номера ТС")
        self.duplicates_button.setToolTip("Показать номера, встречающиеся в разных перевозчиках или типах ТС")
        self.duplicates_button.clicked.connect(self._open_global_duplicates)
        rl.addWidget(self.duplicates_button)
        self.caption = QLabel("Уникальные ТС, маршруты и транзакции по выбранным перевозчикам."); self.caption.setStyleSheet("color:#6B7280;"); rl.addWidget(self.caption)
        self.scroll = QScrollArea(); self.scroll.setWidgetResizable(True); self.scroll.setFrameShape(QFrame.NoFrame); self.scroll.setStyleSheet("QScrollArea { background:#F5F6F8; border:none; } QScrollBar:vertical { background:transparent; width:10px; } QScrollBar::handle:vertical { background:#C7C7CC; border-radius:5px; }")
        self.cards = QWidget(); self.cards.setObjectName("fleetCards"); self.cards.setAttribute(Qt.WA_StyledBackground, True); self.cards.setStyleSheet("QWidget#fleetCards { background:#F5F6F8; }"); self.cards_lay = QVBoxLayout(self.cards); self.cards_lay.setContentsMargins(0, 8, 0, 0); self.cards_lay.setSpacing(10); self.cards_lay.addStretch(1)
        self.scroll.setWidget(self.cards); rl.addWidget(self.scroll, 1); root.addWidget(right, 1)
        self._load_carriers(); self.refresh()
    def _load_carriers(self):
        self.carriers.blockSignals(True)
        for name, count in self.repo.get_fleet_carriers(self.base_where):
            item = QListWidgetItem(f"{name}  ({count:,})".replace(",", " ")); item.setData(Qt.UserRole, name)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable); item.setCheckState(Qt.Checked); self.carriers.addItem(item)
        self.carriers.blockSignals(False)
    def _toggle_all(self, state):
        checked = Qt.Checked if state == Qt.Checked else Qt.Unchecked; self._changing_all = True; self.carriers.blockSignals(True)
        for i in range(self.carriers.count()): self.carriers.item(i).setCheckState(checked)
        self.carriers.blockSignals(False)
        self._changing_all = False
        self.refresh()
    def _carrier_changed(self, item):
        if not self._changing_all:
            self.all_cb.blockSignals(True); self.all_cb.setChecked(False); self.all_cb.blockSignals(False)
            QTimer.singleShot(0, self.refresh)
    def _filter_carriers(self, text):
        text = text.casefold()
        for i in range(self.carriers.count()): self.carriers.item(i).setHidden(text not in self.carriers.item(i).text().casefold())
    def refresh(self):
        selected = [self.carriers.item(i).data(Qt.UserRole) for i in range(self.carriers.count()) if self.carriers.item(i).checkState() == Qt.Checked]
        self.selected_carriers = selected
        data = self.repo.get_fleet_summary(self.base_where, selected) if selected else {}
        while self.cards_lay.count():
            item = self.cards_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        for carrier, item in sorted(data.items(), key=lambda pair: pair[1]["transactions"], reverse=True): self.cards_lay.addWidget(FleetCarrierCard(self.repo, self.base_where, carrier, item))
        totals = self.repo.get_fleet_totals(self.base_where, selected) if selected else {"vehicles": 0, "vehicles_by_type": 0, "routes": 0, "transactions": 0, "work_hours": 0.0}
        vehicle_text = f"{totals['vehicles']:,} уникальных ТС".replace(",", " ")
        if totals["vehicles_by_type"] != totals["vehicles"]:
            vehicle_text += f"   ·   {totals['vehicles_by_type']:,} по типам".replace(",", " ")
        self.total_label.setText(f"Итого по выбранным перевозчикам:  {vehicle_text}   ·   {totals['routes']:,} маршрутов   ·   {totals['transactions']:,} транзакций   ·   {totals['work_hours']:,.1f} ч работы".replace(",", " ").replace(".", ","))
        self.total_work_button.setText(f"Рабочее время по всем ТС · {totals['work_hours']:,.1f} ч".replace(",", " ").replace(".", ","))
        self.cards_lay.addStretch(1); self.caption.setText(f"{len(data)} перевозчиков · уникальные ТС, маршруты и транзакции.")
    def _open_global_duplicates(self):
        GlobalDuplicateVehiclesDialog(self.repo, self.base_where, getattr(self, "selected_carriers", None), self).exec()
    def _open_total_work_hours(self):
        FleetWorkHoursDialog(self.repo, self.base_where, carriers=getattr(self, "selected_carriers", None), parent=self).exec()

class OperationalRouteCard(QWidget):
    """Карточка одного маршрута выбранного перевозчика."""
    def __init__(self, repo, where_clause, carrier, route, rows, parent=None):
        super().__init__(parent); self.repo = repo; self.where_clause = where_clause; self.carrier = carrier; self.route = route
        self.setObjectName("operationalRouteCard"); self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet("QWidget#operationalRouteCard { background:#FFFFFF; border:1px solid #E3E8F1; border-radius:18px; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(18, 16, 18, 16); lay.setSpacing(12)
        transactions = sum(item["transactions"] for item in rows)
        amount = sum(item["calculated_amount"] for item in rows)
        share = (sum(item["benefit_share"] * item["transactions"] for item in rows) / transactions) if transactions else 0
        head = QHBoxLayout(); head.setSpacing(10); route_icon = FleetIcon("Маршрут"); route_icon.setFixedSize(36, 36); head.addWidget(route_icon)
        name_box = QVBoxLayout(); name_box.setSpacing(1); eyebrow = QLabel("МАРШРУТ"); eyebrow.setStyleSheet("font-size:10px; font-weight:700; letter-spacing:0.8px; color:#0A84FF;"); name_box.addWidget(eyebrow)
        title = QLabel(str(route)); title.setStyleSheet("font-size:19px; font-weight:750; color:#1D1D1F;"); name_box.addWidget(title); head.addLayout(name_box); head.addStretch()
        chart = QPushButton("◌  График"); chart.setFixedHeight(30); chart.clicked.connect(lambda: ChartsDialog(repo.get_daily_transactions(where_clause, carrier_name=carrier, route_name=route), self, f"Маршрут {route}", "Количество транзакций по дням для выбранного маршрута.", "Транзакции", False).exec()); head.addWidget(chart)
        self.expand = QPushButton("Состав ▾"); self.expand.setFixedHeight(30); self.expand.clicked.connect(self._toggle); head.addWidget(self.expand); lay.addLayout(head)
        metrics = QHBoxLayout(); metrics.setSpacing(9)
        for color, value, caption in (("#EAF3FF", f"{transactions:,}".replace(",", " "), "транзакций"), ("#EAFBF2", f"{amount:,.2f} ₽".replace(",", " ").replace(".", ","), "расчётная сумма"), ("#F2EDFF", f"{share:.1f}%".replace(".", ","), "с льготой")):
            metric = QLabel(f"<b>{value}</b><br><span>{caption}</span>"); metric.setStyleSheet(f"background:{color}; border-radius:10px; padding:8px 12px; color:#1D1D1F; font-size:12px;"); metrics.addWidget(metric)
        metrics.addStretch(); lay.addLayout(metrics)
        self.details = QWidget(); self.details.setVisible(False); details_lay = QVBoxLayout(self.details); details_lay.setContentsMargins(0, 4, 0, 0); details_lay.setSpacing(10)
        by_vehicle = {}
        for item in rows: by_vehicle.setdefault((item["vehicle"], item.get("vehicle_type", "Не указан")), []).append(item)
        for (vehicle, vehicle_type), vehicle_rows in sorted(by_vehicle.items(), key=lambda pair: sum(x["calculated_amount"] for x in pair[1]), reverse=True):
            vehicle_transactions = sum(x["transactions"] for x in vehicle_rows); vehicle_amount = sum(x["calculated_amount"] for x in vehicle_rows)
            vehicle_box = QWidget(); vehicle_box.setObjectName("operationalVehicle"); vehicle_box.setAttribute(Qt.WA_StyledBackground, True); vehicle_box.setStyleSheet("QWidget#operationalVehicle { background:#F8FAFC; border:1px solid #E6EBF2; border-radius:12px; }")
            vb = QVBoxLayout(vehicle_box); vb.setContentsMargins(13, 11, 13, 11); vb.setSpacing(7)
            vehicle_head = QHBoxLayout(); vehicle_head.setSpacing(8); icon = FleetIcon(vehicle_type); icon.setFixedSize(30, 30); vehicle_head.addWidget(icon)
            vehicle_name_box = QVBoxLayout(); vehicle_name_box.setSpacing(0); vname = QLabel(f"ТС {vehicle}"); vname.setStyleSheet("font-size:14px; font-weight:750; color:#1D1D1F;"); vehicle_name_box.addWidget(vname); vtype = QLabel(vehicle_type); vtype.setStyleSheet("color:#6B7280; font-size:11px;"); vehicle_name_box.addWidget(vtype); vehicle_head.addLayout(vehicle_name_box); vehicle_head.addStretch()
            vstats = QLabel(f"<b>{vehicle_transactions:,}</b> поездок<br><span>{vehicle_amount:,.2f} ₽</span>".replace(",", " ").replace(".", ",")); vstats.setStyleSheet("color:#374151; font-size:12px; text-align:right;"); vehicle_head.addWidget(vstats); vb.addLayout(vehicle_head)
            for driver_row in sorted(vehicle_rows, key=lambda item: item["calculated_amount"], reverse=True):
                driver_box = QWidget(); driver_box.setObjectName("operationalDriver"); driver_box.setAttribute(Qt.WA_StyledBackground, True); driver_box.setStyleSheet("QWidget#operationalDriver { background:#FFFFFF; border-radius:8px; }")
                dl = QHBoxLayout(driver_box); dl.setContentsMargins(9, 7, 9, 7); dl.setSpacing(10)
                initials = "".join(part[0] for part in str(driver_row["driver"]).split()[:2]).upper() or "—"
                avatar = QLabel(initials); avatar.setFixedSize(26, 26); avatar.setAlignment(Qt.AlignCenter); avatar.setStyleSheet("background:#EAF3FF; color:#0A66CC; border-radius:13px; font-size:10px; font-weight:800;"); dl.addWidget(avatar)
                dname = QLabel(driver_row["driver"]); dname.setMinimumWidth(150); dname.setStyleSheet("font-weight:650; color:#25252B;"); dl.addWidget(dname, 1)
                for value, caption in ((f"{driver_row['transactions']:,}".replace(",", " "), "поездок"), (f"{driver_row['calculated_amount']:,.2f} ₽".replace(",", " ").replace(".", ","), "сумма"), (f"{driver_row['average_operation']:,.2f} ₽".replace(",", " ").replace(".", ","), "средняя"), (f"{driver_row['benefit_share']:.1f}%".replace(".", ","), "льгота")):
                    stat = QLabel(f"<b>{value}</b><br><span>{caption}</span>"); stat.setAlignment(Qt.AlignRight | Qt.AlignVCenter); stat.setStyleSheet("color:#374151; font-size:11px;"); dl.addWidget(stat)
                vb.addWidget(driver_box)
            details_lay.addWidget(vehicle_box)
        lay.addWidget(self.details)
    def _toggle(self):
        visible = not self.details.isVisible(); self.details.setVisible(visible); self.expand.setText("Свернуть ▴" if visible else "Состав ▾")

class OperationalAnalyticsDialog(QDialog):
    """Карточная аналитика маршрутов с выбором перевозчика."""
    def __init__(self, repo, where_clause, parent=None):
        super().__init__(parent); self.repo = repo; self.where_clause = where_clause; self.rows = repo.get_operational_analysis(where_clause)
        self.setWindowTitle("Эффективность маршрутов"); self.resize(1280, 760)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QLabel, QListWidget { color:#1D1D1F; } QLineEdit, QListWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; border-radius:8px; padding:6px; } QPushButton { color:#1D1D1F; background:#FFFFFF; border:1px solid #D8DEE9; border-radius:8px; padding:7px 10px; font-weight:600; } QPushButton:hover { background:#F0F6FF; border-color:#0A84FF; }")
        root = QHBoxLayout(self); root.setContentsMargins(18, 18, 18, 18); root.setSpacing(18)
        left = QWidget(); left.setFixedWidth(270); left_lay = QVBoxLayout(left); left_lay.setContentsMargins(0, 0, 0, 0); left_lay.setSpacing(10)
        left_title = QLabel("ПЕРЕВОЗЧИКИ"); left_title.setStyleSheet("font-weight:700;"); left_lay.addWidget(left_title)
        self.carrier_search = QLineEdit(); self.carrier_search.setPlaceholderText("Поиск перевозчика…"); self.carrier_search.setClearButtonEnabled(True); self.carrier_search.textChanged.connect(self._filter_carriers); left_lay.addWidget(self.carrier_search)
        self.carriers = QListWidget(); self.carriers.setSelectionMode(QAbstractItemView.SingleSelection); self.carriers.currentItemChanged.connect(lambda *_: self._refresh_cards()); left_lay.addWidget(self.carriers, 1)
        root.addWidget(left)
        right = QWidget(); right_lay = QVBoxLayout(right); right_lay.setContentsMargins(0, 0, 0, 0); right_lay.setSpacing(8)
        title = QLabel("Эффективность маршрутов"); title.setStyleSheet("font-size:24px; font-weight:700; color:#1D1D1F;"); right_lay.addWidget(title)
        caption = QLabel("Выберите перевозчика — затем откройте маршрут, ТС и водителей. Расчётная сумма выбирает один источник по реализатору; основная сводка не меняется."); caption.setWordWrap(True); caption.setStyleSheet("color:#6B7280;"); right_lay.addWidget(caption)
        self.search = QLineEdit(); self.search.setPlaceholderText("Поиск маршрута, ТС или водителя…"); self.search.setClearButtonEnabled(True); self.search.textChanged.connect(self._refresh_cards); right_lay.addWidget(self.search)
        self.summary = QLabel(); self.summary.setStyleSheet("background:#FFFFFF; border:1px solid #D8DEE9; border-radius:10px; padding:9px 12px; font-weight:700;"); right_lay.addWidget(self.summary)
        scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.NoFrame); scroll.setStyleSheet("QScrollArea { background:#F5F6F8; border:none; }")
        self.cards = QWidget(); self.cards.setAttribute(Qt.WA_StyledBackground, True); self.cards.setStyleSheet("background:#F5F6F8;"); self.cards_lay = QVBoxLayout(self.cards); self.cards_lay.setContentsMargins(0, 6, 0, 0); self.cards_lay.setSpacing(10); scroll.setWidget(self.cards); right_lay.addWidget(scroll, 1); root.addWidget(right, 1)
        self._load_carriers()
    def _load_carriers(self):
        all_item = QListWidgetItem("Все перевозчики"); all_item.setData(Qt.UserRole, None); self.carriers.addItem(all_item)
        for carrier in sorted({row["carrier"] for row in self.rows}, key=str.casefold):
            item = QListWidgetItem(carrier); item.setData(Qt.UserRole, carrier); self.carriers.addItem(item)
        self.carriers.setCurrentRow(0)
    def _filter_carriers(self, text):
        text = text.casefold()
        for index in range(self.carriers.count()):
            item = self.carriers.item(index); item.setHidden(index > 0 and text not in item.text().casefold())
    def _refresh_cards(self):
        while self.cards_lay.count():
            item = self.cards_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        current = self.carriers.currentItem(); selected_carrier = current.data(Qt.UserRole) if current else None
        search = self.search.text().casefold()
        rows = [row for row in self.rows if (selected_carrier is None or row["carrier"] == selected_carrier) and (not search or search in " ".join(str(row[key]) for key in ("route", "vehicle", "driver")).casefold())]
        grouped = {}
        for row in rows: grouped.setdefault((row["carrier"], row["route"]), []).append(row)
        for (carrier, route), route_rows in sorted(grouped.items(), key=lambda pair: sum(row["calculated_amount"] for row in pair[1]), reverse=True):
            self.cards_lay.addWidget(OperationalRouteCard(self.repo, self.where_clause, carrier, route, route_rows, self))
        self.cards_lay.addStretch(1)
        amount = sum(row["calculated_amount"] for row in rows); transactions = sum(row["transactions"] for row in rows)
        self.summary.setText(f"{len(grouped):,} маршрутов   ·   {transactions:,} транзакций   ·   {amount:,.2f} ₽ расчётной суммы".replace(",", " ").replace(".", ","))

class FraudCardsDialog(QDialog):
    """Контроль масок обслуженных карт в безналичной форме оплаты."""
    def __init__(self, repo, where_clause, parent=None):
        super().__init__(parent); self.repo = repo; self.where_clause = where_clause; self.setWindowTitle("ФРОД — безналичные карты"); self.resize(980, 680)
        self.rows = repo.get_fraud_cards(where_clause)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QLineEdit, QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; border-radius:8px; padding:6px; } QTableWidget { border-radius:10px; gridline-color:#EEF0F4; } QHeaderView::section { background:#FFFFFF; color:#6B7280; border:none; border-bottom:1px solid #D8DEE9; padding:8px; font-weight:600; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("ФРОД · маски обслуженных карт"); title.setStyleSheet("font-size:24px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        caption = QLabel("Текущий срез фильтров: только строки, где «Форма оплаты проезда» содержит «Безналич». Каждая строка показывает количество операций одной маски карты.")
        caption.setWordWrap(True); caption.setStyleSheet("color:#6B7280;"); lay.addWidget(caption)
        total_transactions = sum(row["transactions"] for row in self.rows); total_amount = sum(row["amount"] for row in self.rows)
        self.total = QLabel(f"Безналичных операций: {total_transactions:,}   ·   Сумма оплат: {total_amount:,.2f} ₽".replace(",", " ").replace(".", ","))
        self.total.setStyleSheet("background:#FFFFFF; border:1px solid #D8DEE9; border-radius:10px; padding:10px 12px; font-weight:700; color:#1D1D1F;"); lay.addWidget(self.total)
        self.search = QLineEdit(); self.search.setPlaceholderText("Поиск по маске карты…"); self.search.setClearButtonEnabled(True); self.search.textChanged.connect(self._filter); lay.addWidget(self.search)
        self.table = QTableWidget(0, 3); self.table.setHorizontalHeaderLabels(["Маска обслуженной карты", "Количество операций", "Сумма оплат"]); self.table.verticalHeader().setVisible(False); self.table.setEditTriggers(QAbstractItemView.NoEditTriggers); self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents); self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents); lay.addWidget(self.table, 1)
        note = QLabel("Количество — число транзакций; сумма — «Оплаченная сумма». Двойной клик по маске открывает все исходные строки."); note.setStyleSheet("color:#6B7280; font-size:12px;"); lay.addWidget(note)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)
        self.table.cellDoubleClicked.connect(self._open_details)
        self._fill(self.rows)
    def _fill(self, rows):
        self.table.setSortingEnabled(False); self.table.setRowCount(0)
        for row, item in enumerate(rows):
            self.table.insertRow(row)
            values = [item["mask"], f"{item['transactions']:,}".replace(",", " "), f"{item['amount']:,.2f} ₽".replace(",", " ").replace(".", ",")]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if col: cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row, col, cell)
        self.table.setSortingEnabled(True)
    def _filter(self, text):
        text = text.casefold(); self._fill([row for row in self.rows if text in row["mask"].casefold()])
    def _open_details(self, row, _column):
        item = self.table.item(row, 0)
        if item:
            FraudCardDetailsDialog(self.repo, self.where_clause, item.text(), self).exec()

class FraudCardDetailsDialog(QDialog):
    def __init__(self, repo, where_clause, card_mask, parent=None):
        super().__init__(parent); self.setWindowTitle(f"Операции карты {card_mask}"); self.resize(1240, 700)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableView { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; border-radius:10px; alternate-background-color:#F8FAFC; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel(f"Маска карты: {card_mask}"); title.setStyleSheet("font-size:20px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        note = QLabel("Все исходные строки с этой маской и безналичной формой оплаты."); note.setStyleSheet("color:#6B7280;"); lay.addWidget(note)
        table = QTableView(); table.setSelectionBehavior(QAbstractItemView.SelectRows); table.setAlternatingRowColors(True); table.setShowGrid(False); table.verticalHeader().setVisible(False)
        df = repo.get_fraud_card_rows(where_clause, card_mask)
        self.model = PandasTableModel(df); table.setModel(self.model); table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents); table.horizontalHeader().setStretchLastSection(True); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class InvestorLoadDialog(QDialog):
    """Самостоятельное окно подготовки архива инвестора."""
    def __init__(self, repo, parent=None):
        super().__init__(parent); self.repo = repo; self.archive_path = ""; self.worker = None
        self.setWindowTitle("Сверка с ВТБ"); self.setFixedSize(620, 350)
        self.setStyleSheet("""
            QDialog { background:#F5F7FB; color:#172033; }
            QFrame#loadCard { background:#FFFFFF; border:1px solid #E2E8F0; border-radius:16px; }
            QLabel#dialogTitle { color:#172033; font-size:24px; font-weight:750; }
            QLabel#dialogNote, QLabel#loadStatus { color:#536174; font-size:13px; }
            QLabel#readyBadge { color:#155E3A; background:#EAF8F0; border:1px solid #B7E6C8; border-radius:8px; padding:5px 8px; font-weight:650; }
            QLineEdit { background:#F8FAFC; color:#172033; border:1px solid #D8E0EB; border-radius:9px; padding:9px 10px; }
            QPushButton { color:#263247; background:#FFFFFF; border:1px solid #D2DAE6; border-radius:9px; padding:9px 12px; font-weight:650; }
            QPushButton:hover { background:#F0F5FF; border-color:#7EA7F8; }
            QPushButton#primaryAction { background:#1769E0; color:#FFFFFF; border:1px solid #1769E0; padding:10px 15px; font-weight:750; }
            QPushButton#primaryAction:hover { background:#0F57C4; border-color:#0F57C4; }
            QPushButton#primaryAction:disabled { background:#B9C7DE; border-color:#B9C7DE; color:#FFFFFF; }
            QProgressBar { color:#172033; background:#E8EDF5; border:0; border-radius:8px; height:20px; text-align:center; font-weight:750; }
            QProgressBar::chunk { background:#1769E0; border-radius:8px; }
        """)
        lay = QVBoxLayout(self); lay.setContentsMargins(24, 22, 24, 22); lay.setSpacing(12)
        title = QLabel("Сверка с ВТБ"); title.setObjectName("dialogTitle"); lay.addWidget(title)
        note = QLabel("Сопоставление АСОП с отчётом ВТБ по ID транзакции. Архив нужно подготовить только один раз.")
        note.setObjectName("dialogNote"); note.setWordWrap(True); lay.addWidget(note)
        card = QFrame(); card.setObjectName("loadCard"); card_lay = QVBoxLayout(card); card_lay.setContentsMargins(14, 14, 14, 14); card_lay.setSpacing(10)
        row = QHBoxLayout(); self.path = QLineEdit(); self.path.setReadOnly(True); self.path.setPlaceholderText("Выберите финальный файл .zip"); row.addWidget(self.path, 1)
        choose = QPushButton("Выбрать архив…"); choose.clicked.connect(self._choose); row.addWidget(choose); card_lay.addLayout(row)
        self.ready_badge = QLabel(); self.ready_badge.setObjectName("readyBadge"); self.ready_badge.hide(); card_lay.addWidget(self.ready_badge)
        self.progress = QProgressBar(); self.progress.setRange(0, 100); self.progress.setTextVisible(True); self.progress.setFormat("%p%"); card_lay.addWidget(self.progress)
        self.status = QLabel("Выберите архив отчёта ВТБ для начала сверки."); self.status.setObjectName("loadStatus"); card_lay.addWidget(self.status)
        lay.addWidget(card)
        actions = QHBoxLayout(); actions.addStretch(); self.start = QPushButton("Подготовить и сверить"); self.start.setObjectName("primaryAction"); self.start.setEnabled(False); self.start.clicked.connect(self._start); actions.addWidget(self.start); lay.addLayout(actions)
        remembered = QSettings("TransactionAnalyzer", "Analyzer").value("investor/archive_path", "", type=str)
        if remembered and os.path.isfile(remembered):
            self._set_archive(remembered)

    def _output_db_path(self):
        return os.path.splitext(self.archive_path)[0] + "_reconciliation.duckdb"

    def _prepared_db_exists(self, db_path):
        """Проверяет, что ранее подготовленная витрина действительно пригодна для сверки."""
        if not db_path or not os.path.isfile(db_path): return False
        try:
            conn = duckdb.connect(db_path, read_only=True)
            exists = conn.execute("SELECT COUNT(*) FROM information_schema.tables WHERE table_name='investor_data'").fetchone()[0] > 0
            conn.close()
            return exists
        except Exception:
            return False

    def _set_archive(self, path):
        self.archive_path = path; self.path.setText(path)
        QSettings("TransactionAnalyzer", "Analyzer").setValue("investor/archive_path", path)
        ready = self._prepared_db_exists(self._output_db_path())
        self.ready_badge.setVisible(ready)
        if ready:
            self.ready_badge.setText("✓ Архив уже подготовлен — повторная загрузка не нужна")
            self.status.setText("Можно сразу открыть результаты последней подготовки.")
            self.start.setText("Открыть готовую сверку")
            self.progress.setValue(100)
        else:
            self.status.setText("Архив выбран. Все части .z01, .z02 и далее должны лежать рядом с .zip.")
            self.start.setText("Подготовить и сверить")
            self.progress.setValue(0)
        self.start.setEnabled(True)
    def _choose(self):
        path, _ = QFileDialog.getOpenFileName(self, "Открыть отчёт ВТБ", "", "ZIP архив (*.zip)")
        if path:
            self._set_archive(path)
    def _start(self):
        output_db = self._output_db_path()
        if self._prepared_db_exists(output_db):
            self.status.setText("Открываем подготовленную сверку…")
            self._open_reconciliation(output_db)
            return
        self.start.setEnabled(False); self.progress.setValue(0); self.status.setText("Подготовка…")
        self.worker = InvestorArchiveWorker(self.archive_path, output_db, self)
        self.worker.progress.connect(lambda value, text: (self.progress.setValue(value), self.status.setText(text)))
        self.worker.finished.connect(self._done); self.worker.error.connect(self._error); self.worker.start()
    def _done(self, count, db_path, seconds):
        self.progress.setValue(100); self.status.setText(f"Готово: {count:,} строк за {seconds:.1f} сек".replace(",", " "))
        self.ready_badge.setText("✓ Отчёт сохранён для повторной сверки"); self.ready_badge.show()
        self.start.setText("Открыть готовую сверку"); self.start.setEnabled(True)
        self._open_reconciliation(db_path)

    def _open_reconciliation(self, db_path):
        """Открывает результат отдельно: главное окно остаётся рабочим."""
        owner = self.parentWidget()
        dialog = InvestorReconciliationDialog(self.repo, db_path, owner or self)
        dialog.setModal(False)
        if owner is not None:
            owner._reconciliation_dialog = dialog
        else:
            self._reconciliation_dialog = dialog
        dialog.show(); dialog.raise_(); dialog.activateWindow()
        self.hide()
    def _error(self, message):
        self.start.setEnabled(True); self.status.setText("Ошибка подготовки отчёта"); QMessageBox.critical(self, "Ошибка импорта", message)

class ReconciliationPeriodDialog(QDialog):
    """Необязательные независимые периоды для сверки АСОП и ВТБ."""
    def __init__(self, asop_dates=None, vtb_dates=None, parent=None):
        super().__init__(parent); self.setWindowTitle("Периоды сверки"); self.setFixedSize(530, 230)
        self.setStyleSheet("QDialog { background:#F5F7FB; color:#172033; } QLabel { color:#172033; } QDateEdit { background:#FFFFFF; color:#172033; border:1px solid #D8E0EB; border-radius:8px; padding:5px; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(12)
        title = QLabel("Ограничить периоды сверки"); title.setStyleSheet("font-size:20px; font-weight:750;"); lay.addWidget(title)
        note = QLabel("Периоды независимы: АСОП фильтруется по дате поездки, ВТБ — по дате оплаты на терминале."); note.setWordWrap(True); note.setStyleSheet("color:#536174;"); lay.addWidget(note)
        default_from = QDate.currentDate().addMonths(-1); default_to = QDate.currentDate()
        self.asop_check, self.asop_from, self.asop_to = self._date_row(lay, "Период АСОП", asop_dates, default_from, default_to)
        self.vtb_check, self.vtb_from, self.vtb_to = self._date_row(lay, "Период ВТБ", vtb_dates, default_from, default_to)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel); buttons.button(QDialogButtonBox.Ok).setText("Применить"); buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

    def _date_row(self, layout, label, value, default_from, default_to):
        row = QHBoxLayout(); check = QCheckBox(label); check.setChecked(bool(value)); row.addWidget(check)
        date_from = QDateEdit(); date_from.setCalendarPopup(True); date_from.setDisplayFormat("dd.MM.yyyy"); date_from.setDate(QDate.fromString(value[0], "yyyy-MM-dd") if value else default_from); setup_russian_calendar(date_from)
        date_to = QDateEdit(); date_to.setCalendarPopup(True); date_to.setDisplayFormat("dd.MM.yyyy"); date_to.setDate(QDate.fromString(value[1], "yyyy-MM-dd") if value else default_to); setup_russian_calendar(date_to)
        row.addWidget(QLabel("с")); row.addWidget(date_from); row.addWidget(QLabel("по")); row.addWidget(date_to); row.addStretch(); layout.addLayout(row)
        def toggle(checked): date_from.setEnabled(checked); date_to.setEnabled(checked)
        check.toggled.connect(toggle); toggle(check.isChecked())
        return check, date_from, date_to

    def periods(self):
        def value(check, start, end):
            return (start.date().toString("yyyy-MM-dd"), end.date().toString("yyyy-MM-dd")) if check.isChecked() else None
        return value(self.asop_check, self.asop_from, self.asop_to), value(self.vtb_check, self.vtb_from, self.vtb_to)

class InvestorReconciliationDialog(QDialog):
    """Сверка АСОП безнал/без льготы с отчётом инвестора по перевозчикам."""
    def __init__(self, repo, investor_db_path, parent=None, asop_dates=None, vtb_dates=None):
        super().__init__(parent); self.repo = repo; self.investor_db_path = investor_db_path; self.asop_dates = asop_dates; self.vtb_dates = vtb_dates; self.setWindowTitle("Сверка с ВТБ"); self.resize(1320, 700)
        self.rows = repo.get_investor_reconciliation(investor_db_path, asop_dates, vtb_dates)
        self.setStyleSheet("""
            QDialog { background:#F5F6F8; color:#1D1D1F; }
            QLabel { color:#1D1D1F; }
            QScrollArea, QScrollArea::viewport { background:#F5F6F8; border:none; }
            QPushButton { color:#1D1D1F; background:#FFFFFF; border:1px solid #D8DEE9; border-radius:8px; padding:8px 12px; font-weight:650; }
            QPushButton:hover { color:#1D1D1F; background:#F0F6FF; border-color:#7EA7F8; }
            QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; border-radius:10px; gridline-color:#EEF0F4; }
            QTableWidget::item { color:#1D1D1F; }
            QHeaderView::section { background:#FFFFFF; color:#6B7280; border:none; border-bottom:1px solid #D8DEE9; padding:8px; font-weight:600; }
        """)
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Сверка АСОП и отчёта ВТБ"); title.setStyleSheet("font-size:24px; font-weight:700;"); lay.addWidget(title)
        note = QLabel("В АСОП учитываются: пустой реализатор, безналичная форма оплаты, категория «Без льготы» и все статусы. Сопоставление — строго по ID транзакции.")
        note.setWordWrap(True); note.setStyleSheet("color:#6B7280;"); lay.addWidget(note)
        asop_unique = sum(item["only_asop"] for item in self.rows); vtb_unique = sum(item["only_investor"] for item in self.rows); unique_ids = asop_unique + vtb_unique
        metrics = self.rows[0] if self.rows else {}
        matched_ids = metrics.get("matched_ids", 0); repeated_asop = metrics.get("repeated_asop_ids", 0); repeated_vtb = metrics.get("repeated_vtb_ids", 0)
        summary = QLabel(f"Совпавшие ID: {matched_ids:,}   ·   Уникальные: АСОП {asop_unique:,} / ВТБ {vtb_unique:,}   ·   Повторы: АСОП {repeated_asop:,} / ВТБ {repeated_vtb:,}".replace(",", " "))
        summary.setStyleSheet("background:#FFFFFF; border:1px solid #D8DEE9; border-radius:10px; padding:10px 12px; font-weight:700;"); lay.addWidget(summary)
        cards_scroll = QScrollArea(); cards_scroll.setWidgetResizable(True); cards_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded); cards_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff); cards_scroll.setFixedHeight(128); cards_scroll.setFrameShape(QFrame.NoFrame); cards_scroll.setStyleSheet("QScrollArea, QScrollArea::viewport { background:#F5F6F8; border:none; }")
        cards_host = QWidget(); cards_host.setStyleSheet("background:#F5F6F8;"); cards_lay = QHBoxLayout(cards_host); cards_lay.setContentsMargins(0, 0, 0, 4); cards_lay.setSpacing(10)
        for item in self.rows:
            card = QFrame(); card.setFixedWidth(290); card.setStyleSheet("QFrame { background:#FFFFFF; border:1px solid #D8DEE9; border-radius:12px; } QLabel { border:none; }")
            card_lay = QVBoxLayout(card); card_lay.setContentsMargins(12, 10, 12, 10); card_lay.setSpacing(4)
            carrier_label = QLabel(item["carrier"]); carrier_label.setToolTip(item["carrier"]); carrier_label.setStyleSheet("font-weight:750; color:#172033;"); card_lay.addWidget(carrier_label)
            organization_label = QLabel(f"ВТБ · {item['organization']}"); organization_label.setToolTip(item["organization"]); organization_label.setStyleSheet("color:#536174; font-size:12px;"); card_lay.addWidget(organization_label)
            figures = QLabel(f"Уникально · АСОП: {item['only_asop']:,}   ·   ВТБ: {item['only_investor']:,}".replace(",", " ")); figures.setStyleSheet("color:#1769E0; font-size:12px; font-weight:650;"); card_lay.addWidget(figures)
            unique = item["only_asop"] + item["only_investor"]
            mismatch = QLabel(f"Уникальных ID: {unique:,}".replace(",", " ")); mismatch.setStyleSheet("color:#B42318; font-weight:750;" if unique else "color:#16794C; font-weight:750;"); card_lay.addWidget(mismatch)
            cards_lay.addWidget(card)
        cards_lay.addStretch(); cards_scroll.setWidget(cards_host); lay.addWidget(cards_scroll)
        headers = ["Перевозчик АСОП", "Организация ВТБ", "АСОП, операций", "ВТБ, операций", "АСОП, сумма", "ВТБ, сумма", "Уникально в АСОП", "Уникально у ВТБ"]
        table = QTableWidget(0, len(headers)); self.table = table; table.setHorizontalHeaderLabels(headers); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setSelectionBehavior(QAbstractItemView.SelectRows); table.setSortingEnabled(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive); table.horizontalHeader().resizeSection(0, 260)
        # Название организации — один из главных ориентиров сверки, поэтому
        # не отдаём его растягиванию среди числовых колонок.
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive); table.horizontalHeader().resizeSection(1, 340)
        for col in range(2, len(headers)): table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Interactive)
        table.setSortingEnabled(False)
        for index, item in enumerate(self.rows):
            table.insertRow(index)
            values = [item["carrier"], item["organization"], f"{item['asop_operations']:,}".replace(",", " "), f"{item['investor_operations']:,}".replace(",", " "), f"{item['asop_amount']:,.2f} ₽".replace(",", " ").replace(".", ","), f"{item['investor_amount']:,.2f} ₽".replace(",", " ").replace(".", ","), f"{item['only_asop']:,}".replace(",", " "), f"{item['only_investor']:,}".replace(",", " ")]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if col in (0, 1): cell.setToolTip(value)
                if col >= 2: cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(index, col, cell)
        table.resizeColumnsToContents(); table.horizontalHeader().resizeSection(0, 260); table.horizontalHeader().resizeSection(1, 340)
        table.setSortingEnabled(True); table.cellDoubleClicked.connect(lambda *_: self._open_selected_unique_report()); lay.addWidget(table, 1)
        actions = QHBoxLayout(); periods_button = QPushButton("Периоды АСОП / ВТБ…"); periods_button.clicked.connect(self._edit_periods); actions.addWidget(periods_button)
        statuses_button = QPushButton("Статусы АСОП")
        statuses_button.setToolTip("Показывает число уникальных статусов и количество операций в каждом.")
        statuses_button.clicked.connect(self._show_asop_statuses); actions.addWidget(statuses_button)
        daily_ids_button = QPushButton("Транзакции по датам")
        daily_ids_button.setToolTip("Показывает уникальные ID АСОП по датам поездки и результат их сверки с ВТБ.")
        daily_ids_button.clicked.connect(self._show_daily_transactions); actions.addWidget(daily_ids_button)
        report_button = QPushButton("Открыть отчёт ВТБ")
        report_button.setToolTip("Выберите строку или дважды кликните по ней, чтобы увидеть исходные операции ВТБ.")
        report_button.clicked.connect(self._open_investor_report); actions.addWidget(report_button)
        asop_button = QPushButton("Открыть данные АСОП")
        asop_button.setToolTip("Показывает все исходные поля АСОП для выбранного перевозчика.")
        asop_button.clicked.connect(self._open_asop_report); actions.addWidget(asop_button)
        daily_button = QPushButton("По дням АСОП")
        daily_button.setToolTip("Суммарное количество уникальных операций АСОП выбранного перевозчика по датам поездки.")
        daily_button.clicked.connect(self._open_asop_daily); actions.addWidget(daily_button)
        mismatch_button = QPushButton("Показать уникальные ID")
        mismatch_button.setStyleSheet("QPushButton { background:#FFF1F2; color:#B42318; border:1px solid #FECDD3; border-radius:8px; padding:8px 12px; font-weight:700; } QPushButton:hover { background:#FFE4E6; }")
        # Детализация уникальных ID — полный контрольный перечень. Он не
        # зависит от ручных периодов карточек сверки и всегда показывает
        # все расхождения в подготовленных источниках.
        mismatch_button.clicked.connect(lambda: InvestorMismatchDialog(repo, investor_db_path, self).exec()); actions.addWidget(mismatch_button); actions.addStretch(); lay.addLayout(actions)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

    def _edit_periods(self):
        dialog = ReconciliationPeriodDialog(self.asop_dates, self.vtb_dates, self)
        if dialog.exec() == QDialog.Accepted:
            asop_dates, vtb_dates = dialog.periods()
            owner = self.parentWidget()
            replacement = InvestorReconciliationDialog(self.repo, self.investor_db_path, owner, asop_dates, vtb_dates)
            replacement.setModal(False)
            if owner is not None:
                owner._reconciliation_dialog = replacement
            else:
                self._replacement_dialog = replacement
            replacement.show(); replacement.raise_(); replacement.activateWindow()
            self.close()

    def _show_asop_statuses(self):
        AsopStatusComparisonDialog(self.repo, self.investor_db_path, self.asop_dates, self.vtb_dates, self).exec()

    def _show_daily_transactions(self):
        AsopDailyTransactionDialog(self.repo, self.investor_db_path, self.asop_dates, self.vtb_dates, self).exec()

    def _open_selected_unique_report(self):
        """Двойной клик открывает именно ту сторону, где есть уникальные ID."""
        row = self.table.currentRow()
        if row < 0: return
        def count_at(column):
            item = self.table.item(row, column)
            try: return int((item.text() if item else "0").replace(" ", ""))
            except ValueError: return 0
        if count_at(6) > 0:
            self._open_asop_report()
        elif count_at(7) > 0:
            self._open_investor_report()

    def _open_investor_report(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Отчёт инвестора", "Выберите строку в таблице сверки.")
            return
        organization = self.table.item(row, 1).text() if self.table.item(row, 1) else ""
        if not organization or organization == "Нет в отчёте ВТБ":
            QMessageBox.information(self, "Нет строк ВТБ", "Для этой строки в отчёте ВТБ операций нет.")
            return
        InvestorReportDialog(self.repo, self.investor_db_path, organization, self, self.asop_dates, self.vtb_dates).exec()

    def _open_asop_report(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Данные АСОП", "Выберите строку в таблице сверки.")
            return
        carrier = self.table.item(row, 0).text() if self.table.item(row, 0) else ""
        if not carrier or carrier == "Нет в АСОП":
            QMessageBox.information(self, "Нет строк АСОП", "Для этой строки в АСОП операций нет.")
            return
        AsopReportDialog(self.repo, carrier, self.investor_db_path, self, self.asop_dates, self.vtb_dates).exec()

    def _open_asop_daily(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "По дням АСОП", "Выберите перевозчика в таблице сверки.")
            return
        carrier = self.table.item(row, 0).text() if self.table.item(row, 0) else ""
        if not carrier or carrier == "Нет в АСОП":
            QMessageBox.information(self, "Нет данных АСОП", "Для этой строки уникальных операций АСОП нет.")
            return
        CarrierDailyDialog(self.repo, carrier, self.investor_db_path, self.asop_dates, self.vtb_dates, self).exec()

class InvestorReportDialog(QDialog):
    """Просмотр исходных строк подготовленного отчёта инвестора."""
    def __init__(self, repo, investor_db_path, organization, parent=None, asop_dates=None, vtb_dates=None):
        super().__init__(parent); self.repo = repo; self.investor_db_path = investor_db_path; self.organization = organization; self.asop_dates = asop_dates; self.vtb_dates = vtb_dates
        self.setWindowTitle("Отчёт инвестора"); self.resize(1260, 710)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QLineEdit, QTableView { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; border-radius:9px; padding:7px; } QTableView { padding:0; alternate-background-color:#F8FAFC; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Исходный отчёт инвестора"); title.setStyleSheet("font-size:23px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        org = QLabel(organization); org.setWordWrap(True); org.setStyleSheet("color:#1769E0; font-size:15px; font-weight:700;"); lay.addWidget(org)
        note = QLabel("Показаны только операции с ID, уникальными у ВТБ. Поиск работает по ID транзакции; для скорости за один раз выводится до 5 000 строк.")
        note.setWordWrap(True); note.setStyleSheet("color:#6B7280;"); lay.addWidget(note)
        controls = QHBoxLayout(); self.search = QLineEdit(); self.search.setPlaceholderText("Найти по ID транзакции…"); self.search.setClearButtonEnabled(True); self.search.returnPressed.connect(self._load); controls.addWidget(self.search, 1)
        load = QPushButton("Найти"); load.clicked.connect(self._load); controls.addWidget(load); lay.addLayout(controls)
        self.status = QLabel(); self.status.setStyleSheet("color:#6B7280; font-size:12px;"); lay.addWidget(self.status)
        self.table = QTableView(); self.table.setAlternatingRowColors(True); self.table.setSelectionBehavior(QAbstractItemView.SelectRows); self.table.setEditTriggers(QAbstractItemView.NoEditTriggers); self.table.setShowGrid(False); self.table.verticalHeader().setVisible(False); self.table.clicked.connect(self._preview_selected); lay.addWidget(self.table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)
        self._load()

    def _load(self):
        data = self.repo.get_investor_report_rows(self.investor_db_path, self.organization, self.search.text(), asop_dates=self.asop_dates, vtb_dates=self.vtb_dates)
        self.model = PandasTableModel(data); self.table.setModel(self.model)
        self.table.resizeColumnsToContents(); self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive); self.table.horizontalHeader().setStretchLastSection(False)
        self.status.setText(f"Показано строк: {len(data):,}".replace(",", " "))

    def _preview_selected(self, index):
        if not index.isValid() or not hasattr(self, "model") or index.row() >= len(self.model._df): return
        if getattr(self, "_preview_dialog", None): self._preview_dialog.close()
        self._preview_dialog = TransactionPreviewDialog(self.model._df.iloc[index.row()].to_dict(), "Предпросмотр операции ВТБ", self)
        self._preview_dialog.show(); self._preview_dialog.raise_()

class AsopReportDialog(QDialog):
    """Полный первичный реестр АСОП по выбранной строке сверки."""
    def __init__(self, repo, carrier, investor_db_path, parent=None, asop_dates=None, vtb_dates=None):
        super().__init__(parent); self.repo = repo; self.carrier = carrier; self.investor_db_path = investor_db_path; self.asop_dates = asop_dates; self.vtb_dates = vtb_dates
        self.setWindowTitle("Данные АСОП"); self.resize(1320, 720)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QLabel { color:#1D1D1F; } QLineEdit, QTableView { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; border-radius:9px; padding:7px; } QTableView { padding:0; alternate-background-color:#F8FAFC; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Исходные данные АСОП"); title.setStyleSheet("font-size:23px; font-weight:700; color:#1D1D1F;"); lay.addWidget(title)
        carrier_label = QLabel(carrier); carrier_label.setWordWrap(True); carrier_label.setStyleSheet("color:#1769E0; font-size:15px; font-weight:700;"); lay.addWidget(carrier_label)
        note = QLabel("Все поля исходного реестра только для ID, уникальных в АСОП. Учитываются операции с любыми статусами; поиск — по ID транзакции.")
        note.setWordWrap(True); note.setStyleSheet("color:#536174;"); lay.addWidget(note)
        controls = QHBoxLayout(); self.search = QLineEdit(); self.search.setPlaceholderText("Найти по ID транзакции…"); self.search.setClearButtonEnabled(True); self.search.returnPressed.connect(self._load); controls.addWidget(self.search, 1)
        load = QPushButton("Найти"); load.clicked.connect(self._load); controls.addWidget(load); lay.addLayout(controls)
        self.status = QLabel(); self.status.setStyleSheet("color:#536174; font-size:12px;"); lay.addWidget(self.status)
        self.table = QTableView(); self.table.setAlternatingRowColors(True); self.table.setSelectionBehavior(QAbstractItemView.SelectRows); self.table.setEditTriggers(QAbstractItemView.NoEditTriggers); self.table.setShowGrid(False); self.table.verticalHeader().setVisible(False); self.table.clicked.connect(self._preview_selected); lay.addWidget(self.table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)
        self._load()

    def _load(self):
        data = self.repo.get_asop_reconciliation_rows(self.carrier, self.investor_db_path, self.search.text(), asop_dates=self.asop_dates, vtb_dates=self.vtb_dates)
        self.model = PandasTableModel(data); self.table.setModel(self.model)
        self.table.resizeColumnsToContents(); self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive); self.table.horizontalHeader().setStretchLastSection(False)
        self.status.setText(f"Показано строк: {len(data):,}".replace(",", " "))

    def _preview_selected(self, index):
        if not index.isValid() or not hasattr(self, "model") or index.row() >= len(self.model._df): return
        if getattr(self, "_preview_dialog", None): self._preview_dialog.close()
        self._preview_dialog = TransactionPreviewDialog(self.model._df.iloc[index.row()].to_dict(), "Предпросмотр операции АСОП", self)
        self._preview_dialog.show(); self._preview_dialog.raise_()

class CarrierDailyDialog(QDialog):
    """Суммарная дневная таблица уникальных операций АСОП по перевозчику."""
    def __init__(self, repo, carrier, investor_db_path, asop_dates=None, vtb_dates=None, parent=None):
        super().__init__(parent); self.setWindowTitle("Уникальные операции АСОП по дням"); self.resize(540, 600)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; border-radius:10px; } QHeaderView::section { background:#FFFFFF; color:#536174; border:none; border-bottom:1px solid #D8DEE9; padding:8px; font-weight:700; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Уникальные операции АСОП по дням"); title.setStyleSheet("font-size:21px; font-weight:750;"); lay.addWidget(title)
        label = QLabel(carrier); label.setWordWrap(True); label.setStyleSheet("color:#1769E0; font-weight:700;"); lay.addWidget(label)
        rows = repo.get_asop_unique_daily(carrier, investor_db_path, asop_dates, vtb_dates)
        total = sum(count for _, count in rows); note = QLabel(f"Даты поездок · Всего уникальных операций: {total:,}".replace(",", " ")); note.setStyleSheet("color:#536174;"); lay.addWidget(note)
        table = QTableWidget(0, 2); table.setHorizontalHeaderLabels(["Дата поездки", "Уникальных транзакций"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        for index, (day, count) in enumerate(rows):
            table.insertRow(index); table.setItem(index, 0, QTableWidgetItem(day)); cell = QTableWidgetItem(f"{count:,}".replace(",", " ")); cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter); table.setItem(index, 1, cell)
        lay.addWidget(table, 1); buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class AsopStatusComparisonDialog(QDialog):
    """Сопоставление ID ВТБ для каждого статуса АСОП."""
    def __init__(self, repo, investor_db_path, asop_dates=None, vtb_dates=None, parent=None):
        super().__init__(parent); self.setWindowTitle("Статусы АСОП и ВТБ"); self.resize(720, 520)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; border-radius:10px; } QHeaderView::section { background:#FFFFFF; color:#536174; border:0; border-bottom:1px solid #D8DEE9; padding:8px; font-weight:700; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Сравнение статусов АСОП с ВТБ"); title.setStyleSheet("font-size:21px; font-weight:750;"); lay.addWidget(title)
        note = QLabel("Для каждого статуса АСОП показано число уникальных ID, найденных в ВТБ, и число ID, отсутствующих в ВТБ."); note.setWordWrap(True); note.setStyleSheet("color:#536174;"); lay.addWidget(note)
        rows = repo.get_asop_status_comparison(investor_db_path, asop_dates, vtb_dates)
        total_matched = sum(row[1] for row in rows); total_unique = sum(row[2] for row in rows)
        summary = QLabel(f"Уникальных статусов: {len(rows)} · Совпадений ID: {total_matched:,} · Уникальных ID: {total_unique:,}".replace(",", " ")); summary.setStyleSheet("color:#1769E0; font-weight:700;"); lay.addWidget(summary)
        table = QTableWidget(0, 3); table.setHorizontalHeaderLabels(["Статус АСОП", "Совпадений ID", "Уникальных ID"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setSortingEnabled(True)
        table.setSortingEnabled(False)
        for index, (status, matched, unique) in enumerate(rows):
            table.insertRow(index); table.setItem(index, 0, QTableWidgetItem(status))
            for column, value in ((1, matched), (2, unique)):
                cell = QTableWidgetItem(f"{value:,}".replace(",", " ")); cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter); table.setItem(index, column, cell)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents); table.setSortingEnabled(True); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class AsopDailyTransactionDialog(QDialog):
    """Количество уникальных ID АСОП по датам поездки с результатом сверки."""
    def __init__(self, repo, investor_db_path, asop_dates=None, vtb_dates=None, parent=None):
        super().__init__(parent); self.setWindowTitle("Транзакции по датам"); self.resize(820, 600)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #D8DEE9; border-radius:10px; } QHeaderView::section { background:#FFFFFF; color:#536174; border:0; border-bottom:1px solid #D8DEE9; padding:8px; font-weight:700; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Транзакции по датам поездки"); title.setStyleSheet("font-size:21px; font-weight:750;"); lay.addWidget(title)
        note = QLabel("Подсчёт ведётся по уникальным ID транзакций АСОП. Для каждой даты виден результат сопоставления с ВТБ."); note.setWordWrap(True); note.setStyleSheet("color:#536174;"); lay.addWidget(note)
        rows = repo.get_asop_daily_id_comparison(investor_db_path, asop_dates, vtb_dates)
        table = QTableWidget(0, 4); table.setHorizontalHeaderLabels(["Дата поездки", "Уникальных ID", "Совпадений с ВТБ", "Уникальных АСОП"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setSortingEnabled(True)
        table.setSortingEnabled(False)
        for index, (day, all_ids, matched, unique) in enumerate(rows):
            table.insertRow(index); table.setItem(index, 0, QTableWidgetItem(day))
            for column, value in ((1, all_ids), (2, matched), (3, unique)):
                cell = QTableWidgetItem(f"{value:,}".replace(",", " ")); cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter); table.setItem(index, column, cell)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        for column in (1, 2, 3): table.horizontalHeader().setSectionResizeMode(column, QHeaderView.ResizeToContents)
        table.setSortingEnabled(True); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

class InvestorMismatchDialog(QDialog):
    def __init__(self, repo, investor_db_path, parent=None, asop_dates=None, vtb_dates=None):
        super().__init__(parent); self.setWindowTitle("Уникальные ID транзакций"); self.resize(1360, 680)
        self.repo = repo; self.investor_db_path = investor_db_path
        self.rows = repo.get_investor_mismatches(investor_db_path, asop_dates=asop_dates, vtb_dates=vtb_dates)
        self.setStyleSheet("QDialog { background:#F5F6F8; color:#1D1D1F; } QTableWidget { background:#FFFFFF; color:#1D1D1F; border:1px solid #FECACA; border-radius:10px; gridline-color:#FEE2E2; } QHeaderView::section { background:#FFF7F7; color:#B42318; border:none; border-bottom:1px solid #FECACA; padding:8px; font-weight:700; }")
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        title = QLabel("Уникальные ID транзакций"); title.setStyleSheet("font-size:22px; font-weight:700; color:#B42318;"); lay.addWidget(title)
        note = QLabel("Показаны первые 10 000 ID для быстрого просмотра. Экспорт выгружает полный список, без лимита, и разбивает его на листы Excel при необходимости. Список не зависит от выбранных периодов сверки.")
        note.setWordWrap(True); note.setStyleSheet("color:#6B7280;"); lay.addWidget(note)
        headers = ["Где уникальна", "ID транзакции", "Перевозчик", "Организация ВТБ", "Дата поездки АСОП", "Дата обработки АСОП", "Дата оплаты ВТБ", "АСОП, сумма", "ВТБ, сумма"]
        table = QTableWidget(0, len(headers)); table.setHorizontalHeaderLabels(headers); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setSortingEnabled(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch); table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        table.setSortingEnabled(False)
        for index, item in enumerate(self.rows):
            table.insertRow(index)
            values = [item["status"], item["transaction_id"], item["carrier"], item["organization"], item["asop_trip_date"], item["asop_server_datetime"], item["investor_payment_date"], f"{item['main_amount']:,.2f} ₽".replace(",", " ").replace(".", ","), f"{item['investor_amount']:,.2f} ₽".replace(",", " ").replace(".", ",")]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                cell.setBackground(QColor("#FFF1F2")); cell.setForeground(QColor("#7F1D1D") if col == 0 else QColor("#1D1D1F"))
                if col >= 7: cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(index, col, cell)
        table.resizeColumnsToContents()
        for col in range(len(headers)): table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Interactive)
        table.setSortingEnabled(True); lay.addWidget(table, 1)
        actions = QHBoxLayout(); export = QPushButton("Экспортировать в Excel…"); export.clicked.connect(self._export); actions.addWidget(export); actions.addStretch(); lay.addLayout(actions)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.reject); lay.addWidget(buttons)

    def _export(self):
        path, _ = QFileDialog.getSaveFileName(self, "Экспорт уникальных ID", "unique_transaction_ids.xlsx", "Excel (*.xlsx)")
        if not path: return
        if not path.lower().endswith(".xlsx"): path += ".xlsx"
        try:
            headers = ["Где уникальна", "ID транзакции", "Перевозчик АСОП", "Организация ВТБ", "Дата поездки АСОП", "Дата обработки АСОП", "Дата оплаты ВТБ", "АСОП, сумма", "ВТБ, сумма"]
            fields = ["status", "transaction_id", "carrier", "organization", "asop_trip_date", "asop_server_datetime", "investor_payment_date", "main_amount", "investor_amount"]
            workbook = openpyxl.Workbook(write_only=True)
            sheet_number = 0; sheet_rows = 0; total = 0; page_size = 100000; offset = 0; sheet = None
            while True:
                page = self.repo.get_investor_mismatches(self.investor_db_path, limit=page_size, offset=offset)
                if not page: break
                for item in page:
                    if sheet is None or sheet_rows >= 1_048_575:
                        sheet_number += 1; sheet = workbook.create_sheet(f"Уникальные ID {sheet_number}"); sheet.append(headers); sheet_rows = 0
                    sheet.append([item.get(field, "") for field in fields]); sheet_rows += 1; total += 1
                offset += page_size // 2
                QApplication.processEvents()
                if len(page) < page_size: break
            if not total:
                QMessageBox.information(self, "Экспорт", "Нет уникальных ID для экспорта.")
                return
            workbook.save(path)
            QMessageBox.information(self, "Экспорт готов", f"Сохранено строк: {total:,} · листов: {sheet_number}\n{path}".replace(",", " "))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка экспорта", f"Не удалось сохранить файл:\n{e}")

class FilterColumnsDialog(QDialog):
    def __init__(self, all_cols, current_cols, parent=None):
        super().__init__(parent); self.setWindowTitle("Настройка фильтров"); self.setFixedSize(400, 500)
        lay = QVBoxLayout(self); lay.setContentsMargins(16, 16, 16, 16); lay.setSpacing(8)
        info = QLabel("Выберите колонки для фильтров (макс. 8):"); lay.addWidget(info)
        self.search = QLineEdit(); self.search.setPlaceholderText("Поиск колонки..."); self.search.setFixedHeight(26)
        self.search.textChanged.connect(self._filter); lay.addWidget(self.search)
        scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.NoFrame)
        w = QWidget(); self.cl = QVBoxLayout(w); self.cl.setContentsMargins(0, 0, 0, 0); self.cl.setSpacing(2)
        scroll.setWidget(w); lay.addWidget(scroll); self.vars = {}
        for c in all_cols:
            cb = QCheckBox(c); cb.setChecked(c in current_cols); self.vars[c] = cb; self.cl.addWidget(cb)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept); btns.rejected.connect(self.reject); lay.addWidget(btns)
    def _filter(self, t):
        t = t.lower()
        for c, cb in self.vars.items(): cb.setVisible(t in c.lower())
    def result(self): return [c for c, cb in self.vars.items() if cb.isChecked()]

class PandasTableModel(QAbstractTableModel):
    def __init__(self, df, parent=None):
        super().__init__(parent); self._df = df; self._headers = list(df.columns)
    def rowCount(self, parent=None): return len(self._df)
    def columnCount(self, parent=None): return len(self._headers)
    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or role != Qt.DisplayRole: return None
        val = self._df.iloc[index.row(), index.column()]
        if pd.isna(val): return ""
        return str(val)
    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole: return self._headers[section]
        return None

class TransactionPreviewDialog(QDialog):
    """Быстрый просмотр одной исходной операции со всеми без исключения полями."""
    def __init__(self, record, title, parent=None):
        super().__init__(parent); self.setWindowTitle(title); self.resize(780, 620); self.setModal(False)
        self.setStyleSheet("""
            QDialog { background:#F5F7FB; color:#172033; }
            QTableWidget { background:#FFFFFF; color:#172033; border:1px solid #D8E0EB; border-radius:10px; }
            QHeaderView::section { background:#F8FAFC; color:#536174; border:0; border-bottom:1px solid #E2E8F0; padding:8px; font-weight:700; }
        """)
        lay = QVBoxLayout(self); lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(10)
        heading = QLabel(title); heading.setStyleSheet("font-size:20px; font-weight:750; color:#172033;"); lay.addWidget(heading)
        note = QLabel("Все поля выбранной исходной записи. Нажмите другую строку в отчёте, чтобы сменить предпросмотр.")
        note.setStyleSheet("color:#536174;"); note.setWordWrap(True); lay.addWidget(note)
        table = QTableWidget(0, 2); table.setHorizontalHeaderLabels(["Поле", "Значение"]); table.verticalHeader().setVisible(False); table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        for row, (key, value) in enumerate(record.items()):
            table.insertRow(row)
            display = "" if pd.isna(value) else str(value)
            key_item = QTableWidgetItem(str(key)); key_item.setForeground(QColor("#536174"))
            table.setItem(row, 0, key_item); table.setItem(row, 1, QTableWidgetItem(display))
        table.resizeRowsToContents(); lay.addWidget(table, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Close); buttons.rejected.connect(self.close); lay.addWidget(buttons)

class TableCopyController(QObject):
    """Копирование значений из любой таблицы приложения через ⌘C или меню."""
    @staticmethod
    def _table_for(widget):
        current = widget
        while current is not None:
            if isinstance(current, QAbstractItemView): return current
            current = current.parent()
        return None

    @staticmethod
    def _copy_indexes(table, indexes):
        if not indexes: return False
        rows = sorted({index.row() for index in indexes})
        columns = sorted({index.column() for index in indexes})
        index_set = {(index.row(), index.column()) for index in indexes}
        lines = []
        for row in rows:
            values = []
            for column in columns:
                if (row, column) not in index_set:
                    values.append("")
                    continue
                value = table.model().data(table.model().index(row, column), Qt.DisplayRole)
                values.append("" if value is None else str(value))
            lines.append("\t".join(values))
        QApplication.clipboard().setText("\n".join(lines))
        return True

    def eventFilter(self, watched, event):
        table = self._table_for(watched)
        if table is None:
            return False
        is_copy_shortcut = (
            event.type() == QEvent.KeyPress
            and event.key() == Qt.Key_C
            and bool(event.modifiers() & (Qt.ControlModifier | Qt.MetaModifier))
        )
        if is_copy_shortcut:
            return self._copy_indexes(table, table.selectionModel().selectedIndexes())
        if event.type() == QEvent.ContextMenu:
            point = table.viewport().mapFromGlobal(event.globalPos())
            index = table.indexAt(point)
            if not index.isValid(): return False
            menu = QMenu(table)
            action = menu.addAction("Копировать значение")
            action.triggered.connect(lambda: self._copy_indexes(table, [index]))
            menu.exec(event.globalPos())
            return True
        return False

class DetailsViewer(QDialog):
    def __init__(self, repo, carrier, where_clause, parent=None):
        super().__init__(parent); self.setWindowTitle(f"Детали: {carrier}"); self.resize(1200, 700); self.setModal(True)
        self.repo = repo; self.carrier = carrier; self.where_clause = where_clause
        layout = QVBoxLayout(self); layout.setContentsMargins(12, 12, 12, 12); layout.setSpacing(8)
        info_label = QLabel(f"Записи для перевозчика: {carrier}")
        info_label.setFont(QFont(QFontDatabase.systemFont(QFontDatabase.GeneralFont).family(), Theme.Typography.Size.title, Theme.Typography.Weight.bold))
        layout.addWidget(info_label); self.table = QTableView(); self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setShowGrid(False); self.table.setAlternatingRowColors(True); self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True); self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._load_data(); layout.addWidget(self.table)
        btn_box = QDialogButtonBox(QDialogButtonBox.Close); btn_box.rejected.connect(self.reject); layout.addWidget(btn_box)
    def _load_data(self):
        carrier_escaped = self.carrier.replace("'", "''")
        full_where = f"{self.where_clause} AND \"{CARRIER_COLUMN}\" = '{carrier_escaped}'"
        query = f"SELECT * FROM raw_data WHERE {full_where} ORDER BY \"{DATETIME_COLUMN}\""
        try:
            df = self.repo._conn.execute(query).df()
            if "row_id" in df.columns: df = df.drop("row_id", axis=1)
            self.model = PandasTableModel(df); self.table.setModel(self.model)
            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить детали:\n{e}")

class AnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle(f"Анализатор реестра транзакций {APP_VERSION}")
        self.resize(1400, 900); self.setMinimumSize(1100, 700); self.repo = DuckDBRepository()
        QApplication.instance().setStyle(MacProxyStyle())
        self._stats_worker = None; self._db_loaded = False
        self._current_db_path = None; self._update_worker = None
        self._investor_db_path = None; self._investor_worker = None
        self.export_all_rows = False
        self._build_menu(); self._build_ui(); self._build_statusbar(); self._load_settings()
        self._apply_premium_theme()
        # Базу открывает только пользователь через меню: при старте не
        # подключаем предыдущий файл автоматически.
    def _apply_premium_theme(self):
        """Визуальный слой: не влияет на импорт, SQL, фильтры или экспорт."""
        self.setStyleSheet(f"""
            QMainWindow {{ background: {Theme.Colors.bg}; }}
            QMenuBar {{ background: {Theme.Colors.card_bg}; border-bottom: 1px solid {Theme.Colors.border}; padding: 4px 8px; }}
            QMenuBar::item {{ padding: 6px 10px; border-radius: 6px; }} QMenuBar::item:selected {{ background: #EAF3FF; }}
            QMenu {{ background: {Theme.Colors.card_bg}; border: 1px solid {Theme.Colors.border}; border-radius: 10px; padding: 6px; }}
            QMenu::item {{ padding: 7px 28px 7px 10px; border-radius: 6px; }} QMenu::item:selected {{ background: #EAF3FF; color: {Theme.Colors.blue}; }}
            QToolBar#main_toolbar {{ background: {Theme.Colors.card_bg}; border: none; border-bottom: 1px solid {Theme.Colors.border}; spacing: 7px; padding: 7px 14px; }}
            QToolButton {{ border: 1px solid transparent; border-radius: 8px; padding: 6px; }} QToolButton:hover {{ background: #EEF6FF; border-color: #D8EAFD; }} QToolButton:pressed {{ background: #DCEEFF; }}
            QStatusBar {{ background: {Theme.Colors.card_bg}; border-top: 1px solid {Theme.Colors.border}; color: {Theme.Colors.text_secondary}; padding: 2px 10px; }}
            QSplitter::handle {{ background: {Theme.Colors.border}; }} QSplitter::handle:hover {{ background: #BFDBFE; }}
            QTableView {{ color: #1D1D1F; background: {Theme.Colors.card_bg}; alternate-background-color: #FAFBFC; border: none; selection-background-color: #DCEEFF; selection-color: #1D1D1F; }}
            QTableView::item {{ color: #1D1D1F; border-bottom: 1px solid #EFF0F2; padding: 6px 10px; }} QTableView::item:selected {{ color: #1D1D1F; background: #DCEEFF; }} QTableView::item:hover {{ background: #F4F8FF; }}
            QHeaderView::section {{ background: {Theme.Colors.card_bg}; color: {Theme.Colors.text_secondary}; border: none; border-bottom: 1px solid {Theme.Colors.border}; padding: 8px 10px; font-weight: 600; }}
            QScrollBar:vertical {{ width: 10px; background: transparent; margin: 4px; }} QScrollBar::handle:vertical {{ background: #C7C7CC; min-height: 30px; border-radius: 5px; }} QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        """)
    def _auto_load(self):
        db = os.path.abspath(DB_FILE)
        if os.path.exists(db):
            try:
                self.repo.connect(db)
                self._current_db_path = db
                if not self.repo._table_exists():
                    self.repo.close(); return
                self._on_loaded()
                self.status_lbl.setText(f"БД загружена: {DB_FILE}")
            except Exception as e:
                print(f"DB load error: {e}")
    def _on_loaded(self):
        self._db_loaded = True; self.sidebar.clear_filters(); cols = self.repo.get_columns()
        defaults = ["Статус", "Форма оплаты проезда", "Категория пассажира", "Тип тарифа", "Реализатор", "Дата и время совершения поездки"]
        for c in defaults:
            if c in cols: self.sidebar.add_filter_section(c)
        if not self.sidebar.sections and cols:
            for c in cols[:4]: self.sidebar.add_filter_section(c)
        self.sidebar.set_repo(self.repo); self.table.model.set_where_clause("1=1"); self.table.update_info_label(); self._upd_stats()
    def _build_menu(self):
        mb = self.menuBar(); mb.setNativeMenuBar(True); fm = mb.addMenu("Файл")
        a = QAction("Открыть CSV...", self); a.setShortcut(QKeySequence.StandardKey.Open); a.triggered.connect(self.load_csv); fm.addAction(a)
        a = QAction("Открыть БД...", self); a.setShortcut(QKeySequence("Ctrl+Shift+O")); a.triggered.connect(self.load_db); fm.addAction(a)
        fm.addSeparator()
        export_menu = fm.addMenu("Экспорт в Excel")
        a = QAction("Сводная таблица...", self); a.setShortcut(QKeySequence("Ctrl+Shift+E")); a.triggered.connect(self.export_summary); export_menu.addAction(a)
        a = QAction("Исходные строки...", self); a.triggered.connect(self.export_raw_rows); export_menu.addAction(a)
        fm.addSeparator()
        a = QAction("Проверить обновления...", self); a.triggered.connect(self.check_updates); fm.addAction(a)
        fm.addSeparator()
        a = QAction("Выход", self); a.setShortcut(QKeySequence.StandardKey.Quit); a.triggered.connect(self.close); fm.addAction(a)
        tools = mb.addMenu("Инструменты")
        a = QAction("Настройка фильтров...", self); a.setShortcut(QKeySequence("Ctrl+,")); a.triggered.connect(self.open_filter_config); tools.addAction(a)
        a = QAction("Применить фильтры", self); a.setShortcut(QKeySequence("Ctrl+Return")); a.triggered.connect(self.run_filter); tools.addAction(a)
        a = QAction("Показать / скрыть фильтры", self); a.triggered.connect(self.toggle_sidebar); tools.addAction(a)
    def open_filter_config(self):
        if not self._db_loaded: QMessageBox.warning(self, "Нет данных", "Сначала загрузите CSV или БД."); return
        all_cols = self.repo.get_columns(); current = [col for _, (_, col) in self.sidebar.sections.items()]
        dlg = FilterColumnsDialog(all_cols, current, self)
        if dlg.exec() == QDialog.Accepted:
            selected = dlg.result()
            if not selected: QMessageBox.warning(self, "Нет выбора", "Выберите хотя бы одну колонку."); return
            if len(selected) > 8: QMessageBox.warning(self, "Слишком много", "Максимум 8 колонок."); return
            self.sidebar.clear_filters()
            for c in selected: self.sidebar.add_filter_section(c)
            self.sidebar.set_repo(self.repo); self._upd_stats()
    def _build_ui(self):
        central = QWidget(); self.setCentralWidget(central); ml = QVBoxLayout(central); ml.setContentsMargins(0, 0, 0, 0); ml.setSpacing(0)
        self.splitter = QSplitter(Qt.Horizontal); self.splitter.setHandleWidth(1); self.splitter.setOpaqueResize(True); self.splitter.setChildrenCollapsible(False)
        self.sidebar = SidebarFilters(self.repo); self.splitter.addWidget(self.sidebar); self.splitter.setCollapsible(0, False)
        self._filter_apply_timer = QTimer(self); self._filter_apply_timer.setSingleShot(True); self._filter_apply_timer.setInterval(180)
        self._filter_apply_timer.timeout.connect(self.run_filter)
        self.sidebar.filters_changed.connect(self._schedule_filter_apply)
        rw = QWidget(); rw.setObjectName("dashboardSurface")
        rl = QVBoxLayout(rw); rl.setContentsMargins(26, 20, 26, 20); rl.setSpacing(14)
        self.dashboard_header = DashboardHeader(); self.dashboard_header.reset_requested.connect(self.sidebar._on_clear_all); self.dashboard_header.charts_requested.connect(self.open_charts); self.dashboard_header.fleet_requested.connect(self.open_fleet_analytics); self.dashboard_header.operations_requested.connect(self.open_operational_analytics); self.dashboard_header.fraud_requested.connect(self.open_fraud_cards); self.dashboard_header.investor_requested.connect(self.open_investor_reconciliation); rl.addWidget(self.dashboard_header)
        self.summary = SummaryBar(); rl.addWidget(self.summary)
        table_card = QWidget(); table_card.setObjectName("tableCard")
        table_card.setStyleSheet(f"QWidget#tableCard {{ background: {Theme.Colors.card_bg}; border: 1px solid {Theme.Colors.border}; border-radius: 16px; }}")
        table_lay = QVBoxLayout(table_card); table_lay.setContentsMargins(0, 0, 0, 0); table_lay.setSpacing(0)
        table_title = QLabel("Сводная таблица")
        table_title.setStyleSheet(f"color: {Theme.Colors.text_primary}; font-size: 15px; font-weight: 700; padding: 12px 16px 8px 16px; border: none;")
        table_lay.addWidget(table_title)
        self.table = ResultsTableWidget(self.repo); table_lay.addWidget(self.table, 1)
        rl.addWidget(table_card, 1)
        self.splitter.addWidget(rw); self.splitter.setStretchFactor(0, 0); self.splitter.setStretchFactor(1, 1)
        self.splitter.setSizes([Theme.Sidebar.normal_width, 1080])
        ml.addWidget(self.splitter)
        self.loading = LoadingOverlay(central); self.table.tv.doubleClicked.connect(self._on_table_double_click)
    def _schedule_filter_apply(self):
        """Применяет выбор после короткой паузы, чтобы серия кликов не запускала лишние запросы."""
        if self._db_loaded:
            self._filter_apply_timer.start()
    def open_charts(self):
        if not self._db_loaded:
            QMessageBox.warning(self, "Нет данных", "Сначала загрузите CSV или БД.")
            return
        f, date_f = self.sidebar.get_filters()
        rows = self.repo.get_daily_paid_totals(self.repo.build_where_clause(f, date_f))
        ChartsDialog(rows, self).exec()
    def open_fleet_analytics(self):
        if not self._db_loaded:
            QMessageBox.warning(self, "Нет данных", "Сначала загрузите CSV или БД.")
            return
        filters, date_filters = self.sidebar.get_filters()
        FleetAnalyticsDialog(self.repo, self.repo.build_where_clause(filters, date_filters), self).exec()
    def open_operational_analytics(self):
        if not self._db_loaded:
            QMessageBox.warning(self, "Нет данных", "Сначала загрузите CSV или БД.")
            return
        filters, date_filters = self.sidebar.get_filters()
        OperationalAnalyticsDialog(self.repo, self.repo.build_where_clause(filters, date_filters), self).exec()
    def open_fraud_cards(self):
        if not self._db_loaded:
            QMessageBox.warning(self, "Нет данных", "Сначала загрузите CSV или БД.")
            return
        required = {"Номер обслуженной карты", "Форма оплаты проезда"}
        if not required.issubset(self.repo.get_columns()):
            QMessageBox.warning(self, "Недостаточно данных", "Для раздела «ФРОД» нужны колонки «Номер обслуженной карты» и «Форма оплаты проезда».")
            return
        # Витрина использует текущие фильтры, в том числе выбранную форму
        # оплаты, и дополнительно оставляет только безналичные операции.
        filters, date_filters = self.sidebar.get_filters()
        FraudCardsDialog(self.repo, self.repo.build_where_clause(filters, date_filters), self).exec()
    def open_investor_reconciliation(self):
        if not self._db_loaded:
            QMessageBox.warning(self, "Нет данных", "Сначала загрузите основной реестр.")
            return
        source_columns = self.repo.investor_source_columns()
        missing = [label for key, label in (("transaction_id", "ID транзакции"), ("realizator", "Реализатор"), ("payment", "Форма оплаты проезда"), ("category", "Категория пассажира"), ("carrier", "Перевозчик")) if not source_columns.get(key)]
        if missing:
            QMessageBox.warning(self, "Недостаточно данных", "Не найдены колонки: " + ", ".join(missing) + ".\nОткройте БД, созданную из полного реестра транзакций.")
            return
        self._investor_load_dialog = InvestorLoadDialog(self.repo, self)
        self._investor_load_dialog.setModal(False)
        self._investor_load_dialog.show(); self._investor_load_dialog.raise_(); self._investor_load_dialog.activateWindow()
    def _on_investor_loaded(self, count, db_path, seconds):
        self.loading.hide_loading(); self._investor_db_path = db_path; self.status_lbl.setText(f"Отчёт инвестора: {count:,} строк".replace(",", " "))
        QMessageBox.information(self, "Отчёт инвестора готов", f"Подготовлено {count:,} строк за {seconds:.1f} сек.".replace(",", " "))
        InvestorReconciliationDialog(self.repo, db_path, self).exec()
    def _on_investor_error(self, message):
        self.loading.hide_loading(); self.status_lbl.setText("Ошибка отчёта инвестора"); QMessageBox.critical(self, "Ошибка импорта", message)
    def _build_toolbar(self):
        tb = QToolBar(); tb.setObjectName("main_toolbar"); tb.setMovable(False); tb.setFloatable(False)
        tb.setIconSize(QSize(Theme.Toolbar.icon_size, Theme.Toolbar.icon_size)); tb.setToolButtonStyle(Qt.ToolButtonIconOnly)
        p = QApplication.instance().palette()
        self.open_act = QAction(make_icon(ICON_OPEN, p), "Открыть", self); self.open_act.triggered.connect(self.load_csv); tb.addAction(self.open_act)
        tb.addSeparator()
        self.sidebar_act = QAction(make_icon(ICON_SIDEBAR, p), "Sidebar", self); self.sidebar_act.triggered.connect(self.toggle_sidebar); tb.addAction(self.sidebar_act)
        self.addToolBar(tb)
    def export_summary(self):
        self.export_all_rows = False
        self.export_data()
    def export_raw_rows(self):
        self.export_all_rows = True
        self.export_data()
    def _build_statusbar(self):
        self.statusbar = QStatusBar(); self.setStatusBar(self.statusbar)
        self.status_lbl = QLabel("Готово"); self.statusbar.addWidget(self.status_lbl)
    def toggle_sidebar(self):
        sizes = self.splitter.sizes()
        if sizes[0] > 100: self.splitter.setSizes([0, self.width()])
        else: self.splitter.setSizes([Theme.Sidebar.normal_width, self.width() - Theme.Sidebar.normal_width])
    def load_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "Открыть CSV", "", "CSV files (*.csv)")
        if not path: return
        suggested = os.path.splitext(path)[0] + ".duckdb"
        db, _ = QFileDialog.getSaveFileName(self, "Где сохранить базу данных", suggested, "DuckDB files (*.duckdb)")
        if not db: return
        if not db.lower().endswith(".duckdb"): db += ".duckdb"
        if os.path.exists(db):
            if QMessageBox.question(self, "Подтверждение", "Перезаписать БД?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.No: return
        self._current_db_path = os.path.abspath(db)
        self.loading.show_loading("Анализ CSV..."); self.status_lbl.setText("Загрузка...")
        self._iw = ImportWorker(self.repo, path, self._current_db_path)
        self._iw.progress.connect(self.loading.update_progress); self._iw.finished.connect(self._on_import_done)
        self._iw.error.connect(self._on_import_err); self._iw.start()
    def load_db(self):
        # Нативное окно macOS показывает «Все файлы» даже при фильтре.
        # Используем контролируемый диалог: видны только папки и .duckdb.
        dialog = QFileDialog(self, "Открыть БД")
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setNameFilter("Базы DuckDB (*.duckdb)")
        dialog.setFilter(QDir.AllDirs | QDir.Files | QDir.NoDotAndDotDot)
        dialog.setOption(QFileDialog.DontUseNativeDialog, True)
        if dialog.exec() != QDialog.Accepted: return
        selected = dialog.selectedFiles()
        path = selected[0] if selected else ""
        if not path: return
        try: self.repo.connect(path); self._current_db_path = path; self._on_loaded(); self.status_lbl.setText(f"БД: {os.path.basename(path)}"); QMessageBox.information(self, "Успех", f"Открыта:\n{path}")
        except Exception as e: QMessageBox.critical(self, "Ошибка", str(e))
    def _on_import_done(self, n, t):
        self.loading.hide_loading(); self._on_loaded(); self.status_lbl.setText(f"Загружено {n:,} строк за {t:.1f} сек")
        QMessageBox.information(self, "Готово", f"Загружено {n:,} строк за {t:.1f} сек.\nБД: {self._current_db_path}")
    def _on_import_err(self, msg):
        self.loading.hide_loading(); self.status_lbl.setText("Ошибка"); QMessageBox.critical(self, "Ошибка", msg)
    def run_filter(self):
        if not self._db_loaded: QMessageBox.warning(self, "Нет данных", "Сначала загрузите CSV или БД."); return
        f, date_f = self.sidebar.get_filters()
        w = self.repo.build_where_clause(f, date_f)
        self.table.model.set_where_clause(w); self.table.update_info_label(); self._upd_stats(); self.status_lbl.setText("Фильтры применены")
    def _upd_stats(self):
        f, date_f = self.sidebar.get_filters()
        w = self.repo.build_where_clause(f, date_f)
        if self._stats_worker and self._stats_worker.isRunning(): self._stats_worker.requestInterruption(); self._stats_worker.wait(1000)
        QTimer.singleShot(300, lambda: self._run_stats_query(w))
    def _run_stats_query(self, where_clause):
        self._stats_worker = StatsWorker(self.repo, where_clause); self._stats_worker.finished.connect(self._on_stats); self._stats_worker.start()
    def _on_stats(self, s): self.summary.update_stats(s)
    def export_data(self):
        if not self._db_loaded: QMessageBox.warning(self, "Нет данных", "Сначала загрузите данные."); return
        path, _ = QFileDialog.getSaveFileName(self, "Экспорт", "", "Excel (*.xlsx)")
        if not path: return
        f, date_f = self.sidebar.get_filters()
        w = self.repo.build_where_clause(f, date_f)
        self.status_lbl.setText("Экспорт...")
        if self.export_all_rows:
            total = self.repo.get_row_count(w)
            if total > 1048576:
                QMessageBox.warning(self, "Лимит Excel", f"В выборке {total:,} строк.\nExcel поддерживает максимум 1 048 576 строк.")
                self.status_lbl.setText("Экспорт отменен"); return
            if self.repo.export_to_excel(w, path): 
                self.status_lbl.setText(f"Экспорт: {os.path.basename(path)}"); QMessageBox.information(self, "Экспорт", f"Сохранено:\n{path}")
            else: 
                self.status_lbl.setText("Ошибка экспорта"); QMessageBox.critical(self, "Ошибка", "Не удалось экспортировать")
        else:
            if self.repo.export_aggregated_to_excel(w, path): 
                self.status_lbl.setText(f"Экспорт: {os.path.basename(path)}"); QMessageBox.information(self, "Экспорт", f"Сводка сохранена:\n{path}")
            else: 
                self.status_lbl.setText("Ошибка экспорта"); QMessageBox.critical(self, "Ошибка", "Не удалось экспортировать сводку")
    def _on_table_double_click(self, idx):
        if not idx.isValid(): return
        carrier = self.table.model.data(self.table.model.index(idx.row(), 0), Qt.DisplayRole)
        if not carrier or carrier == "(Пусто)": return
        f, date_f = self.sidebar.get_filters()
        where = self.repo.build_where_clause(f, date_f)
        dlg = DetailsViewer(self.repo, carrier, where, self); dlg.exec()
    def check_updates(self):
        self.status_lbl.setText("Проверка обновлений...")
        self._update_worker = UpdateWorker(self)
        self._update_worker.completed.connect(self._on_update_checked)
        self._update_worker.error.connect(self._on_update_error)
        self._update_worker.start()
    def _on_update_checked(self, manifest):
        self.status_lbl.setText("Готово")
        version = str(manifest.get("version", ""))
        if not version or version <= APP_VERSION:
            QMessageBox.information(self, "Обновления", "Установлена актуальная версия.")
            return
        notes = str(manifest.get("notes", "Нет описания изменений."))
        script_url = str(manifest.get("script_url", "")).strip()
        answer = QMessageBox.question(
            self, "Доступно обновление",
            f"Доступна версия {version}.\n\n{notes}\n\nСкачать обновление и заменить текущий файл программы?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if answer != QMessageBox.Yes: return
        if not script_url:
            QMessageBox.warning(self, "Обновление", "В манифесте нет ссылки на файл программы.")
            return
        self.status_lbl.setText("Скачивание обновления…")
        self._script_update_worker = ScriptUpdateWorker(script_url, str(manifest.get("script_sha256", "")), self)
        self._script_update_worker.completed.connect(lambda backup: self._on_script_updated(version, backup))
        self._script_update_worker.error.connect(self._on_update_error)
        self._script_update_worker.start()

    def _on_script_updated(self, version, backup_path):
        self.status_lbl.setText("Обновление загружено")
        answer = QMessageBox.question(
            self, "Обновление готово",
            f"Версия {version} установлена. Предыдущая версия сохранена:\n{backup_path}\n\nПерезапустить приложение сейчас?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if answer == QMessageBox.Yes:
            subprocess.Popen([sys.executable, os.path.abspath(sys.argv[0])], start_new_session=True)
            QApplication.quit()
    def _on_update_error(self, message):
        self.status_lbl.setText("Не удалось проверить обновления")
        QMessageBox.information(self, "Обновления", message)
    def _save_settings(self):
        s = QSettings("TransactionAnalyzer", "Analyzer"); s.setValue("window/geometry", self.saveGeometry()); s.setValue("window/state", self.saveState()); s.setValue("splitter", self.splitter.saveState())
    def _load_settings(self):
        s = QSettings("TransactionAnalyzer", "Analyzer"); g = s.value("window/geometry")
        if g: self.restoreGeometry(g)
        st = s.value("window/state")
        if st: self.restoreState(st)
        sp = s.value("splitter")
        if sp: self.splitter.restoreState(sp)
    def closeEvent(self, e): self._save_settings(); self.repo.close(); super().closeEvent(e)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Один контроллер обслуживает все таблицы — включая те, что открываются
    # позднее в отдельных окнах АСОП, ВТБ и сверки.
    copy_controller = TableCopyController(app)
    app.installEventFilter(copy_controller)
    f = QFontDatabase.systemFont(QFontDatabase.GeneralFont)
    f.setHintingPreference(QFont.PreferNoHinting); f.setStyleStrategy(QFont.PreferAntialias)
    app.setFont(f); w = AnalyzerApp(); w.show(); sys.exit(app.exec())

import sqlite3
import hashlib
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import calendar

# Попытка импорта openpyxl
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, GradientFill
    from openpyxl.utils import get_column_letter
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False

# ══════════════════════════════════════════════════════════════
#  ПАЛИТРА
# ══════════════════════════════════════════════════════════════
BG_COLOR      = "#1c2033"   # основной фон — насыщенный тёмно-синий (светлее)
PANEL_COLOR   = "#252a3d"   # панели — заметно светлее фона
CARD_COLOR    = "#2e3450"   # карточки/поля — синевато-серый
PRIMARY_COLOR = "#2a3058"   # заголовки-шапки
COMBO_COLOR   = "#384068"   # цвет комбобокса АЗС — отличается от фона
ACCENT_COLOR  = "#5b9bff"   # синий акцент — ярче
ACCENT2_COLOR = "#9b72ff"   # фиолетовый акцент — ярче
SUCCESS_COLOR = "#30d974"   # зелёный — ярче
WARNING_COLOR = "#ffab2e"   # оранжевый — ярче
DANGER_COLOR  = "#ff4f4f"   # красный — ярче
TEXT_COLOR    = "#f0f4ff"   # основной текст — ярче, чище
TEXT_DIM      = "#8d9bbf"   # приглушённый — виднее
BORDER_COLOR  = "#3d4878"   # границы — заметнее

# ══════════════════════════════════════════════════════════════
#  УТИЛИТЫ
# ══════════════════════════════════════════════════════════════
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


def apply_global_style():
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('.',            background=BG_COLOR,    foreground=TEXT_COLOR, font=('Segoe UI', 10))
    style.configure('TFrame',       background=BG_COLOR)
    style.configure('TLabel',       background=BG_COLOR,    foreground=TEXT_COLOR, font=('Segoe UI', 10))
    style.configure('TLabelframe',  background=PANEL_COLOR, foreground=TEXT_COLOR, bordercolor=BORDER_COLOR, relief='flat')
    style.configure('TLabelframe.Label', background=PANEL_COLOR, foreground=TEXT_COLOR, font=('Segoe UI', 10, 'bold'))
    style.configure('TButton',      background=CARD_COLOR,  foreground=TEXT_COLOR, font=('Segoe UI', 10, 'bold'), padding=(12, 7), relief='flat', borderwidth=0)
    style.map('TButton', background=[('active', ACCENT_COLOR), ('pressed', ACCENT2_COLOR), ('disabled', '#2a2d3e')], foreground=[('disabled', TEXT_DIM)])
    style.configure('TEntry',       fieldbackground=CARD_COLOR, foreground=TEXT_COLOR, insertcolor=TEXT_COLOR, bordercolor=BORDER_COLOR, lightcolor=BORDER_COLOR, darkcolor=BORDER_COLOR, font=('Segoe UI', 11), padding=(8, 6))
    style.map('TEntry', bordercolor=[('focus', ACCENT_COLOR)])
    style.configure('TCombobox',    fieldbackground=CARD_COLOR, background=CARD_COLOR, foreground=TEXT_COLOR, selectbackground=ACCENT_COLOR, selectforeground='white', font=('Segoe UI', 10), padding=(6, 5))
    style.map('TCombobox', fieldbackground=[('readonly', CARD_COLOR)], selectbackground=[('readonly', CARD_COLOR)], selectforeground=[('readonly', TEXT_COLOR)])
    style.configure('TSpinbox',     fieldbackground=CARD_COLOR, foreground=TEXT_COLOR, background=CARD_COLOR, font=('Segoe UI', 10), padding=(6, 4))
    style.configure('Treeview',     background=PANEL_COLOR, foreground=TEXT_COLOR, rowheight=32, fieldbackground=PANEL_COLOR, borderwidth=0, font=('Segoe UI', 10))
    style.configure('Treeview.Heading', background=PRIMARY_COLOR, foreground=ACCENT_COLOR, font=('Segoe UI', 10, 'bold'), padding=(8, 8), relief='flat')
    style.map('Treeview.Heading',   background=[('active', CARD_COLOR)])
    style.map('Treeview',           background=[('selected', ACCENT_COLOR)], foreground=[('selected', 'white')])
    style.configure('TScrollbar',   background=CARD_COLOR, troughcolor=PANEL_COLOR, bordercolor=PANEL_COLOR, arrowcolor=TEXT_DIM, relief='flat')
    style.configure('TNotebook',    background=BG_COLOR, borderwidth=0)
    style.configure('TNotebook.Tab', background=CARD_COLOR, foreground=TEXT_DIM, padding=(14, 7), font=('Segoe UI', 10, 'bold'))
    style.map('TNotebook.Tab', background=[('selected', ACCENT_COLOR), ('active', BORDER_COLOR)], foreground=[('selected', 'white'), ('active', TEXT_COLOR)])
    return style


def styled_entry(parent, textvariable=None, show=None, width=22, font_size=12):
    kw = dict(textvariable=textvariable, font=('Segoe UI', font_size),
              bg=CARD_COLOR, fg=TEXT_COLOR, insertbackground=TEXT_COLOR,
              relief='flat', bd=0, highlightthickness=2,
              highlightcolor=ACCENT_COLOR, highlightbackground=BORDER_COLOR,
              width=width)
    if show:
        kw['show'] = show
    return Entry(parent, **kw)


def mk_btn(parent, text, cmd, style='normal', state=NORMAL, padx=12, pady=6):
    colors = {
        'normal':  (CARD_COLOR,    TEXT_COLOR,  '#2d3250'),
        'success': (SUCCESS_COLOR, 'white',     '#27ae60'),
        'danger':  (DANGER_COLOR,  'white',     '#c0392b'),
        'accent':  (ACCENT_COLOR,  'white',     '#3a7aed'),
        'warn':    (WARNING_COLOR, 'white',     '#d68910'),
        'purple':  (ACCENT2_COLOR, 'white',     '#6a4ce0'),
    }
    bg, fg, abg = colors.get(style, colors['normal'])
    return Button(parent, text=text, font=('Segoe UI', 10, 'bold'),
                  bg=bg, fg=fg, activebackground=abg, activeforeground=fg,
                  disabledforeground=TEXT_DIM,
                  relief='flat', bd=0, padx=padx, pady=pady,
                  cursor='hand2', command=cmd, state=state)


def dialog_header(win, icon_text, title_text, color=None):
    """Стандартный заголовок диалогового окна"""
    c = color or PRIMARY_COLOR
    hdr = Frame(win, bg=c, height=54)
    hdr.pack(fill=X)
    hdr.pack_propagate(False)
    Label(hdr, text=f"{icon_text}  {title_text}", font=('Segoe UI', 13, 'bold'),
          bg=c, fg=TEXT_COLOR).pack(side=LEFT, padx=20, pady=14)
    Frame(win, bg=ACCENT_COLOR, height=2).pack(fill=X)


def dialog_footer(win, ok_text, ok_cmd, ok_style='success', cancel_cmd=None):
    """Стандартный подвал диалогового окна"""
    Frame(win, bg=BORDER_COLOR, height=1).pack(fill=X)
    bar = Frame(win, bg=PANEL_COLOR, height=58)
    bar.pack(fill=X)
    bar.pack_propagate(False)
    inner = Frame(bar, bg=PANEL_COLOR)
    inner.pack(side=RIGHT, padx=18, pady=10)
    mk_btn(inner, ok_text, ok_cmd, ok_style, padx=18, pady=8).pack(side=LEFT, padx=(0, 8))
    if cancel_cmd:
        mk_btn(inner, "✕  Отмена", cancel_cmd, 'normal', padx=18, pady=8).pack(side=LEFT)


# ══════════════════════════════════════════════════════════════
#  БАЗА ДАННЫХ
# ══════════════════════════════════════════════════════════════
def init_db():
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS azs_stations (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, code TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, color TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, category_id INTEGER, price REAL,
        FOREIGN KEY (category_id) REFERENCES categories(id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT,
        full_name TEXT, role TEXT, station_id INTEGER,
        FOREIGN KEY (station_id) REFERENCES azs_stations(id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT, station_id INTEGER, product_id INTEGER,
        quantity REAL, total REAL, sale_date TEXT, shift INTEGER, user_id INTEGER,
        FOREIGN KEY (station_id) REFERENCES azs_stations(id),
        FOREIGN KEY (product_id) REFERENCES products(id),
        FOREIGN KEY (user_id) REFERENCES users(id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS action_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, action TEXT, timestamp TEXT,
        FOREIGN KEY (user_id) REFERENCES users(id))''')
    # Таблица запросов на сброс пароля
    c.execute('''CREATE TABLE IF NOT EXISTS password_reset_requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER, username TEXT, full_name TEXT,
        requested_at TEXT, status TEXT DEFAULT 'pending',
        FOREIGN KEY (user_id) REFERENCES users(id))''')

    for name, code in [('АЗС №1 (Центральная)', 'AZS001'), ('АЗС №2 (Северная)', 'AZS002'), ('АЗС №3 (Южная)', 'AZS003')]:
        c.execute("INSERT OR IGNORE INTO azs_stations (name, code) VALUES (?, ?)", (name, code))
    for name, color in [('Топливо', '#e74c3c'), ('Снеки', '#f1c40f'), ('Напитки', '#3498db'), ('Разное', '#95a5a6')]:
        c.execute("INSERT OR IGNORE INTO categories (name, color) VALUES (?, ?)", (name, color))

    c.execute("SELECT id, name FROM categories")
    cat_ids = {n: i for i, n in c.fetchall()}
    for name, cat, price in [
        ('АИ-92', 'Топливо', 51.20), ('АИ-95', 'Топливо', 55.30), ('ДТ', 'Топливо', 60.10),
        ('Чипсы Lays', 'Снеки', 120.0), ('Шоколад Snickers', 'Снеки', 80.0),
        ('Вода 0.5л', 'Напитки', 40.0), ('Кола 0.5л', 'Напитки', 70.0),
        ('Кофе американо', 'Напитки', 120.0), ('Жвачка', 'Разное', 30.0)]:
        c.execute("INSERT OR IGNORE INTO products (name, category_id, price) VALUES (?, ?, ?)", (name, cat_ids[cat], price))

    for username, pwd, full, role, st_id in [
        ('admin',   hash_password('admin'),   'Администратор', 'admin', None),
        ('user1',   hash_password('111'),     'Иванов Иван',   'user',  1),
        ('user2',   hash_password('222'),     'Петров Пётр',   'user',  2),
        ('user3',   hash_password('333'),     'Сидоров Сидор', 'user',  3),
        ('manager', hash_password('manager'), 'Управляющий',   'admin', None)]:
        c.execute("INSERT OR IGNORE INTO users (username, password, full_name, role, station_id) VALUES (?, ?, ?, ?, ?)",
                  (username, pwd, full, role, st_id))

    c.execute("SELECT COUNT(*) FROM sales")
    if c.fetchone()[0] == 0:
        c.execute("SELECT id, name FROM products")
        prod_ids = {n: i for i, n in c.fetchall()}
        c.execute("SELECT id, username FROM users")
        user_ids = {u: i for i, u in c.fetchall()}
        for st_id, prod, qty, date, shift, uid in [
            (1,'АИ-95',150.5,'2025-02-20',1,'user1'), (1,'ДТ',200.0,'2025-02-20',1,'user1'),
            (1,'Чипсы Lays',3,'2025-02-20',1,'user1'), (1,'Кола 0.5л',5,'2025-02-20',1,'user1'),
            (2,'АИ-92',410.5,'2025-02-20',1,'user2'), (2,'Вода 0.5л',10,'2025-02-20',1,'user2'),
            (3,'ДТ',175.0,'2025-02-20',2,'user3'), (3,'Кофе американо',8,'2025-02-20',2,'user3'),
            (3,'Жвачка',15,'2025-02-20',2,'user3')]:
            c.execute("SELECT price FROM products WHERE id=?", (prod_ids[prod],))
            price = c.fetchone()[0]
            c.execute("INSERT INTO sales (station_id,product_id,quantity,total,sale_date,shift,user_id) VALUES (?,?,?,?,?,?,?)",
                      (st_id, prod_ids[prod], qty, qty*price, date, shift, user_ids[uid]))

    conn.commit()
    conn.close()


# ── DB helpers ──────────────────────────────────────────────
def check_user(username, password):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("SELECT id,username,full_name,role,station_id FROM users WHERE username=? AND password=?",
              (username, hash_password(password)))
    r = c.fetchone(); conn.close(); return r


def get_stations(role, station_id):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    if role == 'admin':
        c.execute("SELECT id,name FROM azs_stations ORDER BY name")
    else:
        c.execute("SELECT id,name FROM azs_stations WHERE id=?", (station_id,))
    r = c.fetchall(); conn.close(); return r


def get_station_name(station_id):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("SELECT name FROM azs_stations WHERE id=?", (station_id,))
    r = c.fetchone(); conn.close()
    return r[0] if r else ""


def get_products():
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("SELECT p.id,p.name,cat.name,p.price FROM products p JOIN categories cat ON p.category_id=cat.id ORDER BY cat.name,p.name")
    r = c.fetchall(); conn.close(); return r


def get_users(role='all'):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    if role == 'all':
        c.execute("SELECT id,full_name,username,role,station_id FROM users ORDER BY full_name")
    else:
        c.execute("SELECT id,full_name,username,role,station_id FROM users WHERE role=? ORDER BY full_name", (role,))
    r = c.fetchall(); conn.close(); return r


def get_sales(station_ids, month=None, user_id_filter=None):
    if not station_ids: return []
    ph = ','.join('?'*len(station_ids))
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    query = f'''SELECT s.id,s.sale_date,p.name,cat.name,s.quantity,p.price,s.total,s.shift,st.name,u.full_name
        FROM sales s
        JOIN products p ON s.product_id=p.id
        JOIN categories cat ON p.category_id=cat.id
        JOIN azs_stations st ON s.station_id=st.id
        LEFT JOIN users u ON s.user_id=u.id
        WHERE s.station_id IN ({ph})'''
    params = list(station_ids)
    if month:
        y, m = month
        ld = calendar.monthrange(y, m)[1]
        query += " AND s.sale_date BETWEEN ? AND ?"
        params += [f"{y}-{m:02d}-01", f"{y}-{m:02d}-{ld}"]
    if user_id_filter:
        query += " AND s.user_id=?"
        params.append(user_id_filter)
    query += " ORDER BY s.sale_date DESC,s.id DESC"
    c.execute(query, params); r = c.fetchall(); conn.close(); return r


def add_sale(station_id, product_id, quantity, sale_date, shift, user_id):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("SELECT price FROM products WHERE id=?", (product_id,))
    price = c.fetchone()[0]
    c.execute("INSERT INTO sales (station_id,product_id,quantity,total,sale_date,shift,user_id) VALUES (?,?,?,?,?,?,?)",
              (station_id, product_id, quantity, quantity*price, sale_date, shift, user_id))
    conn.commit(); conn.close()


def update_sale(sale_id, product_id, quantity, sale_date, shift, user_id):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("SELECT price FROM products WHERE id=?", (product_id,))
    price = c.fetchone()[0]
    c.execute("UPDATE sales SET product_id=?,quantity=?,total=?,sale_date=?,shift=?,user_id=? WHERE id=?",
              (product_id, quantity, quantity*price, sale_date, shift, user_id, sale_id))
    conn.commit(); conn.close()


def delete_sale(sale_id):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("DELETE FROM sales WHERE id=?", (sale_id,))
    conn.commit(); conn.close()


def update_password(user_id, old_password, new_password):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("SELECT password FROM users WHERE id=?", (user_id,))
    stored = c.fetchone()[0]
    if stored != hash_password(old_password):
        conn.close(); return False
    c.execute("UPDATE users SET password=? WHERE id=?", (hash_password(new_password), user_id))
    conn.commit(); conn.close(); return True


def admin_set_password(user_id, new_password):
    """Администратор устанавливает пароль без проверки старого"""
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("UPDATE users SET password=? WHERE id=?", (hash_password(new_password), user_id))
    conn.commit(); conn.close()


def log_action(user_id, action):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("INSERT INTO action_log (user_id,action,timestamp) VALUES (?,?,?)",
              (user_id, action, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit(); conn.close()


def get_last_actions(limit=50):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute('''SELECT al.timestamp,u.full_name,al.action FROM action_log al
        JOIN users u ON al.user_id=u.id ORDER BY al.id DESC LIMIT ?''', (limit,))
    r = c.fetchall(); conn.close(); return r


# ── Password reset requests ──────────────────────────────────
def create_reset_request(user_id, username, full_name):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    # Удалим старые pending-запросы этого пользователя
    c.execute("DELETE FROM password_reset_requests WHERE user_id=? AND status='pending'", (user_id,))
    c.execute("INSERT INTO password_reset_requests (user_id,username,full_name,requested_at,status) VALUES (?,?,?,?,'pending')",
              (user_id, username, full_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit(); conn.close()


def get_reset_requests():
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("SELECT id,user_id,username,full_name,requested_at FROM password_reset_requests WHERE status='pending' ORDER BY requested_at DESC")
    r = c.fetchall(); conn.close(); return r


def resolve_reset_request(req_id):
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("UPDATE password_reset_requests SET status='resolved' WHERE id=?", (req_id,))
    conn.commit(); conn.close()


def get_pending_reset_count():
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM password_reset_requests WHERE status='pending'")
    r = c.fetchone()[0]; conn.close(); return r


# ══════════════════════════════════════════════════════════════
#  EXCEL EXPORT (улучшенный)
# ══════════════════════════════════════════════════════════════
def _xl_header_fill():
    return PatternFill("solid", fgColor="1E2235")

def _xl_accent_fill():
    return PatternFill("solid", fgColor="4F8EF7")

def _xl_alt_fill():
    return PatternFill("solid", fgColor="22263A")

def _xl_total_fill():
    return PatternFill("solid", fgColor="2D3250")

def _xl_thin_border():
    s = Side(style='thin', color='2D3250')
    return Border(left=s, right=s, top=s, bottom=s)

def _xl_header_font(size=11):
    return Font(name='Calibri', bold=True, color='E8ECF4', size=size)

def _xl_title_font():
    return Font(name='Calibri', bold=True, color='FFFFFF', size=14)

def _xl_data_font():
    return Font(name='Calibri', color='E8ECF4', size=10)

def _xl_total_font():
    return Font(name='Calibri', bold=True, color='4F8EF7', size=11)

def _xl_center():
    return Alignment(horizontal='center', vertical='center', wrap_text=False)

def _xl_left():
    return Alignment(horizontal='left', vertical='center')

def write_report_sheet(sheet, rows, year, month, employee_name=None):
    # Фон листа (всех ячеек) — через tab color
    sheet.sheet_properties.tabColor = "4F8EF7"

    # ── Строка 1: Заголовок ─────────────────────────────────
    emp_label = f" | Сотрудник: {employee_name}" if employee_name else ""
    sheet.merge_cells('A1:I1')
    c = sheet['A1']
    c.value = f"  ОТЧЁТ ЗА {month:02d}.{year}{emp_label}"
    c.font = _xl_title_font()
    c.fill = _xl_accent_fill()
    c.alignment = _xl_center()
    sheet.row_dimensions[1].height = 32

    # ── Строка 2: подпись ───────────────────────────────────
    sheet.merge_cells('A2:I2')
    c2 = sheet['A2']
    c2.value = f"  Сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    c2.font = Font(name='Calibri', color='7A84A0', size=9, italic=True)
    c2.fill = _xl_header_fill()
    c2.alignment = _xl_left()
    sheet.row_dimensions[2].height = 18

    # ── Строка 3: заголовки столбцов ────────────────────────
    headers = ['Дата', 'Товар', 'Категория', 'Кол-во', 'Цена, руб.', 'Сумма, руб.', 'Смена', 'АЗС', 'Оператор']
    widths   = [14,     26,      14,           10,        13,           14,             8,       22,     20]
    for col, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = sheet.cell(row=3, column=col)
        cell.value = h
        cell.font = _xl_header_font()
        cell.fill = _xl_header_fill()
        cell.alignment = _xl_center()
        cell.border = _xl_thin_border()
        sheet.column_dimensions[get_column_letter(col)].width = w
    sheet.row_dimensions[3].height = 24

    # ── Данные ───────────────────────────────────────────────
    CATEGORY_COLORS = {
        'Топливо': 'FF4040',
        'Снеки':   'F5C842',
        'Напитки': '7EC8E3',
        'Разное':  '95A5A6',
    }
    total_sum = 0
    for r_idx, row in enumerate(rows):
        excel_row = r_idx + 4
        is_alt = (r_idx % 2 == 1)
        row_fill = _xl_alt_fill() if is_alt else PatternFill("solid", fgColor="1A1D27")
        values = [row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9] or '—']
        for col, val in enumerate(values, start=1):
            cell = sheet.cell(row=excel_row, column=col)
            cell.value = val
            cell.border = _xl_thin_border()
            cell.alignment = _xl_center() if col not in (2, 8, 9) else _xl_left()
            # Категория — цветной текст
            cat = row[3]
            cat_color = CATEGORY_COLORS.get(cat, 'E8ECF4')
            if col == 3:
                cell.font = Font(name='Calibri', color=cat_color, size=10, bold=True)
            else:
                cell.font = _xl_data_font()
            cell.fill = row_fill
        sheet.row_dimensions[excel_row].height = 20
        total_sum += row[6]

    # ── Итого ────────────────────────────────────────────────
    total_row = len(rows) + 4
    sheet.merge_cells(f'A{total_row}:E{total_row}')
    c = sheet.cell(row=total_row, column=1)
    c.value = "ИТОГО:"
    c.font = _xl_total_font()
    c.fill = _xl_total_fill()
    c.alignment = Alignment(horizontal='right', vertical='center')
    c.border = _xl_thin_border()
    for col in range(2, 6):
        sheet.cell(row=total_row, column=col).fill = _xl_total_fill()
        sheet.cell(row=total_row, column=col).border = _xl_thin_border()
    tc = sheet.cell(row=total_row, column=6)
    tc.value = total_sum
    tc.number_format = '#,##0.00 "руб."'
    tc.font = _xl_total_font()
    tc.fill = _xl_total_fill()
    tc.alignment = _xl_center()
    tc.border = _xl_thin_border()
    for col in range(7, 10):
        sheet.cell(row=total_row, column=col).fill = _xl_total_fill()
        sheet.cell(row=total_row, column=col).border = _xl_thin_border()
    sheet.row_dimensions[total_row].height = 24

    # Freeze top rows
    sheet.freeze_panes = 'A4'

    return total_sum


def write_summary_sheet(sheet, stations_dict, year, month, employee_name=None):
    sheet.sheet_properties.tabColor = "2ECC71"
    emp_label = f" | {employee_name}" if employee_name else ""
    sheet.merge_cells('A1:D1')
    c = sheet['A1']
    c.value = f"  СВОДКА ПО ВСЕМ АЗС — {month:02d}.{year}{emp_label}"
    c.font = _xl_title_font()
    c.fill = PatternFill("solid", fgColor="2ECC71")
    c.alignment = _xl_center()
    sheet.row_dimensions[1].height = 32

    sheet.merge_cells('A2:D2')
    c2 = sheet['A2']
    c2.value = f"  Сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    c2.font = Font(name='Calibri', color='7A84A0', size=9, italic=True)
    c2.fill = _xl_header_fill()
    c2.alignment = _xl_left()

    for col, (h, w) in enumerate(zip(['АЗС', 'Продаж', 'Сумма, руб.', 'Доля, %'], [28, 12, 18, 12]), start=1):
        cell = sheet.cell(row=3, column=col)
        cell.value = h
        cell.font = _xl_header_font()
        cell.fill = _xl_header_fill()
        cell.alignment = _xl_center()
        cell.border = _xl_thin_border()
        sheet.column_dimensions[get_column_letter(col)].width = w
    sheet.row_dimensions[3].height = 24

    grand_total = sum(sum(r[6] for r in rows) for rows in stations_dict.values())
    for i, (station, rows) in enumerate(stations_dict.items()):
        cnt = len(rows); summ = sum(r[6] for r in rows)
        pct = round(summ/grand_total*100, 1) if grand_total else 0
        fill = _xl_alt_fill() if i % 2 else PatternFill("solid", fgColor="1A1D27")
        for col, val in enumerate([station, cnt, summ, f"{pct}%"], start=1):
            cell = sheet.cell(row=4+i, column=col)
            cell.value = val; cell.font = _xl_data_font()
            cell.fill = fill; cell.border = _xl_thin_border()
            cell.alignment = _xl_left() if col == 1 else _xl_center()
        sheet.row_dimensions[4+i].height = 20

    tr = 4 + len(stations_dict)
    for col, val in enumerate(['ВСЕГО:', len([r for rows in stations_dict.values() for r in rows]), grand_total, '100%'], start=1):
        cell = sheet.cell(row=tr, column=col)
        cell.value = val; cell.font = _xl_total_font()
        cell.fill = _xl_total_fill(); cell.border = _xl_thin_border()
        cell.alignment = _xl_left() if col == 1 else _xl_center()
    sheet.row_dimensions[tr].height = 24


def generate_excel_report(station_ids, year, month, user_role, user_name="", user_id_filter=None, employee_name=None):
    if not EXCEL_SUPPORT:
        messagebox.showerror("Ошибка", "Установите openpyxl:\n  pip install openpyxl")
        return

    rows = get_sales(station_ids, (year, month), user_id_filter)
    if not rows:
        messagebox.showinfo("Отчёт", "Нет данных за выбранный период")
        return

    stations_dict = {}
    for row in rows:
        st = row[8]
        stations_dict.setdefault(st, []).append(row)

    wb = Workbook()
    wb.remove(wb.active)

    if user_role == 'admin' and len(stations_dict) > 1 and not user_id_filter:
        ws = wb.create_sheet("Сводка", 0)
        write_summary_sheet(ws, stations_dict, year, month, employee_name)
        for st_name, st_rows in stations_dict.items():
            ws2 = wb.create_sheet(st_name[:31])
            write_report_sheet(ws2, st_rows, year, month, employee_name)
    else:
        ws = wb.create_sheet("Отчёт")
        write_report_sheet(ws, rows, year, month, employee_name)

    init_name = f"report_{year}_{month:02d}"
    if employee_name:
        init_name += f"_{employee_name.replace(' ', '_')}"
    elif user_name:
        init_name += f"_{user_name.replace(' ', '_')}"

    filename = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Сохранить отчёт Excel",
        initialfile=init_name + ".xlsx"
    )
    if filename:
        wb.save(filename)
        messagebox.showinfo("✔  Готово", f"Отчёт сохранён:\n{filename}")
        log_action_ext(user_name, f"Экспорт Excel: {filename}")


def log_action_ext(user_name, action):
    """Логирование без user_id (для вспомогательных функций)"""
    conn = sqlite3.connect('azs_shop.db')
    c = conn.cursor()
    c.execute("SELECT id FROM users WHERE full_name=? LIMIT 1", (user_name,))
    r = c.fetchone()
    if r:
        c.execute("INSERT INTO action_log (user_id,action,timestamp) VALUES (?,?,?)",
                  (r[0], action, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
    conn.close()


# ══════════════════════════════════════════════════════════════
#  ОКНО ВХОДА
# ══════════════════════════════════════════════════════════════
class LoginWindow:
    def __init__(self):
        self.win = Tk()
        self.win.title("АЗС-Магазин — Вход")
        self.win.geometry("480x620")
        self.win.configure(bg=BG_COLOR)
        self.win.resizable(False, False)
        apply_global_style()
        self.win.update_idletasks()
        x = (self.win.winfo_screenwidth() - 480) // 2
        y = (self.win.winfo_screenheight() - 620) // 2
        self.win.geometry(f"480x620+{x}+{y}")

        # Шапка
        hdr = Frame(self.win, bg=PRIMARY_COLOR, height=170)
        hdr.pack(fill=X); hdr.pack_propagate(False)
        Label(hdr, text="⛽", font=('Segoe UI Emoji', 44), bg=PRIMARY_COLOR, fg=ACCENT_COLOR).pack(pady=(22, 0))
        Label(hdr, text="АЗС-МАГАЗИН", font=('Segoe UI', 17, 'bold'), bg=PRIMARY_COLOR, fg=TEXT_COLOR).pack()
        Label(hdr, text="Система учёта продаж", font=('Segoe UI', 10), bg=PRIMARY_COLOR, fg=TEXT_DIM).pack()
        Frame(self.win, bg=ACCENT_COLOR, height=3).pack(fill = X)

        # Форма
        form = Frame(self.win, bg=BG_COLOR)
        form.pack(expand=True, fill=BOTH, padx=52, pady=28)

        Label(form, text="Логин", font=('Segoe UI', 10, 'bold'), bg=BG_COLOR, fg=TEXT_DIM).pack(anchor=W)
        self.entry_user = styled_entry(form, width=30, font_size=12)
        self.entry_user.pack(fill=X, pady=(4, 18), ipady=8)

        Label(form, text="Пароль", font=('Segoe UI', 10, 'bold'), bg=BG_COLOR, fg=TEXT_DIM).pack(anchor=W)
        self.entry_pass = styled_entry(form, show="•", width=30, font_size=12)
        self.entry_pass.pack(fill=X, pady=(4, 24), ipady=8)

        Button(form, text="  ВОЙТИ  →", font=('Segoe UI', 13, 'bold'),
               bg=ACCENT_COLOR, fg='white', activebackground=ACCENT2_COLOR,
               activeforeground='white', relief='flat', bd=0, cursor='hand2',
               command=self.login).pack(fill=X, ipady=10)

        # Кнопка «Забыл пароль»
        Frame(form, bg=BG_COLOR, height=12).pack()
        Button(form, text="Забыл пароль?", font=('Segoe UI', 10),
               bg=BG_COLOR, fg=TEXT_DIM, activebackground=BG_COLOR,
               activeforeground=ACCENT_COLOR, relief='flat', bd=0, cursor='hand2',
               command=self.forgot_password).pack()

        hint = Frame(self.win, bg=CARD_COLOR)
        hint.pack(fill=X, padx=52, pady=(4, 24))
        Label(hint, text="  admin / admin   •   user1 / 111",
              font=('Segoe UI', 9), bg=CARD_COLOR, fg=TEXT_DIM).pack(pady=8)

        self.win.bind('<Return>', lambda e: self.login())
        self.entry_user.focus_set()
        self.win.mainloop()

    def login(self):
        username = self.entry_user.get().strip()
        password = self.entry_pass.get()
        user = check_user(username, password)
        if user:
            self.win.destroy()
            MainApp(user)
        else:
            self.entry_pass.config(highlightbackground=DANGER_COLOR, highlightcolor=DANGER_COLOR)
            self.win.after(700, lambda: self.entry_pass.config(
                highlightbackground=BORDER_COLOR, highlightcolor=ACCENT_COLOR))
            messagebox.showerror("Ошибка входа", "Неверный логин или пароль")

    def forgot_password(self):
        username = self.entry_user.get().strip()
        if not username:
            messagebox.showinfo("Забыл пароль",
                                "Введите ваш логин в поле выше,\nзатем нажмите «Забыл пароль».")
            return
        # Ищем пользователя
        conn = sqlite3.connect('azs_shop.db')
        c = conn.cursor()
        c.execute("SELECT id,full_name FROM users WHERE username=?", (username,))
        row = c.fetchone(); conn.close()
        if not row:
            messagebox.showerror("Ошибка", f"Пользователь «{username}» не найден.")
            return
        user_id, full_name = row
        create_reset_request(user_id, username, full_name)
        messagebox.showinfo("Запрос отправлен",
                            f"Запрос на сброс пароля для «{full_name}» отправлен администратору.\n\n"
                            "Обратитесь к администратору — он установит вам новый пароль.")


# ══════════════════════════════════════════════════════════════
#  ДИАЛОГ ДОБАВЛЕНИЯ / РЕДАКТИРОВАНИЯ ПРОДАЖИ
# ══════════════════════════════════════════════════════════════
class EditSaleWindow:
    def __init__(self, parent, sale_data=None, allowed_station_id=None, user_role='user', current_user_id=None):
        self.parent = parent
        self.sale_data = sale_data
        self.allowed_station_id = allowed_station_id
        self.user_role = user_role
        self.current_user_id = current_user_id
        self.result = None

        self.win = Toplevel(parent.root if hasattr(parent, 'root') else parent)
        title_text = "✏  Редактирование продажи" if sale_data else "➕  Добавление продажи"
        self.win.title(title_text)
        self.win.geometry("560x520")
        self.win.configure(bg=BG_COLOR)
        self.win.grab_set()
        self.win.resizable(False, False)
        apply_global_style()

        dialog_header(self.win, "✏" if sale_data else "➕", title_text.split("  ", 1)[1])

        form = Frame(self.win, bg=BG_COLOR, padx=28, pady=16)
        form.pack(fill=BOTH, expand=True)
        form.columnconfigure(1, weight=1)

        def lbl(row, text):
            Label(form, text=text, font=('Segoe UI', 10, 'bold'),
                  bg=BG_COLOR, fg=TEXT_DIM, anchor=W).grid(row=row, column=0, sticky=W, pady=(10, 2), padx=(0, 16))

        # АЗС
        lbl(0, "АЗС")
        self.station_var = StringVar()
        if user_role == 'admin' and not sale_data:
            stations = get_stations('admin', None)
            self.station_combo = ttk.Combobox(form, textvariable=self.station_var, state='readonly', width=38, font=('Segoe UI', 10))
            self.station_combo['values'] = [f"{sid} - {name}" for sid, name in stations]
            if stations: self.station_combo.current(0)
            self.station_combo.grid(row=0, column=1, sticky=EW, pady=(10, 2))
        else:
            if sale_data:
                st_name = sale_data[8]
                st_id = self._get_station_id_by_name(st_name)
            else:
                st_id = allowed_station_id
                st_name = get_station_name(st_id)
            self.station_var.set(f"{st_id} - {st_name}")
            Label(form, textvariable=self.station_var, font=('Segoe UI', 10),
                  bg=CARD_COLOR, fg=TEXT_COLOR, relief='flat', anchor=W, padx=8, pady=5
                  ).grid(row=0, column=1, sticky=EW, pady=(10, 2))

        # Товар
        lbl(1, "Товар")
        self.product_var = StringVar()
        self.product_combo = ttk.Combobox(form, textvariable=self.product_var, state='readonly', width=44, font=('Segoe UI', 10))
        prods = get_products()
        self.product_dict = {f"{p[1]} ({p[2]}) — {p[3]:.2f} руб.": (p[0], p[3]) for p in prods}
        self.product_combo['values'] = list(self.product_dict.keys())
        if prods: self.product_combo.current(0)
        self.product_combo.grid(row=1, column=1, sticky=EW, pady=(10, 2))

        # Количество
        lbl(2, "Количество")
        self.quantity_var = DoubleVar(value=1.0)
        ttk.Entry(form, textvariable=self.quantity_var, width=18, font=('Segoe UI', 11)).grid(row=2, column=1, sticky=W, pady=(10, 2))

        # Дата
        lbl(3, "Дата (ГГГГ-ММ-ДД)")
        self.date_var = StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(form, textvariable=self.date_var, width=18, font=('Segoe UI', 11)).grid(row=3, column=1, sticky=W, pady=(10, 2))

        # Смена
        lbl(4, "Смена")
        self.shift_var = IntVar(value=1)
        ttk.Spinbox(form, from_=1, to=2, textvariable=self.shift_var, width=6, font=('Segoe UI', 11)).grid(row=4, column=1, sticky=W, pady=(10, 2))

        # Оператор
        lbl(5, "Оператор")
        self.operator_var = StringVar()
        if user_role == 'admin':
            users_list = get_users('all')
            self.user_dict = {f"{u[1]} ({u[2]})": u[0] for u in users_list}
            self.operator_combo = ttk.Combobox(form, textvariable=self.operator_var, state='readonly', width=38, font=('Segoe UI', 10))
            self.operator_combo['values'] = list(self.user_dict.keys())
            if sale_data and sale_data[9]:
                for key in self.user_dict:
                    if key.startswith(sale_data[9]):
                        self.operator_var.set(key); break
            else:
                if current_user_id:
                    for key, uid in self.user_dict.items():
                        if uid == current_user_id:
                            self.operator_var.set(key); break
                elif self.user_dict:
                    self.operator_combo.current(0)
            self.operator_combo.grid(row=5, column=1, sticky=EW, pady=(10, 2))
        else:
            for u in get_users('all'):
                if u[0] == current_user_id:
                    self.operator_var.set(f"{u[1]} ({u[2]})"); break
            Label(form, textvariable=self.operator_var, font=('Segoe UI', 10),
                  bg=CARD_COLOR, fg=TEXT_COLOR, relief='flat', anchor=W, padx=8, pady=5
                  ).grid(row=5, column=1, sticky=EW, pady=(10, 2))

        dialog_footer(self.win, "✔  Сохранить", self.save, 'success', self.win.destroy)
        if sale_data: self._load_sale_data()

    def _get_station_id_by_name(self, name):
        for sid, sname in get_stations('admin', None):
            if sname == name: return sid
        return None

    def _load_sale_data(self):
        self.date_var.set(self.sale_data[1])
        product_full = f"{self.sale_data[2]} ({self.sale_data[3]}) — {self.sale_data[5]:.2f} руб."
        if product_full in self.product_dict: self.product_var.set(product_full)
        self.quantity_var.set(self.sale_data[4])
        self.shift_var.set(self.sale_data[7])

    def save(self):
        try:
            if self.user_role == 'admin' and not self.sale_data:
                st_str = self.station_var.get()
                if not st_str: messagebox.showerror("Ошибка", "Выберите АЗС"); return
                station_id = int(st_str.split(' - ')[0])
            else:
                station_id = self._get_station_id_by_name(self.sale_data[8]) if self.sale_data else self.allowed_station_id

            product_full = self.product_var.get()
            if not product_full or product_full not in self.product_dict:
                messagebox.showerror("Ошибка", "Выберите товар"); return
            product_id, _ = self.product_dict[product_full]

            quantity = self.quantity_var.get()
            if quantity <= 0: messagebox.showerror("Ошибка", "Количество > 0"); return

            sale_date = self.date_var.get().strip()
            try: datetime.strptime(sale_date, "%Y-%m-%d")
            except: messagebox.showerror("Ошибка", "Формат даты: ГГГГ-ММ-ДД"); return

            shift = self.shift_var.get()
            if shift not in (1, 2): messagebox.showerror("Ошибка", "Смена 1 или 2"); return

            if self.user_role == 'admin':
                op_key = self.operator_var.get()
                if not op_key: messagebox.showerror("Ошибка", "Выберите оператора"); return
                user_id = self.user_dict[op_key]
            else:
                user_id = self.current_user_id

            if self.sale_data:
                update_sale(self.sale_data[0], product_id, quantity, sale_date, shift, user_id)
                log_action(self.current_user_id, f"Редактирование продажи ID={self.sale_data[0]}")
            else:
                add_sale(station_id, product_id, quantity, sale_date, shift, user_id)
                log_action(self.current_user_id, f"Добавление продажи: {product_full[:30]}, {quantity}, {sale_date}")
            self.result = True
            self.win.destroy()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


# ══════════════════════════════════════════════════════════════
#  СМЕНА ПАРОЛЯ (для себя)
# ══════════════════════════════════════════════════════════════
class ChangePasswordWindow:
    def __init__(self, parent, user_id):
        self.user_id = user_id
        self.win = Toplevel(parent.root)
        self.win.title("Смена пароля")
        self.win.geometry("440x440")
        self.win.configure(bg=BG_COLOR)
        self.win.grab_set()
        self.win.resizable(False, False)
        apply_global_style()

        dialog_header(self.win, "🔑", "Смена пароля")
        form = Frame(self.win, bg=BG_COLOR, padx=36, pady=24)
        form.pack(fill=BOTH, expand=True)
        form.columnconfigure(0, weight=1)

        def field(r, text):
            Label(form, text=text, font=('Segoe UI', 10, 'bold'),
                  bg=BG_COLOR, fg=TEXT_DIM).grid(row=r*2, column=0, sticky=W, pady=(14, 2))
            e = styled_entry(form, width=26)
            e.grid(row=r*2+1, column=0, sticky=EW, ipady=8)
            return e

        self.old_pass     = field(0, "Текущий пароль")
        self.new_pass     = field(1, "Новый пароль")
        self.confirm_pass = field(2, "Подтверждение")

        dialog_footer(self.win, "✔  Сменить", self.change, 'success', self.win.destroy)

    def change(self):
        old = self.old_pass.get(); new = self.new_pass.get(); confirm = self.confirm_pass.get()
        if not old or not new: messagebox.showerror("Ошибка", "Заполните все поля"); return
        if new != confirm: messagebox.showerror("Ошибка", "Пароли не совпадают"); return
        if len(new) < 3: messagebox.showerror("Ошибка", "Минимум 3 символа"); return
        if update_password(self.user_id, old, new):
            log_action(self.user_id, "Смена пароля")
            messagebox.showinfo("Успех", "Пароль успешно изменён")
            self.win.destroy()
        else:
            messagebox.showerror("Ошибка", "Неверный текущий пароль")


# ══════════════════════════════════════════════════════════════
#  УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ (только для admin)
# ══════════════════════════════════════════════════════════════
class UserManagementWindow:
    def __init__(self, parent):
        self.parent = parent
        self.win = Toplevel(parent.root)
        self.win.title("Управление пользователями")
        self.win.geometry("760x540")
        self.win.configure(bg=BG_COLOR)
        self.win.grab_set()
        apply_global_style()

        dialog_header(self.win, "👥", "Управление пользователями")

        # Панель кнопок
        btn_panel = Frame(self.win, bg=PANEL_COLOR, height=50)
        btn_panel.pack(fill=X); btn_panel.pack_propagate(False)
        Frame(btn_panel, bg=BORDER_COLOR, height=1).pack(fill=X, side=BOTTOM)
        bp = Frame(btn_panel, bg=PANEL_COLOR)
        bp.pack(side=LEFT, padx=12, pady=8)
        self.set_pwd_btn  = mk_btn(bp, "🔑 Установить пароль", self.set_password, 'accent', DISABLED)
        self.set_pwd_btn.pack(side=LEFT, padx=(0, 6))
        mk_btn(bp, "🔄 Обновить", self.refresh, 'normal').pack(side=LEFT)

        # Индикатор запросов
        self.req_lbl = Label(btn_panel, text="", font=('Segoe UI', 10, 'bold'),
                             bg=PANEL_COLOR, fg=DANGER_COLOR)
        self.req_lbl.pack(side=RIGHT, padx=16)

        # Таблица пользователей
        tbl_frame = Frame(self.win, bg=BG_COLOR)
        tbl_frame.pack(fill=BOTH, expand=True, padx=12, pady=8)

        cols = ('ID', 'ФИО', 'Логин', 'Роль', 'АЗС', 'Запрос пароля')
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show='headings', selectmode='browse')
        self.tree.column('ID',     width=40,  anchor='center')
        self.tree.column('ФИО',   width=180, anchor='w')
        self.tree.column('Логин', width=110, anchor='center')
        self.tree.column('Роль',  width=80,  anchor='center')
        self.tree.column('АЗС',   width=170, anchor='w')
        self.tree.column('Запрос пароля', width=140, anchor='center')
        for col in cols:
            self.tree.heading(col, text=col)

        self.tree.tag_configure('admin',   foreground='#f5c842')
        self.tree.tag_configure('user',    foreground=TEXT_COLOR)
        self.tree.tag_configure('pending', foreground=DANGER_COLOR)

        vsb = ttk.Scrollbar(tbl_frame, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        vsb.pack(side=RIGHT, fill=Y)

        self.tree.bind('<<TreeviewSelect>>', self._on_select)

        # Закрыть
        Frame(self.win, bg=BORDER_COLOR, height=1).pack(fill=X)
        close_bar = Frame(self.win, bg=PANEL_COLOR, height=50)
        close_bar.pack(fill=X); close_bar.pack_propagate(False)
        mk_btn(close_bar, "✕  Закрыть", self.win.destroy, 'normal', padx=18, pady=8).pack(side=RIGHT, padx=14, pady=9)

        self.refresh()

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        users = get_users('all')
        # Загружаем pending-запросы
        pending_requests = {r[1]: r for r in get_reset_requests()}  # keyed by user_id

        for u in users:
            uid, full_name, username, role, station_id = u
            station_name = get_station_name(station_id) if station_id else "—"
            has_req = uid in pending_requests
            req_text = f"⚠ {pending_requests[uid][4][:16]}" if has_req else "—"
            tags = ('pending',) if has_req else (role,)
            self.tree.insert('', END, values=(uid, full_name, username, role, station_name, req_text), tags=tags)

        pending_count = get_pending_reset_count()
        if pending_count > 0:
            self.req_lbl.config(text=f"⚠  Запросов на сброс: {pending_count}")
        else:
            self.req_lbl.config(text="")
        self._on_select()

    def _on_select(self, event=None):
        sel = self.tree.selection()
        self.set_pwd_btn.config(state=NORMAL if sel else DISABLED,
                                bg=ACCENT_COLOR if sel else '#2a2d3e',
                                fg='white' if sel else TEXT_DIM)

    def set_password(self):
        sel = self.tree.selection()
        if not sel: return
        values = self.tree.item(sel[0])['values']
        user_id = values[0]
        full_name = values[1]
        username = values[2]

        # Диалог ввода нового пароля
        dlg = AdminSetPasswordDialog(self, user_id, full_name, username)
        self.win.wait_window(dlg.win)
        if dlg.result:
            self.refresh()
            self.parent.refresh_actions()


class AdminSetPasswordDialog:
    def __init__(self, parent, user_id, full_name, username):
        self.user_id = user_id
        self.result = False
        self.win = Toplevel(parent.win)
        self.win.title("Установить пароль")
        self.win.geometry("440x380")
        self.win.configure(bg=BG_COLOR)
        self.win.grab_set()
        self.win.resizable(False, False)
        apply_global_style()

        dialog_header(self.win, "🔑", "Установить новый пароль")

        form = Frame(self.win, bg=BG_COLOR, padx=36, pady=24)
        form.pack(fill=BOTH, expand=True)
        form.columnconfigure(0, weight=1)

        Label(form, text=f"Пользователь: {full_name} ({username})",
              font=('Segoe UI', 10), bg=BG_COLOR, fg=TEXT_DIM).grid(row=0, column=0, sticky=W, pady=(0, 18))

        def field(r, text):
            Label(form, text=text, font=('Segoe UI', 10, 'bold'),
                  bg=BG_COLOR, fg=TEXT_DIM).grid(row=r*2+1, column=0, sticky=W, pady=(10, 2))
            e = styled_entry(form, width=26)
            e.grid(row=r*2+2, column=0, sticky=EW, ipady=8)
            return e

        self.new_pass     = field(0, "Новый пароль")
        self.confirm_pass = field(1, "Подтверждение")

        # Проверяем есть ли pending запрос
        pending = get_reset_requests()
        self.req_id = None
        for r in pending:
            if r[1] == user_id:
                self.req_id = r[0]; break
        if self.req_id:
            Label(form, text="⚠  Пользователь запрашивал сброс пароля",
                  font=('Segoe UI', 9, 'italic'), bg=BG_COLOR, fg=WARNING_COLOR
                  ).grid(row=5, column=0, sticky=W, pady=(10, 0))

        dialog_footer(self.win, "✔  Установить", self.save, 'success', self.win.destroy)

    def save(self):
        new = self.new_pass.get(); confirm = self.confirm_pass.get()
        if not new: messagebox.showerror("Ошибка", "Введите новый пароль"); return
        if new != confirm: messagebox.showerror("Ошибка", "Пароли не совпадают"); return
        if len(new) < 3: messagebox.showerror("Ошибка", "Минимум 3 символа"); return
        admin_set_password(self.user_id, new)
        if self.req_id:
            resolve_reset_request(self.req_id)
        messagebox.showinfo("✔  Готово", "Пароль успешно установлен")
        self.result = True
        self.win.destroy()


# ══════════════════════════════════════════════════════════════
#  ДИАЛОГ ОТЧЁТА
# ══════════════════════════════════════════════════════════════
class ReportDialog:
    def __init__(self, parent, station_ids, user_role, user_name, current_user_id):
        self.station_ids = station_ids
        self.user_role = user_role
        self.user_name = user_name
        self.current_user_id = current_user_id
        self.result = False

        self.win = Toplevel(parent.root)
        self.win.title("Параметры отчёта")
        self.win.geometry("480x420")
        self.win.configure(bg=BG_COLOR)
        self.win.grab_set()
        self.win.resizable(False, False)
        apply_global_style()

        dialog_header(self.win, "📊", "Параметры отчёта Excel")

        form = Frame(self.win, bg=BG_COLOR, padx=36, pady=20)
        form.pack(fill=BOTH, expand=True)
        form.columnconfigure(1, weight=1)

        def lbl(r, text):
            Label(form, text=text, font=('Segoe UI', 10, 'bold'),
                  bg=BG_COLOR, fg=TEXT_DIM, anchor=W).grid(row=r, column=0, sticky=W, pady=(12, 2), padx=(0, 14))

        # Год
        lbl(0, "Год:")
        self.year_var = IntVar(value=datetime.now().year)
        ttk.Entry(form, textvariable=self.year_var, width=10, font=('Segoe UI', 11)).grid(row=0, column=1, sticky=W, pady=(12, 2))

        # Месяц
        lbl(1, "Месяц:")
        self.month_var = IntVar(value=datetime.now().month)
        ttk.Spinbox(form, from_=1, to=12, textvariable=self.month_var, width=6, font=('Segoe UI', 11)).grid(row=1, column=1, sticky=W, pady=(12, 2))

        # Сотрудник
        lbl(2, "Сотрудник:")
        self.employee_var = StringVar()
        self.employee_combo = ttk.Combobox(form, textvariable=self.employee_var, state='readonly',
                                            width=30, font=('Segoe UI', 10))
        users_list = get_users('all')
        self.user_map = {"— Все сотрудники —": None}
        for u in users_list:
            self.user_map[f"{u[1]} ({u[2]})"] = u[0]
        self.employee_combo['values'] = list(self.user_map.keys())
        self.employee_combo.current(0)
        self.employee_combo.grid(row=2, column=1, sticky=EW, pady=(12, 2))

        # Описание
        Frame(form, bg=BORDER_COLOR, height=1).grid(row=3, column=0, columnspan=2, sticky=EW, pady=(18, 12))
        Label(form, text="Отчёт будет сохранён в формате .xlsx с\nцветовым оформлением и итоговыми суммами.",
              font=('Segoe UI', 9), bg=BG_COLOR, fg=TEXT_DIM, justify=LEFT
              ).grid(row=4, column=0, columnspan=2, sticky=W)

        dialog_footer(self.win, "📊  Создать отчёт", self.ok, 'accent', self.win.destroy)

    def ok(self):
        year = self.year_var.get()
        month = self.month_var.get()
        if not (1 <= month <= 12 and year > 2000):
            messagebox.showerror("Ошибка", "Некорректный год или месяц"); return
        emp_key = self.employee_var.get()
        emp_uid = self.user_map.get(emp_key)
        emp_name = None if emp_uid is None else emp_key.split(" (")[0]
        self.win.destroy()
        generate_excel_report(self.station_ids, year, month, self.user_role,
                               self.user_name, emp_uid, emp_name)


# ══════════════════════════════════════════════════════════════
#  ГЛАВНОЕ ПРИЛОЖЕНИЕ
# ══════════════════════════════════════════════════════════════
class MainApp:
    def __init__(self, user):
        self.user = user  # (id, username, full_name, role, station_id)
        self.root = Tk()
        self.root.title(f"АЗС-Магазин  —  {user[2]}")
        self.root.geometry("1480x860")
        self.root.configure(bg=BG_COLOR)
        self.root.minsize(1100, 660)
        apply_global_style()

        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth(); sh = self.root.winfo_screenheight()
        self.root.geometry(f"1480x860+{max(0,(sw-1480)//2)}+{max(0,(sh-860)//2)}")

        self._build_titlebar()
        self._build_toolbar()
        self._build_body()
        self._build_statusbar()

        self.stations = get_stations(self.user[3], self.user[4])
        self.station_combo['values'] = [f"{sid} - {name}" for sid, name in self.stations]
        if self.stations:
            self.station_combo.current(0)
            self.on_station_selected()

        self.refresh_actions()
        log_action(self.user[0], "Вход в систему")
        self._check_pending_requests()
        self.root.mainloop()

    def _check_pending_requests(self):
        """Проверяем наличие запросов на сброс пароля при входе админа"""
        if self.user[3] == 'admin':
            cnt = get_pending_reset_count()
            if cnt > 0:
                self.root.after(800, lambda: messagebox.showinfo(
                    "⚠  Запросы на сброс пароля",
                    f"Есть {cnt} запрос(а/ов) на сброс пароля от пользователей.\n\n"
                    "Перейдите: Управление пользователями → выберите пользователя → Установить пароль"))

    # ── Title bar ─────────────────────────────────────────────
    def _build_titlebar(self):
        bar = Frame(self.root, bg=PRIMARY_COLOR, height=58)
        bar.pack(fill=X); bar.pack_propagate(False)

        left = Frame(bar, bg=PRIMARY_COLOR)
        left.pack(side=LEFT, padx=18, pady=8)
        Label(left, text="⛽", font=('Segoe UI Emoji', 22), bg=PRIMARY_COLOR, fg=ACCENT_COLOR).pack(side=LEFT)
        Label(left, text="  АЗС-МАГАЗИН", font=('Segoe UI', 16, 'bold'), bg=PRIMARY_COLOR, fg=TEXT_COLOR).pack(side=LEFT)

        right = Frame(bar, bg=PRIMARY_COLOR)
        right.pack(side=RIGHT, padx=18)

        role_icon = "👑" if self.user[3] == 'admin' else "👤"
        Label(right, text=f"{role_icon}  {self.user[2]}",
              font=('Segoe UI', 11), bg=PRIMARY_COLOR, fg=TEXT_DIM).pack(side=LEFT, padx=(0, 14))

        if self.user[3] == 'admin':
            self.users_btn = Button(right, text="👥 Пользователи", font=('Segoe UI', 10, 'bold'),
                   bg=ACCENT2_COLOR, fg='white', activebackground='#6a4ce0',
                   activeforeground='white', relief='flat', bd=0, padx=12, pady=6,
                   cursor='hand2', command=self.open_user_management)
            self.users_btn.pack(side=LEFT, padx=4)



        Button(right, text="⏏  Выход", font=('Segoe UI', 10, 'bold'),
               bg=DANGER_COLOR, fg='white', activebackground='#c0392b',
               activeforeground='white', relief='flat', bd=0, padx=12, pady=6,
               cursor='hand2', command=self.logout).pack(side=LEFT, padx=4)

        Frame(self.root, bg=ACCENT_COLOR, height=2).pack(fill=X)

    # ── Toolbar ───────────────────────────────────────────────
    def _build_toolbar(self):
        bar = Frame(self.root, bg=PANEL_COLOR, height=62)
        bar.pack(fill=X); bar.pack_propagate(False)

        inner = Frame(bar, bg=PANEL_COLOR)
        inner.pack(side=LEFT, fill=Y, padx=14, pady=10)

        Label(inner, text="Выберите АЗС:", font=('Segoe UI', 10, 'bold'),
              bg=PANEL_COLOR, fg=TEXT_DIM).pack(side=LEFT, padx=(0, 8))

        self.station_var = StringVar()
        self.station_combo = ttk.Combobox(inner, textvariable=self.station_var,
                                           state='readonly', width=36, font=('Segoe UI', 11))
        self.station_combo.pack(side=LEFT, padx=(0, 14))
        self.station_combo.bind('<<ComboboxSelected>>', self.on_station_selected)

        self.add_btn    = mk_btn(inner, "➕ Добавить",   self.add_sale,       'success')
        self.edit_btn   = mk_btn(inner, "✏  Изменить",   self.edit_sale,      'accent',  DISABLED)
        self.delete_btn = mk_btn(inner, "🗑  Удалить",    self.delete_sale,    'danger',  DISABLED)
        self.export_btn = mk_btn(inner, "📊 Excel",       self.export_excel,   'warn',    DISABLED)
        self.report_btn = mk_btn(inner, "📋 Отчёт",       self.monthly_report, 'normal')


        for b in (self.add_btn, self.edit_btn, self.delete_btn,
                  self.export_btn, self.report_btn):
            b.pack(side=LEFT, padx=3)

    # ── Body ──────────────────────────────────────────────────
    def _build_body(self):
        body = Frame(self.root, bg=BG_COLOR)
        body.pack(fill=BOTH, expand=True, padx=10, pady=(10, 0))

        # ── Таблица ──
        left_panel = Frame(body, bg=BG_COLOR)
        left_panel.pack(side=LEFT, fill=BOTH, expand=True)

        cols = ('ID', 'Дата', 'Товар', 'Категория', 'Кол-во', 'Цена', 'Сумма', 'Смена', 'АЗС', 'Оператор')
        self.tree = ttk.Treeview(left_panel, columns=cols, show='headings', selectmode='browse')
        self.tree.column('ID', width=0, stretch=False, minwidth=0)
        for col, w, anc in [
            ('Дата',100,'center'),('Товар',170,'w'),('Категория',110,'center'),
            ('Кол-во',80,'center'),('Цена',80,'center'),('Сумма',95,'center'),
            ('Смена',60,'center'),('АЗС',165,'w'),('Оператор',130,'w')]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor=anc, minwidth=40)

        self.tree.tag_configure('Топливо', background='#1f1420', foreground='#f08080')
        self.tree.tag_configure('Снеки',   background='#1e1c12', foreground='#f5c842')
        self.tree.tag_configure('Напитки', background='#111c2a', foreground='#7ec8e3')
        self.tree.tag_configure('Разное',  background=PANEL_COLOR, foreground=TEXT_COLOR)

        vsb = ttk.Scrollbar(left_panel, orient=VERTICAL,   command=self.tree.yview)
        hsb = ttk.Scrollbar(left_panel, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        left_panel.grid_rowconfigure(0, weight=1)
        left_panel.grid_columnconfigure(0, weight=1)

        self.tree.bind('<<TreeviewSelect>>', self.on_row_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_sale())

        # ── Боковая панель ──
        sidebar = Frame(body, bg=PANEL_COLOR, width=360)
        sidebar.pack(side=RIGHT, fill=Y, padx=(10, 0))
        sidebar.pack_propagate(False)

        Frame(sidebar, bg=ACCENT_COLOR, height=3).pack(fill=X)
        Label(sidebar, text="  📋  Последние действия",
              font=('Segoe UI', 11, 'bold'), bg=PANEL_COLOR, fg=TEXT_COLOR, anchor=W
              ).pack(fill=X, pady=(12, 4), padx=10)
        Frame(sidebar, bg=BORDER_COLOR, height=1).pack(fill=X, padx=10)

        list_frame = Frame(sidebar, bg=PANEL_COLOR)
        list_frame.pack(fill=BOTH, expand=True, padx=6, pady=6)

        self.actions_listbox = Listbox(
            list_frame, bg=CARD_COLOR, fg=TEXT_COLOR,
            selectbackground=ACCENT_COLOR, selectforeground='white',
            font=('Segoe UI', 9), relief='flat', bd=0,
            highlightthickness=0, activestyle='none')
        scroll_a = ttk.Scrollbar(list_frame, orient=VERTICAL, command=self.actions_listbox.yview)
        hscroll_a = ttk.Scrollbar(list_frame, orient=HORIZONTAL, command=self.actions_listbox.xview)
        self.actions_listbox.configure(yscrollcommand=scroll_a.set, xscrollcommand=hscroll_a.set)
        self.actions_listbox.grid(row=0, column=0, sticky='nsew')
        scroll_a.grid(row=0, column=1, sticky='ns')
        hscroll_a.grid(row=1, column=0, sticky='ew')
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

    # ── Status bar ────────────────────────────────────────────
    def _build_statusbar(self):
        bar = Frame(self.root, bg=PRIMARY_COLOR, height=32)
        bar.pack(fill=X, side=BOTTOM); bar.pack_propagate(False)
        Frame(bar, bg=ACCENT_COLOR, height=2).pack(fill=X, side=TOP)
        self.status = Label(bar, text="  Готово", font=('Segoe UI', 9),
                            bg=PRIMARY_COLOR, fg=TEXT_DIM, anchor=W)
        self.status.pack(side=LEFT, fill=Y, padx=8)

    # ── Event handlers ────────────────────────────────────────
    def on_row_select(self, event=None):
        sel = self.tree.selection()
        if sel:
            self.edit_btn.config(state=NORMAL,   bg=ACCENT_COLOR, fg='white', activebackground='#3a7aed')
            self.delete_btn.config(state=NORMAL, bg=DANGER_COLOR, fg='white', activebackground='#c0392b')
        else:
            self.edit_btn.config(state=DISABLED,   bg='#2a2d3e', fg=TEXT_DIM)
            self.delete_btn.config(state=DISABLED, bg='#2a2d3e', fg=TEXT_DIM)

    def on_station_selected(self, event=None):
        selected = self.station_var.get()
        if not selected: return
        try: station_id = int(selected.split(' - ')[0])
        except: return

        station_ids = [station_id] if self.user[3] == 'admin' else [self.user[4]]
        rows = get_sales(station_ids)

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.edit_btn.config(state=DISABLED,   bg='#2a2d3e', fg=TEXT_DIM)
        self.delete_btn.config(state=DISABLED, bg='#2a2d3e', fg=TEXT_DIM)

        total_sum = 0
        for row in rows:
            self.tree.insert('', END, values=row, tags=(row[3],))
            total_sum += row[6]

        self.status.config(text=f"  ✔  Записей: {len(rows)}   |   Итого: {total_sum:,.2f} руб.")
        if rows:
            self.export_btn.config(state=NORMAL, bg=WARNING_COLOR, fg='white', activebackground='#d68910')
        else:
            self.export_btn.config(state=DISABLED, bg='#2a2d3e', fg=TEXT_DIM)

    # ── Actions ───────────────────────────────────────────────
    def add_sale(self):
        w = EditSaleWindow(self, sale_data=None,
                           allowed_station_id=self.user[4] if self.user[3]=='user' else None,
                           user_role=self.user[3], current_user_id=self.user[0])
        self.root.wait_window(w.win)
        if w.result: self.on_station_selected(); self.refresh_actions()

    def edit_sale(self):
        sel = self.tree.selection()
        if not sel: return
        sale_data = self.tree.item(sel[0])['values']
        if self.user[3] == 'user' and sale_data[8] != get_station_name(self.user[4]):
            messagebox.showerror("Ошибка", "Нельзя редактировать продажи другой АЗС"); return
        w = EditSaleWindow(self, sale_data=sale_data, allowed_station_id=self.user[4],
                           user_role=self.user[3], current_user_id=self.user[0])
        self.root.wait_window(w.win)
        if w.result: self.on_station_selected(); self.refresh_actions()

    def delete_sale(self):
        sel = self.tree.selection()
        if not sel: return
        if not messagebox.askyesno("Подтверждение", "Удалить выбранную запись?"):  return
        if not messagebox.askyesno("Подтверждение", "Действие необратимо. Продолжить?"): return
        item = self.tree.item(sel[0])['values']
        if self.user[3] == 'user' and item[8] != get_station_name(self.user[4]):
            messagebox.showerror("Ошибка", "Нельзя удалять продажи другой АЗС"); return
        delete_sale(item[0])
        log_action(self.user[0], f"Удаление продажи ID={item[0]}")
        self.on_station_selected(); self.refresh_actions()

    def export_excel(self):
        """Быстрый экспорт текущего вида в Excel"""
        if not EXCEL_SUPPORT:
            messagebox.showerror("Ошибка", "Установите openpyxl: pip install openpyxl"); return
        selected = self.station_var.get()
        if not selected: return
        try: station_id = int(selected.split(' - ')[0])
        except: return
        station_ids = [station_id] if self.user[3]=='admin' else [self.user[4]]

        rows = get_sales(station_ids)
        if not rows: messagebox.showinfo("Экспорт", "Нет данных для экспорта"); return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить текущие данные в Excel",
            initialfile=f"export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        if not filename: return

        wb = Workbook(); wb.remove(wb.active)
        ws = wb.create_sheet("Данные")
        write_report_sheet(ws, rows, datetime.now().year, datetime.now().month)
        wb.save(filename)
        log_action(self.user[0], f"Экспорт Excel: {filename}")
        messagebox.showinfo("✔  Готово", f"Сохранено:\n{filename}")

    def monthly_report(self):
        selected = self.station_var.get()
        if not selected: return
        try: station_id = int(selected.split(' - ')[0])
        except: return
        station_ids = [station_id] if self.user[3]=='admin' else [self.user[4]]

        dlg = ReportDialog(self, station_ids, self.user[3], self.user[2], self.user[0])
        self.root.wait_window(dlg.win)
        self.refresh_actions()



    def change_password(self):
        ChangePasswordWindow(self, self.user[0])
        self.refresh_actions()

    def open_user_management(self):
        UserManagementWindow(self)
        self.refresh_actions()
        self._update_users_btn_badge()

    def _update_users_btn_badge(self):
        if self.user[3] == 'admin':
            cnt = get_pending_reset_count()
            text = f"👥 Пользователи ⚠{cnt}" if cnt > 0 else "👥 Пользователи"
            self.users_btn.config(text=text)

    def refresh_actions(self):
        self.actions_listbox.delete(0, END)
        for ts, user, action in get_last_actions(50):
            self.actions_listbox.insert(END, f"  {ts}  {user}: {action}")
        self._update_users_btn_badge()

    def logout(self):
        log_action(self.user[0], "Выход из системы")
        self.root.destroy()
        LoginWindow()


# ══════════════════════════════════════════════════════════════
#  ЗАПУСК
# ══════════════════════════════════════════════════════════════
if __name__ == '__main__':
    init_db()
    if not EXCEL_SUPPORT:
        # Показываем предупреждение только через Tk, не раньше
        root = Tk(); root.withdraw()
        messagebox.showwarning("openpyxl не найден",
                               "Для экспорта в Excel установите библиотеку:\n\n  pip install openpyxl\n\nОстальные функции работают в обычном режиме.")
        root.destroy()
    LoginWindow()

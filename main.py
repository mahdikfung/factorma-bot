import logging
import os
import sqlite3
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler, MessageHandler,
    ConversationHandler, ContextTypes, filters
)
from openpyxl import load_workbook
import jdatetime
import tempfile
import shutil
import subprocess

# ---------- Logging ----------
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ---------- Constants & Paths ----------
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(ROOT, "فاکتور رسمی مهدی خواجه.xlsx")
OUTPUT_DIR = os.path.join(ROOT, "invoices")
DB_PATH = os.path.join(ROOT, "factorma.db")

# ---------- Conversation States ----------
(
    MAIN_MENU,
    NEW_INVOICE_MENU,
    ENTER_NEW_CUSTOMER_NAME,
    ENTER_NEW_CUSTOMER_PHONE,
    ENTER_DATE_SHAMSI,
    ENTER_SEARCH_NAME,
    SELECT_EXISTING_CUSTOMER,
    ENTER_ITEM_DESC,
    ENTER_ITEM_QTY,
    ENTER_ITEM_UNIT,
    ENTER_ITEM_PRICE,
    ASK_MORE_ITEMS,
    REVIEW_INVOICE,
    CONFIRM_INVOICE,
    CUSTOMER_LIST_MENU,
    ENTER_SEARCH_CUSTOMER,
    VIEW_CUSTOMER,
) = range(17)

# ---------- DB helpers ----------
def init_db():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT NOT NULL
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_number INTEGER UNIQUE,
        date_shamsi TEXT,
        customer_id INTEGER,
        total_amount INTEGER,
        file_path TEXT,
        FOREIGN KEY(customer_id) REFERENCES customers(id)
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS invoice_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_id INTEGER,
        description TEXT,
        quantity REAL,
        unit TEXT,
        unit_price INTEGER,
        total_price INTEGER,
        FOREIGN KEY(invoice_id) REFERENCES invoices(id)
    )""")
    conn.commit()
    conn.close()


def get_next_invoice_number():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT MAX(id) FROM invoices")
    row = c.fetchone()
    conn.close()
    next_id = (row[0] or 0) + 1
    return 10000 + next_id


def create_customer(name, phone):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("INSERT INTO customers (name, phone) VALUES (?, ?)", (name, phone))
    conn.commit()
    cust_id = c.lastrowid
    conn.close()
    return cust_id


def search_customers(query):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        "SELECT id, name, phone FROM customers WHERE name LIKE ? ORDER BY name",
        ("%" + query + "%",),
    )
    rows = c.fetchall()
    conn.close()
    return rows


def get_customer(cust_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, name, phone FROM customers WHERE id=?", (cust_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return {"id": row[0], "name": row[1], "phone": row[2]}
    return None


def update_customer(cust_id, name, phone):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE customers SET name=?, phone=? WHERE id=?", (name, phone, cust_id))
    conn.commit()
    conn.close()


def save_invoice(invoice_number, date_shamsi, customer_id, total, file_path):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        """INSERT INTO invoices (invoice_number, date_shamsi, customer_id, total_amount, file_path)
           VALUES (?,?,?,?,?)""",
        (invoice_number, date_shamsi, customer_id, total, file_path),
    )
    inv_id = c.lastrowid
    conn.commit()
    conn.close()
    return inv_id


def save_items(invoice_id, items):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    for it in items:
        total = int(it["quantity"] * it["unit_price"])
        c.execute(
            """INSERT INTO invoice_items
               (invoice_id, description, quantity, unit, unit_price, total_price)
               VALUES (?,?,?,?,?,?)""",
            (invoice_id, it["description"], it["quantity"], it["unit"], it["unit_price"], total),
        )
    conn.commit()
    conn.close()


# ---------- PDF generation ----------
def generate_invoice_pdf(data):
    """Fill the Excel template and convert to PDF using libreoffice headless."""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Customer info (buyer section)
    ws["B11"] = data["customer_name"]
    ws["F14"] = data["customer_phone"]

    # Date and serial
    ws["F2"] = str(data["invoice_number"])
    ws["B4"] = data["date_shamsi"]

    # Items rows 16..(16+n-1)
    start_row = 16
    for i, item in enumerate(data["items"]):
        r = start_row + i
        ws[f"B{r}"] = i + 1
        ws[f"C{r}"] = item["description"]
        ws[f"D{r}"] = item["quantity"]
        ws[f"E{r}"] = item["unit"]
        ws[f"F{r}"] = item["unit_price"]
        ws[f"G{r}"] = int(item["quantity"] * item["unit_price"])

    tmp_dir = tempfile.mkdtemp()
    try:
        tmp_xlsx = os.path.join(tmp_dir, "invoice.xlsx")
        wb.save(tmp_xlsx)
        pdf_dir = os.path.join(tmp_dir, "pdf")
        os.makedirs(pdf_dir, exist_ok=True)
        subprocess.run(
            [
                "libreoffice", "--headless", "--convert-to", "pdf",
                "--outdir", pdf_dir, tmp_xlsx,
            ],
            check=True,
            timeout=120,
        )
        tmp_pdf = os.path.join(pdf_dir, "invoice.pdf")
        # Build final path
        shamsi_clean = data["date_shamsi"].replace("/", "")
        year_month = shamsi_clean[:6]  # 140506
        month_dir = os.path.join(OUTPUT_DIR, year_month)
        os.makedirs(month_dir, exist_ok=True)
        safe_name = data["customer_name"].replace(" ", "_")
        final_name = f"{data['invoice_number']}-{shamsi_clean}-{safe_name}.pdf"
        final_path = os.path.join(month_dir, final_name)
        shutil.move(tmp_pdf, final_path)
        return final_path
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------- Keyboards ----------
def main_menu_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📝 فاکتور جدید", callback_data="new_invoice")],
        [InlineKeyboardButton("📂 فاکتورهای سابق", callback_data="old_invoices")],
        [InlineKeyboardButton("👥 لیست مشتریان", callback_data="customer_list")],
    ])


def back_to_main_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🔙 بازگشت به منوی اصلی", callback_data="back_to_main")]
    ])


# ---------- Handlers ----------
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "خوش آمدید. لطفاً یک گزینه را انتخاب کنید:",
        reply_markup=main_menu_kb(),
    )
    return MAIN_MENU


async def main_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "new_invoice":
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("مشتری جدید", callback_data="new_customer")],
            [InlineKeyboardButton("مشتری قبلی", callback_data="existing_customer")],
            [InlineKeyboardButton("🔙 بازگشت", callback_data="back_to_main")],
        ])
        await query.edit_message_text(
            "آیا مشتری جدید است یا مشتری قبلی؟", reply_markup=kb
        )
        return NEW_INVOICE_MENU

    if data == "old_invoices":
        await query.edit_message_text(
            "نام مشتری را برای جستجو وارد کنید:",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔙 بازگشت", callback_data="back_to_main")]
            ]),
        )
        context.user_data["mode"] = "search_invoice"
        return ENTER_SEARCH_CUSTOMER

    if data == "customer_list":
        await query.edit_message_text(
            "لیست مشتریان - نام را برای جستجو وارد کنید (یا * همه مشتریان):",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔙 بازگشت", callback_data="back_to_main")]
            ]),
        )
        context.user_data["mode"] = "search_customer"
        return ENTER_SEARCH_CUSTOMER

    if data == "back_to_main":
        context.user_data.clear()
        await query.edit_message_text(
            "منوی اصلی:", reply_markup=main_menu_kb()
        )
        return MAIN_MENU


# ---- New invoice flow ----
async def new_invoice_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "new_customer":
        context.user_data["invoice"] = {"is_new": True, "items": []}
        await query.edit_message_text("نام و نام خانوادگی مشتری جدید را وارد کنید:")
        return ENTER_NEW_CUSTOMER_NAME

    if data == "existing_customer":
        context.user_data["invoice"] = {"is_new": False, "items": []}
        await query.edit_message_text("نام مشتری قبلی را برای جستجو وارد کنید:")
        return ENTER_SEARCH_NAME

    if data == "back_to_main":
        context.user_data.clear()
        await query.edit_message_text("منوی اصلی:", reply_markup=main_menu_kb())
        return MAIN_MENU


async def enter_new_customer_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    if not name:
        await update.message.reply_text("نام نمی‌تواند خالی باشد. دوباره وارد کنید:")
        return ENTER_NEW_CUSTOMER_NAME
    context.user_data["invoice"]["customer_name"] = name
    await update.message.reply_text("شماره تلفن مشتری را وارد کنید:")
    return ENTER_NEW_CUSTOMER_PHONE


async def enter_new_customer_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    phone = update.message.text.strip()
    if not phone:
        await update.message.reply_text("شماره تلفن نمی‌تواند خالی باشد. دوباره وارد کنید:")
        return ENTER_NEW_CUSTOMER_PHONE
    context.user_data["invoice"]["customer_phone"] = phone
    await update.message.reply_text(
        "تاریخ فاکتور را به صورت شمسی وارد کنید (مثال: 1405/06/06):"
    )
    return ENTER_DATE_SHAMSI


async def enter_date_shamsi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    date_text = update.message.text.strip()
    # Basic validation
    if not date_text or len(date_text.replace("/", "")) < 8:
        await update.message.reply_text(
            "تاریخ نامعتبر است. مثال: 1405/06/06"
        )
        return ENTER_DATE_SHAMSI
    context.user_data["invoice"]["date_shamsi"] = date_text
    # Save customer and start items
    inv = context.user_data["invoice"]
    cust_id = create_customer(inv["customer_name"], inv["customer_phone"])
    inv["customer_id"] = cust_id
    inv["invoice_number"] = get_next_invoice_number()
    context.user_data["current_item"] = {}
    await update.message.reply_text(
        "کالای اول را وارد کنید - شرح کالا چیست؟"
    )
    return ENTER_ITEM_DESC


async def enter_search_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    if not name:
        await update.message.reply_text("نام نمی‌تواند خالی باشد.")
        return ENTER_SEARCH_NAME
    rows = search_customers(name)
    if not rows:
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🔙 بازگشت", callback_data="back_to_main")]
        ])
        await update.message.reply_text(
            "مشتری یافت نشد. نام دیگری وارد کنید یا بازگشت بزنید.",
            reply_markup=kb,
        )
        return ENTER_SEARCH_NAME
    kb_rows = [
        [InlineKeyboardButton(f"{r[1]} - {r[2]}", callback_data=f"pickcust_{r[0]}")]
        for r in rows
    ]
    kb_rows.append([InlineKeyboardButton("🔙 بازگشت", callback_data="back_to_main")])
    await update.message.reply_text(
        "مشتری را انتخاب کنید:", reply_markup=InlineKeyboardMarkup(kb_rows)
    )
    return SELECT_EXISTING_CUSTOMER


async def select_existing_customer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "back_to_main":
        context.user_data.clear()
        await query.edit_message_text("منوی اصلی:", reply_markup=main_menu_kb())
        return MAIN_MENU
    cust_id = int(data.split("_", 1)[1])
    cust = get_customer(cust_id)
    if not cust:
        await query.edit_message_text("مشتری یافت نشد.", reply_markup=back_to_main_kb())
        return MAIN_MENU
    inv = context.user_data["invoice"]
    inv["customer_id"] = cust["id"]
    inv["customer_name"] = cust["name"]
    inv["customer_phone"] = cust["phone"]
    inv["invoice_number"] = get_next_invoice_number()
    context.user_data["current_item"] = {}
    await query.edit_message_text(
        f"مشتری انتخاب شد: {cust['name']}\n\n"
        f"تاریخ فاکتور را به صورت شمسی وارد کنید (مثال: 1405/06/06):"
    )
    # After customer selection we go to date input
    return ENTER_DATE_SHAMSI


# ---- Items flow ----
async def enter_item_desc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    desc = update.message.text.strip()
    if not desc:
        await update.message.reply_text("شرح کالا نمی‌تواند خالی باشد.")
        return ENTER_ITEM_DESC
    context.user_data["current_item"] = {"description": desc}
    await update.message.reply_text("تعداد کالا را وارد کنید (مثال: 5):")
    return ENTER_ITEM_QTY


async def enter_item_qty(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    try:
        qty = float(text)
        if qty <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text("عدد معتبر وارد کنید (مثلاً 5 یا 2.5):")
        return ENTER_ITEM_QTY
    context.user_data["current_item"]["quantity"] = qty
    await update.message.reply_text("واحد کالا را وارد کنید (مثال: عدد، کیلو، متر):")
    return ENTER_ITEM_UNIT


async def enter_item_unit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    unit = update.message.text.strip()
    if not unit:
        await update.message.reply_text("واحد نمی‌تواند خالی باشد.")
        return ENTER_ITEM_UNIT
    context.user_data["current_item"]["unit"] = unit
    await update.message.reply_text("قیمت واحد را به ریال وارد کنید (مثال: 150000):")
    return ENTER_ITEM_PRICE


async def enter_item_price(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip().replace(",", "").replace("،", "")
    try:
        price = int(text)
        if price < 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text("عدد معتبر وارد کنید (مثال: 150000):")
        return ENTER_ITEM_PRICE
    item = context.user_data["current_item"]
    item["unit_price"] = price
    item["total_price"] = int(item["quantity"] * price)
    context.user_data["invoice"]["items"].append(item)
    context.user_data["current_item"] = {}
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("➕ افزودن کالای بعدی", callback_data="add_more")],
        [InlineKeyboardButton("✅ پایان و بررسی", callback_data="finish_items")],
        [InlineKeyboardButton("❌ انصراف", callback_data="cancel_invoice")],
    ])
    await update.message.reply_text(
        f"کالا ثبت شد: {item['description']} - {item['quantity']} {item['unit']} - {item['total_price']:,} ریال\n\n"
        "می‌خواهید کالای دیگری اضافه کنید؟",
        reply_markup=kb,
    )
    return ASK_MORE_ITEMS


async def ask_more_items(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "add_more":
        await query.edit_message_text("شرح کالای جدید را وارد کنید:")
        return ENTER_ITEM_DESC
    if query.data == "cancel_invoice":
        context.user_data.clear()
        await query.edit_message_text("فاکتور لغو شد.", reply_markup=main_menu_kb())
        return MAIN_MENU
    # finish_items
    inv = context.user_data["invoice"]
    items = inv["items"]
    lines = []
    total = 0
    for i, it in enumerate(items, 1):
        line_total = int(it["quantity"] * it["unit_price"])
        total += line_total
        lines.append(
            f"{i}. {it['description']} | {it['quantity']} {it['unit']} | "
            f"{it['unit_price']:,} ریال | {line_total:,} ریال"
        )
    inv["total"] = total
    msg = (
        f"📋 پیش‌نمایش فاکتور:\n\n"
        f"شماره فاکتور: {inv['invoice_number']}\n"
        f"تاریخ: {inv['date_shamsi']}\n"
        f"خریدار: {inv['customer_name']} - {inv['customer_phone']}\n\n"
        f"کالاها:\n" + "\n".join(lines) + f"\n\n💰 جمع کل: {total:,} ریال\n\n"
        "آیا تأیید می‌کنید؟"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ تأیید و ارسال PDF", callback_data="confirm_invoice")],
        [InlineKeyboardButton("✏️ ویرایش کالاها", callback_data="edit_items")],
        [InlineKeyboardButton("❌ انصراف", callback_data="cancel_invoice")],
    ])
    await query.edit_message_text(msg, reply_markup=kb)
    return REVIEW_INVOICE


async def review_invoice_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "cancel_invoice":
        context.user_data.clear()
        await query.edit_message_text("فاکتور لغو شد.", reply_markup=main_menu_kb())
        return MAIN_MENU
    if data == "edit_items":
        # Simple: restart items, keep customer and date
        context.user_data["invoice"]["items"] = []
        context.user_data["current_item"] = {}
        await query.edit_message_text("کالاها پاک شدند. شرح کالای اول را وارد کنید:")
        return ENTER_ITEM_DESC
    # confirm
    inv = context.user_data["invoice"]
    pdf_data = {
        "customer_name": inv["customer_name"],
        "customer_phone": inv["customer_phone"],
        "invoice_number": inv["invoice_number"],
        "date_shamsi": inv["date_shamsi"],
        "items": inv["items"],
    }
    try:
        await query.edit_message_text("در حال ساخت PDF...")
        pdf_path = generate_invoice_pdf(pdf_data)
    except subprocess.TimeoutExpired:
        await query.edit_message_text("تبدیل PDF بیش از حد طول کشید. دوباره تلاش کنید.")
        return REVIEW_INVOICE
    except Exception as e:
        logger.exception("PDF generation failed")
        await query.edit_message_text(f"خطا در ساخت PDF: {e}")
        return REVIEW_INVOICE

    total = sum(int(it["quantity"] * it["unit_price"]) for it in inv["items"])
    inv_id = save_invoice(
        inv["invoice_number"], inv["date_shamsi"],
        inv["customer_id"], total, pdf_path,
    )
    save_items(inv_id, inv["items"])
    # Send PDF
    with open(pdf_path, "rb") as f:
        await context.bot.send_document(
            chat_id=query.message.chat_id,
            document=f,
            filename=os.path.basename(pdf_path),
        )
    context.user_data.clear()
    await query.message.reply_text(
        "✅ فاکتور با موفقیت ساخته و ارسال شد.",
        reply_markup=main_menu_kb(),
    )
    return MAIN_MENU


# ---- Customer list (search & view & edit) ----
async def enter_search_customer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    rows = []
    if name == "*":
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT id, name, phone FROM customers ORDER BY name LIMIT 50")
        rows = c.fetchall()
        conn.close()
    else:
        rows = search_customers(name)
    if not rows:
        await update.message.reply_text(
            "نتیجه‌ای یافت نشد. نام دیگری وارد کنید یا * برای همه بزنید.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔙 بازگشت", callback_data="back_to_main")]
            ]),
        )
        return ENTER_SEARCH_CUSTOMER
    kb_rows = [
        [InlineKeyboardButton(f"{r[1]} - {r[2]}", callback_data=f"viewcust_{r[0]}")]
        for r in rows
    ]
    kb_rows.append([InlineKeyboardButton("🔙 بازگشت", callback_data="back_to_main")])
    await update.message.reply_text(
        "مشتری را انتخاب کنید:", reply_markup=InlineKeyboardMarkup(kb_rows)
    )
    return VIEW_CUSTOMER


async def view_customer_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "back_to_main":
        context.user_data.clear()
        await query.edit_message_text("منوی اصلی:", reply_markup=main_menu_kb())
        return MAIN_MENU
    cust_id = int(data.split("_", 1)[1])
    cust = get_customer(cust_id)
    if not cust:
        await query.edit_message_text("مشتری یافت نشد.", reply_markup=back_to_main_kb())
        return MAIN_MENU
    context.user_data["viewing_customer_id"] = cust_id
    msg = (
        f"👤 اطلاعات مشتری:\n\n"
        f"نام: {cust['name']}\n"
        f"تلفن: {cust['phone']}\n"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("✏️ ویرایش نام", callback_data="edit_name")],
        [InlineKeyboardButton("✏️ ویرایش تلفن", callback_data="edit_phone")],
        [InlineKeyboardButton("🔙 بازگشت به لیست", callback_data="back_to_list")],
    ])
    await query.edit_message_text(msg, reply_markup=kb)
    return VIEW_CUSTOMER


async def customer_edit_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "back_to_list":
        # Re-show list - go back to search but use the same name if we have one
        await query.edit_message_text(
            "نام را برای جستجو وارد کنید (یا * برای همه):"
        )
        return ENTER_SEARCH_CUSTOMER
    cust_id = context.user_data.get("viewing_customer_id")
    if not cust_id:
        await query.edit_message_text("خطا. دوباره شروع کنید.", reply_markup=main_menu_kb())
        return MAIN_MENU
    context.user_data["editing_field"] = "name" if data == "edit_name" else "phone"
    field_name = "نام" if data == "edit_name" else "شماره تلفن"
    await query.edit_message_text(f"{field_name} جدید را وارد کنید:")
    return CONFIRM_INVOICE  # Reusing state for "waiting for new value"


async def customer_save_edit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Reuses CONFIRM_INVOICE state to capture the new value."""
    new_value = update.message.text.strip()
    cust_id = context.user_data.get("viewing_customer_id")
    field = context.user_data.get("editing_field")
    if not cust_id or not field:
        await update.message.reply_text("خطا. دوباره شروع کنید.", reply_markup=main_menu_kb())
        return MAIN_MENU
    if not new_value:
        await update.message.reply_text("مقدار نمی‌تواند خالی باشد. دوباره وارد کنید:")
        return CONFIRM_INVOICE
    cust = get_customer(cust_id)
    if field == "name":
        update_customer(cust_id, new_value, cust["phone"])
    else:
        update_customer(cust_id, cust["name"], new_value)
    # Refresh view (stay on customer page)
    cust = get_customer(cust_id)
    msg = (
        f"👤 اطلاعات مشتری (به‌روز شد):\n\n"
        f"نام: {cust['name']}\n"
        f"تلفن: {cust['phone']}\n"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("✏️ ویرایش نام", callback_data="edit_name")],
        [InlineKeyboardButton("✏️ ویرایش تلفن", callback_data="edit_phone")],
        [InlineKeyboardButton("🔙 بازگشت به لیست", callback_data="back_to_list")],
    ])
    await update.message.reply_text(msg, reply_markup=kb)
    return VIEW_CUSTOMER


# ---- Cancel ----
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("عملیات لغو شد.", reply_markup=main_menu_kb())
    return MAIN_MENU


# ---------- Main ----------
def main():
    init_db()
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN env var not set")
    app = Application.builder().token(token).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", cmd_start)],
        states={
            MAIN_MENU: [CallbackQueryHandler(main_menu_handler)],
            NEW_INVOICE_MENU: [CallbackQueryHandler(new_invoice_menu)],
            ENTER_NEW_CUSTOMER_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, enter_new_customer_name)
            ],
            ENTER_NEW_CUSTOMER_PHONE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, enter_new_customer_phone)
            ],
            ENTER_DATE_SHAMSI: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, enter_date_shamsi)
            ],
            ENTER_SEARCH_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, enter_search_name)
            ],
            SELECT_EXISTING_CUSTOMER: [
                CallbackQueryHandler(select_existing_customer)
            ],
            ENTER_ITEM_DESC: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, enter_item_desc)
            ],
            ENTER_ITEM_QTY: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, enter_item_qty)
            ],
            ENTER_ITEM_UNIT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, enter_item_unit)
            ],
            ENTER_ITEM_PRICE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, enter_item_price)
            ],
            ASK_MORE_ITEMS: [CallbackQueryHandler(ask_more_items)],
            REVIEW_INVOICE: [CallbackQueryHandler(review_invoice_handler)],
            ENTER_SEARCH_CUSTOMER: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, enter_search_customer)
            ],
            VIEW_CUSTOMER: [
                CallbackQueryHandler(view_customer_handler),
                CallbackQueryHandler(customer_edit_field, pattern="^(edit_name|edit_phone|back_to_list)$"),
            ],
            CONFIRM_INVOICE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, customer_save_edit)
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(conv)
    app.run_polling()


if __name__ == "__main__":
    main()
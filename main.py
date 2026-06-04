import os
import json
import tempfile
import logging
from datetime import datetime
from io import BytesIO

from flask import Flask, request

import qrcode
import gspread
from google.oauth2.service_account import Credentials
from num2words import num2words

from telegram import (
    Update,
    ReplyKeyboardMarkup,
    KeyboardButton,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

from docx import Document
from docx.shared import Inches


logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "asia_mx_secret")

if not BOT_TOKEN:
    raise RuntimeError("Не задан BOT_TOKEN / TELEGRAM_BOT_TOKEN")

app = Flask(__name__)
tg_app = Application.builder().token(BOT_TOKEN).build()

USER_STATE = {}

AGENT_NAME = "ЭЙЖА ЭМ ЭКС ТРЕЙД ЛИМИТЕД (ASIA MX TRADE LIMITED)"
AGENT_SHORT = "ЭЙЖА ЭМ ЭКС ТРЕЙД ЛИМИТЕД"
AGENT_INN = "9909722497"
AGENT_KPP = "270087001"
AGENT_ACCOUNT = "40807810350710000002"
AGENT_ACCOUNT_PRINT = "40807 810 3 5071 0000002"
AGENT_BANK = "ДАЛЬНЕВОСТОЧНЫЙ БАНК ПАО СБЕРБАНК"
AGENT_BIK = "040813608"
AGENT_CORR = "30101810600000000608"
AGENT_CORR_PRINT = "30101 810 6 0000 0000608"

MAIN_MENU = ReplyKeyboardMarkup(
    [
        [KeyboardButton("📄 Выставить счет")],
        [KeyboardButton("📑 Договор"), KeyboardButton("📎 Приложение")],
        [KeyboardButton("❌ Отмена")],
    ],
    resize_keyboard=True,
)


def get_sheet():
    if not GOOGLE_SHEET_ID or not GOOGLE_SERVICE_ACCOUNT_JSON:
        return None

    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(GOOGLE_SHEET_ID)


def get_next_invoice_number():
    sheet = get_sheet()
    if sheet is None:
        return str(int(datetime.now().strftime("%H%M%S")))

    try:
        ws = sheet.worksheet("СЧЕТА")
    except Exception:
        ws = sheet.add_worksheet(title="СЧЕТА", rows=1000, cols=10)
        ws.append_row(["invoice_number", "created_at"])

    values = ws.get_all_values()
    last_number = 14252

    for row in values[1:]:
        if row and row[0].isdigit():
            last_number = max(last_number, int(row[0]))

    new_number = last_number + 1
    ws.append_row([str(new_number), datetime.now().strftime("%d.%m.%Y %H:%M:%S")])
    return str(new_number)


def money_words(amount):
    rub = int(float(str(amount).replace(" ", "").replace(",", ".")))
    text = num2words(rub, lang="ru")
    return f"{text.capitalize()} рублей 00 копеек"


def format_amount(amount):
    value = float(str(amount).replace(" ", "").replace(",", "."))
    return f"{value:,.2f}".replace(",", " ")


def amount_kopecks(amount):
    value = float(str(amount).replace(" ", "").replace(",", "."))
    return str(int(round(value * 100)))


def make_qr(data):
    purpose = f"Оплата по счету №{data['invoice_number']} по договору {data['contract_number']} от {data['contract_date']}"
    qr_text = (
        f"ST00012|"
        f"Name={AGENT_SHORT}|"
        f"PersonalAcc={AGENT_ACCOUNT}|"
        f"BankName={AGENT_BANK}|"
        f"BIC={AGENT_BIK}|"
        f"CorrespAcc={AGENT_CORR}|"
        f"PayeeINN={AGENT_INN}|"
        f"KPP={AGENT_KPP}|"
        f"Purpose={purpose}|"
        f"Sum={amount_kopecks(data['amount'])}"
    )

    img = qrcode.make(qr_text)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    img.save(tmp.name)
    return tmp.name


def replace_text(doc, mapping):
    for p in doc.paragraphs:
        for key, value in mapping.items():
            if key in p.text:
                for run in p.runs:
                    run.text = run.text.replace(key, str(value))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in mapping.items():
                        if key in p.text:
                            for run in p.runs:
                                run.text = run.text.replace(key, str(value))


def create_invoice_docx(data):
    doc = Document()

    doc.add_paragraph(AGENT_BANK)
    doc.add_paragraph(f"БИК {AGENT_BIK}")
    doc.add_paragraph(f"Корсчёт {AGENT_CORR_PRINT}")
    doc.add_paragraph(f"ИНН {AGENT_INN}   КПП {AGENT_KPP}")
    doc.add_paragraph(f"Сч. № {AGENT_ACCOUNT_PRINT}")
    doc.add_paragraph(f"Получатель: {AGENT_SHORT}")

    doc.add_heading(
        f"Счет на оплату № {data['invoice_number']} от {data['invoice_date']}",
        level=1,
    )

    doc.add_paragraph(
        f"Поставщик (Исполнитель): {AGENT_NAME}, ИНН {AGENT_INN}, КПП {AGENT_KPP}, "
        "Unit 704c 7/f Block 3, Nan Fung Industrial City, 18 Tin Hau Road, Tuen Mun, Hong Kong"
    )

    doc.add_paragraph(f"Покупатель (Заказчик): {data['buyer_full_name']}")

    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    headers = ["№", "Товары (работы, услуги)", "Кол-во", "Ед.", "Цена", "Сумма"]

    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h

    row = table.add_row().cells
    service = (
        f"Оплата по договору {data['contract_number']} от {data['contract_date']}, "
        f"инвойс {data['foreign_invoice_number']}, за {data['car_name']}, "
        f"{data['identifier_type']} {data['identifier_value']}"
    )

    row[0].text = "1"
    row[1].text = service
    row[2].text = "1"
    row[3].text = "Шт"
    row[4].text = format_amount(data["amount"])
    row[5].text = format_amount(data["amount"])

    doc.add_paragraph(f"Итого: {format_amount(data['amount'])}")
    doc.add_paragraph("НДС не облагается: -")
    doc.add_paragraph(f"Всего к оплате: {format_amount(data['amount'])}")
    doc.add_paragraph(f"Всего наименований 1, на сумму {format_amount(data['amount'])} руб.")
    doc.add_paragraph(money_words(data["amount"]))

    qr_path = make_qr(data)
    doc.add_paragraph("QR-код для оплаты:")
    doc.add_picture(qr_path, width=Inches(1.7))

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    return tmp.name


def create_contract_docx(data):
    doc = Document()

    doc.add_paragraph("ASIA MX TRADE LIMITED")
    doc.add_paragraph("Unit 704c 7/f Block 3, Nan Fung Industrial City, 18 Tin Hau Road, Tuen Mun, Hong Kong")
    doc.add_paragraph("Email: general@asiamx.ltd")
    doc.add_paragraph("Web: asiamx.ltd")
    doc.add_paragraph("Tel: +852-9129-2658")

    doc.add_heading(f"АГЕНТСКИЙ ДОГОВОР № {data['contract_number']}", level=1)
    doc.add_paragraph("по организации оплаты автомобилей")
    doc.add_paragraph(f"г. Владивосток     «{data['contract_date']}»")

    doc.add_paragraph(
        f"{data['buyer_full_name']}, именуемый(ая) в дальнейшем «Принципал», "
        f"с одной стороны, и {AGENT_NAME}, зарегистрированная на территории САР Гонконг, "
        f"Unit 704c 7/f Block 3, Nan Fung Industrial City, 18 Tin Hau Road, Tuen Mun, Hong Kong, "
        f"ИНН {AGENT_INN}, КПП {AGENT_KPP}, в лице Антона Фалкона Пичардо, "
        f"действующего на основании Устава, именуемая в дальнейшем «Агент», "
        f"с другой стороны, заключили настоящий договор о нижеследующем."
    )

    doc.add_heading("1. Предмет договора", level=2)
    doc.add_paragraph(
        "1.1. Принципал поручает, а Агент принимает на себя обязательства по организации "
        "и проведению оплаты автомобилей, приобретаемых Принципалом."
    )
    doc.add_paragraph(
        "1.2. Агент осуществляет функции по приёму денежных средств от Принципала "
        "и их перечислению на счета поставщиков, указанных Принципалом."
    )

    doc.add_heading("2. Вознаграждение и порядок расчетов", level=2)
    doc.add_paragraph("2.1. Вознаграждение Агента составляет 1% от суммы каждого платежа.")
    doc.add_paragraph(
        "2.2. Оплата вознаграждения производится путем безналичного перечисления "
        "на расчетный счет Агента."
    )

    doc.add_heading("3. Банковские реквизиты Агента", level=2)
    doc.add_paragraph(
        f"Агент: ASIA MX TRADE LIMITED\n"
        f"Unit 704c 7/f Block 3, Nan Fung Industrial City, 18 Tin Hau Road, Tuen Mun, Hong Kong\n"
        f"Email: general@asiamx.ltd\n"
        f"Web: asiamx.ltd\n"
        f"Tel: +852-9129-2658\n\n"
        f"Банковские реквизиты (Россия):\n"
        f"Наименование: {AGENT_NAME}\n"
        f"ИНН: {AGENT_INN}, КПП: {AGENT_KPP}\n"
        f"Р/с: {AGENT_ACCOUNT_PRINT}\n"
        f"Банк: {AGENT_BANK}\n"
        f"БИК: {AGENT_BIK}\n"
        f"Корсчёт: {AGENT_CORR_PRINT}"
    )

    doc.add_paragraph("\nПодписи сторон:")
    doc.add_paragraph(f"Агент: _____________________ /Антон Фалкон Пичардо/")
    doc.add_paragraph(f"Принципал: _____________________ /{data['buyer_full_name']}/")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    return tmp.name


def create_appendix_docx(data):
    doc = Document()

    doc.add_heading(
        f"Приложение №1 к Агентскому договору № {data['contract_number']} от {data['contract_date']}",
        level=1,
    )
    doc.add_heading(
        f"Поручение Агенту №1 от {data['invoice_date']}",
        level=2,
    )

    doc.add_paragraph(
        f"{data['buyer_full_name']} (Принципал) поручает {AGENT_NAME} (Агент) "
        f"обеспечить своевременную оплату в пользу третьего лица по договору поставки."
    )

    table = doc.add_table(rows=8, cols=2)
    table.style = "Table Grid"

    rows = [
        ("Валюта-1", "Российский рубль"),
        ("Сумма платежа, Валюта-1", f"{format_amount(data['amount'])} руб."),
        (
            "Банковские реквизиты Агента",
            f"{AGENT_NAME}\nИНН: {AGENT_INN}, КПП: {AGENT_KPP}\nР/с: {AGENT_ACCOUNT_PRINT}\n"
            f"Банк: {AGENT_BANK}\nБИК: {AGENT_BIK}\nКорсчёт: {AGENT_CORR_PRINT}",
        ),
        ("Вознаграждение Агента", "1%"),
        ("Валюта-2", "Российский рубль"),
        ("Сумма после удержания вознаграждения", f"{format_amount(float(data['amount']) * 0.99)} руб."),
        ("Контрагент", data.get("contractor_name", "Не указано")),
        ("Другие условия", ""),
    ]

    for i, row_data in enumerate(rows):
        table.rows[i].cells[0].text = row_data[0]
        table.rows[i].cells[1].text = row_data[1]

    doc.add_paragraph("\nАгент:")
    doc.add_paragraph(f"{AGENT_NAME}\n_____________________ /Антон Фалкон Пичардо/")

    doc.add_paragraph("\nПринципал:")
    doc.add_paragraph(f"_____________________ /{data['buyer_full_name']}/")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    return tmp.name


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Бот документов ASIA MX запущен ✅\n\nВыберите действие:",
        reply_markup=MAIN_MENU,
    )


def reset_user(user_id):
    USER_STATE[user_id] = {"mode": None, "step": None, "data": {}}


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    reset_user(update.effective_user.id)
    await update.message.reply_text("Действие отменено.", reply_markup=MAIN_MENU)


async def start_invoice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    USER_STATE[user_id] = {
        "mode": "invoice",
        "step": "buyer_full_name",
        "data": {},
    }
    await update.message.reply_text("Введите ФИО покупателя:")


async def start_contract(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    state = USER_STATE.get(user_id, {"data": {}})
    data = state.get("data", {})

    if not data.get("buyer_full_name"):
        await update.message.reply_text("Сначала сформируйте счет, чтобы договор подтянул данные.")
        return

    docx = create_contract_docx(data)
    await update.message.reply_document(
        document=open(docx, "rb"),
        filename=f"Договор_{data['contract_number']}.docx",
        caption="Договор сформирован ✅",
    )


async def start_appendix(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    state = USER_STATE.get(user_id, {"data": {}})
    data = state.get("data", {})

    if not data.get("buyer_full_name"):
        await update.message.reply_text("Сначала сформируйте счет, чтобы приложение подтянуло данные.")
        return

    docx = create_appendix_docx(data)
    await update.message.reply_document(
        document=open(docx, "rb"),
        filename=f"Приложение_{data['contract_number']}.docx",
        caption="Приложение сформировано ✅",
    )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    user_id = update.effective_user.id

    if text == "📄 Выставить счет":
        await start_invoice(update, context)
        return

    if text == "📑 Договор":
        await start_contract(update, context)
        return

    if text == "📎 Приложение":
        await start_appendix(update, context)
        return

    if text == "❌ Отмена":
        await cancel(update, context)
        return

    state = USER_STATE.get(user_id)
    if not state or state.get("mode") != "invoice":
        await update.message.reply_text("Выберите действие:", reply_markup=MAIN_MENU)
        return

    step = state["step"]
    data = state["data"]

    if step == "buyer_full_name":
        data["buyer_full_name"] = text
        state["step"] = "contract_number"
        await update.message.reply_text("Введите номер договора:")
        return

    if step == "contract_number":
        data["contract_number"] = text
        state["step"] = "contract_date"
        await update.message.reply_text("Введите дату договора:")
        return

    if step == "contract_date":
        data["contract_date"] = text
        state["step"] = "foreign_invoice_number"
        await update.message.reply_text("Введите номер инвойса:")
        return

    if step == "foreign_invoice_number":
        data["foreign_invoice_number"] = text
        state["step"] = "car_name"
        await update.message.reply_text("Введите наименование авто / марка, модель:")
        return

    if step == "car_name":
        data["car_name"] = text
        state["step"] = "identifier_type"
        keyboard = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton("VIN", callback_data="idtype:VIN"),
                    InlineKeyboardButton("Chassis", callback_data="idtype:Chassis"),
                ]
            ]
        )
        await update.message.reply_text("Выберите тип идентификатора:", reply_markup=keyboard)
        return

    if step == "identifier_value":
        data["identifier_value"] = text
        state["step"] = "amount"
        await update.message.reply_text("Введите сумму счета:")
        return

    if step == "amount":
        data["amount"] = text.replace(" ", "").replace(",", ".")
        data["invoice_number"] = get_next_invoice_number()
        data["invoice_date"] = datetime.now().strftime("%d.%m.%Y")

        preview = (
            "Проверьте данные:\n\n"
            f"ФИО: {data['buyer_full_name']}\n"
            f"Счет №: {data['invoice_number']} от {data['invoice_date']}\n"
            f"Договор: {data['contract_number']} от {data['contract_date']}\n"
            f"Инвойс: {data['foreign_invoice_number']}\n"
            f"Авто: {data['car_name']}\n"
            f"{data['identifier_type']}: {data['identifier_value']}\n"
            f"Сумма: {format_amount(data['amount'])} руб."
        )

        keyboard = InlineKeyboardMarkup(
            [
                [InlineKeyboardButton("✅ Сформировать счет", callback_data="make_invoice")],
                [InlineKeyboardButton("❌ Отмена", callback_data="cancel")],
            ]
        )
        await update.message.reply_text(preview, reply_markup=keyboard)
        return


async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = update.effective_user.id
    state = USER_STATE.get(user_id)

    if query.data.startswith("idtype:"):
        if not state:
            return
        identifier_type = query.data.split(":", 1)[1]
        state["data"]["identifier_type"] = identifier_type
        state["step"] = "identifier_value"
        await query.message.reply_text(f"Введите значение {identifier_type}:")
        return

    if query.data == "cancel":
        reset_user(user_id)
        await query.message.reply_text("Отменено.", reply_markup=MAIN_MENU)
        return

    if query.data == "make_invoice":
        if not state:
            await query.message.reply_text("Нет данных для счета.")
            return

        data = state["data"]
        docx = create_invoice_docx(data)

        await query.message.reply_document(
            document=open(docx, "rb"),
            filename=f"Счет_{data['invoice_number']}.docx",
            caption="Счет сформирован ✅",
        )

        keyboard = InlineKeyboardMarkup(
            [
                [InlineKeyboardButton("📑 Сформировать договор", callback_data="make_contract")],
                [InlineKeyboardButton("📎 Сформировать приложение", callback_data="make_appendix")],
            ]
        )
        await query.message.reply_text("Что формируем дальше?", reply_markup=keyboard)
        return

    if query.data == "make_contract":
        if not state:
            return
        data = state["data"]
        docx = create_contract_docx(data)
        await query.message.reply_document(
            document=open(docx, "rb"),
            filename=f"Договор_{data['contract_number']}.docx",
            caption="Договор сформирован ✅",
        )
        return

    if query.data == "make_appendix":
        if not state:
            return
        data = state["data"]
        docx = create_appendix_docx(data)
        await query.message.reply_document(
            document=open(docx, "rb"),
            filename=f"Приложение_{data['contract_number']}.docx",
            caption="Приложение сформировано ✅",
        )
        return


tg_app.add_handler(CommandHandler("start", start))
tg_app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
tg_app.add_handler(CallbackQueryHandler(handle_callback))


@app.route(f"/webhook/{WEBHOOK_SECRET}", methods=["GET", "POST"])
def webhook():
    if request.method == "GET":
        return "Webhook endpoint is alive", 200

    update = Update.de_json(request.get_json(force=True), tg_app.bot)

    import asyncio
    asyncio.run(tg_app.initialize())
    asyncio.run(tg_app.process_update(update))

    return "ok", 200


@app.route("/")
def home():
    return "ASIA MX Docs Bot is running ✅"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

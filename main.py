import os
import json
import tempfile
from datetime import datetime

import requests
import qrcode
import gspread

from flask import Flask, request
from google.oauth2.service_account import Credentials
from docx import Document
from docx.shared import Inches
from num2words import num2words


# =========================
# НАСТРОЙКИ
# =========================

BOT_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")

WEBHOOK_SECRET = "asia_mx_secret"

if not BOT_TOKEN:
    raise RuntimeError("Не задан BOT_TOKEN в Render Environment")

app = Flask(__name__)
USER_STATE = {}


# =========================
# РЕКВИЗИТЫ ASIA MX
# =========================

AGENT_NAME = "ЭЙЖА ЭМ ЭКС ТРЕЙД ЛИМИТЕД (ASIA MX TRADE LIMITED)"
AGENT_SHORT = "ЭЙЖА ЭМ ЭКС ТРЕЙД ЛИМИТЕД"

AGENT_INN = "9909722497"
AGENT_KPP = "270087001"

AGENT_ACCOUNT_QR = "40807810350710000002"
AGENT_BANK = "ДАЛЬНЕВОСТОЧНЫЙ БАНК ПАО СБЕРБАНК"
AGENT_BIK = "040813608"
AGENT_CORR_QR = "30101810600000000608"


# =========================
# TELEGRAM
# =========================

def tg(method, payload=None, files=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/{method}"
    r = requests.post(url, data=payload, files=files, timeout=60)
    print(method, r.status_code, r.text[:500])
    return r


def send_message(chat_id, text, keyboard=None):
    payload = {
        "chat_id": chat_id,
        "text": text,
    }

    if keyboard:
        payload["reply_markup"] = json.dumps(keyboard, ensure_ascii=False)

    return tg("sendMessage", payload)


def send_document(chat_id, file_path, filename, caption=None):
    with open(file_path, "rb") as f:
        files = {"document": (filename, f)}
        payload = {"chat_id": chat_id}

        if caption:
            payload["caption"] = caption

        return tg("sendDocument", payload, files)


def main_menu():
    return {
        "keyboard": [
            [{"text": "📄 Выставить счет"}],
            [{"text": "📑 Договор"}, {"text": "📎 Приложение"}],
            [{"text": "❌ Отмена"}],
        ],
        "resize_keyboard": True,
    }


def identifier_keyboard():
    return {
        "inline_keyboard": [
            [
                {"text": "VIN", "callback_data": "identifier:VIN"},
                {"text": "Chassis", "callback_data": "identifier:Chassis"},
            ]
        ]
    }


def after_invoice_keyboard():
    return {
        "keyboard": [
            [{"text": "📑 Сформировать договор"}],
            [{"text": "📎 Сформировать приложение"}],
            [{"text": "🏠 В меню"}],
        ],
        "resize_keyboard": True,
    }


# =========================
# GOOGLE SHEETS
# =========================

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


def get_next_invoice_number(data):
    sheet = get_sheet()

    if sheet is None:
        return datetime.now().strftime("%H%M%S")

    try:
        ws = sheet.worksheet("СЧЕТА")
    except Exception:
        ws = sheet.add_worksheet(title="СЧЕТА", rows=1000, cols=20)
        ws.append_row([
            "invoice_number",
            "created_at",
            "buyer_full_name",
            "contract_number",
            "contract_date",
            "foreign_invoice_number",
            "car_name",
            "identifier_type",
            "identifier_value",
            "amount",
        ])

    values = ws.get_all_values()
    last_number = 14252

    for row in values[1:]:
        if row and str(row[0]).isdigit():
            last_number = max(last_number, int(row[0]))

    new_number = last_number + 1

    ws.append_row([
        str(new_number),
        datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
        data.get("buyer_full_name", ""),
        data.get("contract_number", ""),
        data.get("contract_date", ""),
        data.get("foreign_invoice_number", ""),
        data.get("car_name", ""),
        data.get("identifier_type", ""),
        data.get("identifier_value", ""),
        data.get("amount", ""),
    ])

    return str(new_number)


# =========================
# ДЕНЬГИ / QR
# =========================

def normalize_amount(amount):
    return float(str(amount).replace(" ", "").replace(",", "."))


def format_amount(amount):
    value = normalize_amount(amount)
    return f"{value:,.2f}".replace(",", " ")


def amount_words(amount):
    value = int(normalize_amount(amount))
    text = num2words(value, lang="ru")
    return f"{text.capitalize()} рублей 00 копеек"


def amount_kopecks(amount):
    return str(int(round(normalize_amount(amount) * 100)))


def create_qr(data):
    purpose = (
        f"Оплата по счету №{data['invoice_number']} "
        f"по договору {data['contract_number']} от {data['contract_date']}"
    )

    qr_text = (
        f"ST00012|"
        f"Name={AGENT_SHORT}|"
        f"PersonalAcc={AGENT_ACCOUNT_QR}|"
        f"BankName={AGENT_BANK}|"
        f"BIC={AGENT_BIK}|"
        f"CorrespAcc={AGENT_CORR_QR}|"
        f"PayeeINN={AGENT_INN}|"
        f"KPP={AGENT_KPP}|"
        f"Purpose={purpose}|"
        f"Sum={amount_kopecks(data['amount'])}"
    )

    img = qrcode.make(qr_text)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    img.save(tmp.name)

    return tmp.name


# =========================
# ШАБЛОНЫ WORD
# =========================

def replace_placeholders(doc, mapping):
    def replace_in_paragraph(paragraph):
        if not paragraph.runs:
            return

        full_text = "".join(run.text for run in paragraph.runs)
        new_text = full_text

        for key, value in mapping.items():
            new_text = new_text.replace(key, str(value))

        if new_text != full_text:
            for run in paragraph.runs:
                run.text = ""
            paragraph.runs[0].text = new_text

    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph)


def insert_qr_code(doc, qr_path):
    for paragraph in doc.paragraphs:
        if "{{QR_CODE}}" in paragraph.text:
            paragraph.text = ""
            run = paragraph.add_run()
            run.add_picture(qr_path, width=Inches(1.7))
            return

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if "{{QR_CODE}}" in paragraph.text:
                        paragraph.text = ""
                        run = paragraph.add_run()
                        run.add_picture(qr_path, width=Inches(1.7))
                        return


def create_invoice_docx(data):
    doc = Document("templates/invoice_template.docx")

    mapping = {
        "{{INVOICE_NUMBER}}": data["invoice_number"],
        "{{INVOICE_DATE}}": data["invoice_date"],
        "{{BUYER_FULL_NAME}}": data["buyer_full_name"],
        "{{CONTRACT_NUMBER}}": data["contract_number"],
        "{{CONTRACT_DATE}}": data["contract_date"],
        "{{FOREIGN_INVOICE_NUMBER}}": data["foreign_invoice_number"],
        "{{CAR_NAME}}": data["car_name"],
        "{{IDENTIFIER_TYPE}}": data["identifier_type"],
        "{{IDENTIFIER_VALUE}}": data["identifier_value"],
        "{{AMOUNT}}": format_amount(data["amount"]),
        "{{AMOUNT_WORDS}}": amount_words(data["amount"]),
    }

    replace_placeholders(doc, mapping)

    qr_path = create_qr(data)
    insert_qr_code(doc, qr_path)

    path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    doc.save(path)
    return path


def create_contract_docx(data):
    doc = Document("templates/contract_template.docx")

    mapping = {
        "{{CONTRACT_NUMBER}}": data["contract_number"],
        "{{CONTRACT_DATE}}": data["contract_date"],
        "{{BUYER_FULL_NAME}}": data["buyer_full_name"],
        "{{BUYER_BIRTH_DATE}}": data.get("buyer_birth_date", ""),
        "{{BUYER_PASSPORT_SERIES}}": data.get("buyer_passport_series", ""),
        "{{BUYER_PASSPORT_NUMBER}}": data.get("buyer_passport_number", ""),
        "{{BUYER_PASSPORT_ISSUED_BY}}": data.get("buyer_passport_issued_by", ""),
        "{{BUYER_PASSPORT_ISSUE_DATE}}": data.get("buyer_passport_issue_date", ""),
        "{{BUYER_PASSPORT_CODE}}": data.get("buyer_passport_code", ""),
        "{{BUYER_INN}}": data.get("buyer_inn", ""),
        "{{BUYER_REG_ADDRESS}}": data.get("buyer_reg_address", ""),
    }

    replace_placeholders(doc, mapping)

    path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    doc.save(path)
    return path


def create_appendix_docx(data):
    doc = Document("templates/appendix_template.docx")

    amount_after_fee = normalize_amount(data["amount"]) * 0.99

    mapping = {
        "{{CONTRACT_NUMBER}}": data["contract_number"],
        "{{CONTRACT_DATE}}": data["contract_date"],
        "{{ORDER_DATE}}": data["invoice_date"],
        "{{BUYER_FULL_NAME}}": data["buyer_full_name"],
        "{{AMOUNT}}": format_amount(data["amount"]),
        "{{AMOUNT_WORDS}}": amount_words(data["amount"]),
        "{{AMOUNT_AFTER_FEE}}": format_amount(amount_after_fee),
        "{{CONTRACTOR}}": data.get("contractor", ""),
        "{{CONTRACTOR_BANK}}": data.get("contractor_bank", ""),
    }

    replace_placeholders(doc, mapping)

    path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    doc.save(path)
    return path


# =========================
# ЛОГИКА БОТА
# =========================

def reset(chat_id):
    USER_STATE.pop(str(chat_id), None)


def start_invoice(chat_id):
    USER_STATE[str(chat_id)] = {
        "mode": "invoice",
        "step": "buyer_full_name",
        "data": {},
    }

    send_message(chat_id, "Введите ФИО покупателя:")


def process_invoice_step(chat_id, text):
    state = USER_STATE.get(str(chat_id))

    if not state:
        send_message(chat_id, "Выберите действие:", main_menu())
        return

    step = state["step"]
    data = state["data"]

    if step == "buyer_full_name":
        data["buyer_full_name"] = text
        state["step"] = "buyer_birth_date"
        send_message(chat_id, "Введите дату рождения:")
        return

    if step == "buyer_birth_date":
        data["buyer_birth_date"] = text
        state["step"] = "buyer_passport_series"
        send_message(chat_id, "Введите серию паспорта:")
        return

    if step == "buyer_passport_series":
        data["buyer_passport_series"] = text
        state["step"] = "buyer_passport_number"
        send_message(chat_id, "Введите номер паспорта:")
        return

    if step == "buyer_passport_number":
        data["buyer_passport_number"] = text
        state["step"] = "buyer_passport_issued_by"
        send_message(chat_id, "Введите кем выдан паспорт:")
        return

    if step == "buyer_passport_issued_by":
        data["buyer_passport_issued_by"] = text
        state["step"] = "buyer_passport_issue_date"
        send_message(chat_id, "Введите дату выдачи паспорта:")
        return

    if step == "buyer_passport_issue_date":
        data["buyer_passport_issue_date"] = text
        state["step"] = "buyer_passport_code"
        send_message(chat_id, "Введите код подразделения:")
        return

    if step == "buyer_passport_code":
        data["buyer_passport_code"] = text
        state["step"] = "buyer_inn"
        send_message(chat_id, "Введите ИНН физлица:")
        return

    if step == "buyer_inn":
        data["buyer_inn"] = text
        state["step"] = "buyer_reg_address"
        send_message(chat_id, "Введите адрес регистрации:")
        return

    if step == "buyer_reg_address":
        data["buyer_reg_address"] = text
        state["step"] = "contract_number"
        send_message(chat_id, "Введите номер договора:")
        return

    if step == "contract_number":
        data["contract_number"] = text
        state["step"] = "contract_date"
        send_message(chat_id, "Введите дату договора:")
        return

    if step == "contract_date":
        data["contract_date"] = text
        state["step"] = "foreign_invoice_number"
        send_message(chat_id, "Введите номер инвойса:")
        return

    if step == "foreign_invoice_number":
        data["foreign_invoice_number"] = text
        state["step"] = "car_name"
        send_message(chat_id, "Введите наименование авто / марка, модель:")
        return

    if step == "car_name":
        data["car_name"] = text
        state["step"] = "identifier_type"
        send_message(chat_id, "Выберите тип идентификатора:", identifier_keyboard())
        return

    if step == "identifier_type":
        if text not in ["VIN", "Chassis"]:
            send_message(chat_id, "Выберите VIN или Chassis:", identifier_keyboard())
            return

        data["identifier_type"] = text
        state["step"] = "identifier_value"
        send_message(chat_id, f"Введите значение {text}:")
        return

    if step == "identifier_value":
        data["identifier_value"] = text
        state["step"] = "amount"
        send_message(chat_id, "Введите сумму счета:")
        return

    if step == "amount":
        data["amount"] = text.replace(" ", "").replace(",", ".")
        state["step"] = "contractor"
        send_message(chat_id, "Введите контрагента для приложения:")
        return

    if step == "contractor":
        data["contractor"] = text
        state["step"] = "contractor_bank"
        send_message(chat_id, "Введите банковские реквизиты контрагента:")
        return

    if step == "contractor_bank":
        data["contractor_bank"] = text

        data["invoice_date"] = datetime.now().strftime("%d.%m.%Y")
        data["invoice_number"] = get_next_invoice_number(data)

        preview = (
            "Проверьте данные:\n\n"
            f"ФИО: {data['buyer_full_name']}\n"
            f"Счет № {data['invoice_number']} от {data['invoice_date']}\n"
            f"Договор: {data['contract_number']} от {data['contract_date']}\n"
            f"Инвойс: {data['foreign_invoice_number']}\n"
            f"Авто: {data['car_name']}\n"
            f"{data['identifier_type']}: {data['identifier_value']}\n"
            f"Сумма: {format_amount(data['amount'])} руб.\n\n"
            f"Формирую счет..."
        )

        send_message(chat_id, preview)

        invoice_path = create_invoice_docx(data)

        send_document(
            chat_id,
            invoice_path,
            f"Счет_{data['invoice_number']}.docx",
            "Счет сформирован ✅",
        )

        send_message(chat_id, "Что формируем дальше?", after_invoice_keyboard())

        state["step"] = "done"
        return

    if step == "done":
        send_message(chat_id, "Документы уже сформированы. Выберите действие:", after_invoice_keyboard())
        return


def create_contract_from_last(chat_id):
    state = USER_STATE.get(str(chat_id))

    if not state or not state.get("data"):
        send_message(chat_id, "Сначала сформируйте счет, чтобы договор подтянул данные.")
        return

    data = state["data"]
    path = create_contract_docx(data)

    send_document(
        chat_id,
        path,
        f"Договор_{data['contract_number']}.docx",
        "Договор сформирован ✅",
    )


def create_appendix_from_last(chat_id):
    state = USER_STATE.get(str(chat_id))

    if not state or not state.get("data"):
        send_message(chat_id, "Сначала сформируйте счет, чтобы приложение подтянуло данные.")
        return

    data = state["data"]
    path = create_appendix_docx(data)

    send_document(
        chat_id,
        path,
        f"Приложение_{data['contract_number']}.docx",
        "Приложение сформировано ✅",
    )


def handle_message(message):
    chat_id = message["chat"]["id"]
    text = message.get("text", "").strip()

    if not text:
        send_message(chat_id, "Я пока работаю только с текстом.")
        return

    if text == "/start":
        reset(chat_id)
        send_message(
            chat_id,
            "Бот документов ASIA MX запущен ✅\n\nВыберите действие:",
            main_menu(),
        )
        return

    if text in ["🏠 В меню", "❌ Отмена"]:
        reset(chat_id)
        send_message(chat_id, "Выберите действие:", main_menu())
        return

    if text == "📄 Выставить счет":
        start_invoice(chat_id)
        return

    if text in ["📑 Договор", "📑 Сформировать договор"]:
        create_contract_from_last(chat_id)
        return

    if text in ["📎 Приложение", "📎 Сформировать приложение"]:
        create_appendix_from_last(chat_id)
        return

    state = USER_STATE.get(str(chat_id))

    if state and state.get("mode") == "invoice":
        process_invoice_step(chat_id, text)
        return

    send_message(chat_id, "Выберите действие:", main_menu())


# =========================
# FLASK ROUTES
# =========================

@app.route("/")
def home():
    return "ASIA MX Docs Bot is running ✅", 200


@app.route("/test")
def test():
    return "TEST OK", 200


@app.route(f"/webhook/{WEBHOOK_SECRET}", methods=["GET", "POST"])
def webhook():
    if request.method == "GET":
        return "Webhook endpoint is alive", 200

    update = request.get_json(force=True)
    print(update)

    try:
        if "message" in update:
            handle_message(update["message"])

        if "callback_query" in update:
            callback = update["callback_query"]
            chat_id = callback["message"]["chat"]["id"]
            data = callback["data"]

            if data.startswith("identifier:"):
                value = data.split(":", 1)[1]
                state = USER_STATE.get(str(chat_id))

                if state:
                    state["data"]["identifier_type"] = value
                    state["step"] = "identifier_value"
                    send_message(chat_id, f"Введите значение {value}:")

    except Exception as e:
        print("BOT ERROR:", repr(e))

        try:
            if "message" in update:
                chat_id = update["message"]["chat"]["id"]
            else:
                chat_id = update["callback_query"]["message"]["chat"]["id"]

            send_message(chat_id, f"Ошибка: {e}")
        except Exception:
            pass

    return "ok", 200

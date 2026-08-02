import os  
import logging
from datetime import datetime
import asyncio
import json
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.utils.keyboard import ReplyKeyboardBuilder, InlineKeyboardBuilder
from aiohttp import web, ClientSession
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.page import PageMargins
from aiogram.types import FSInputFile
import gspread
from google.oauth2.service_account import Credentials
from PIL import Image, ImageDraw, ImageFont
import qrcode
logging.basicConfig(level=logging.INFO)

API_TOKEN = os.getenv("BOT_TOKEN")
if not API_TOKEN:
    raise ValueError("Xatolik: BOT_TOKEN topilmadi! Render sozlamalarini tekshiring.")

bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

EXCEL_FILE = "students.xlsx"

def generate_certificate(full_name: str, user_id: int) -> str:
    template = Image.open("certificate_template.png").convert("RGB")
    draw = ImageDraw.Draw(template)
    font = ImageFont.truetype(
        "/usr/share/fonts/truetype/liberation/LiberationSerif-BoldItalic.ttf", 42
    )
    text = full_name.upper()
    bbox = draw.textbbox((0, 0), text, font=font)
    w = bbox[2] - bbox[0]
    x = (template.width - w) / 2
    y = 340  # 470 dan 520 ga tushirildi — chiziq ustiga to'g'ri kelishi uchun
    draw.text((x, y), text, font=font, fill=(255, 255, 255))

    date_text = datetime.now().strftime("%d.%m.%Y")
    draw.text(
        (100, template.height - 60), f"Sana: {date_text}",
        font=ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 24),
        fill=(180, 180, 180)
    )

    qr_link = f"https://t.me/AngrenAkademiyaBot?start=cabinet_{user_id}"
    qr_link = f"https://t.me/SIZNING_BOT_USERNAME?start=cabinet_{user_id}"
    qr_img = qrcode.make(qr_link).resize((110, 110))
    template.paste(qr_img, (template.width - 200, 330))

    out_path = f"cert_{user_id}.png"
    template.save(out_path)
    return out_path
 
DIPLOMA_TIERS = [
    (100, "mutlaq_g'olib.png", "MUTLAQ G'OLIB"),
    (95, "a'lo_darajali.png", "A'LO DARAJALI O'QUVCHI"),
    (90, "faol_va_iqtidorli.png", "FAOL VA IQTIDORLI O'QUVCHI"),
    (85, "iqtidorli.png", "IQTIDORLI O'QUVCHI"),
    (80, "bilimdon.png",  "BILIMDON O'QUVCHI"),
    (70, "bilimga_intiluvchi.png", "BILIMGA INTILUVCHI O'QUVCHI"),
]


def get_diploma_tier(percent: int):
    for min_percent, filename, title in DIPLOMA_TIERS:
        if percent >= min_percent:
            return filename, title
    return None, None


def generate_diploma(full_name: str, subject: str, percent: int, user_id: int) -> str:
    filename, title = get_diploma_tier(percent)
    if not filename:
        return None

    template = Image.open(filename).convert("RGB")
    draw = ImageDraw.Draw(template)

    name_font = ImageFont.truetype(
        "/usr/share/fonts/truetype/liberation/LiberationSerif-BoldItalic.ttf", 28
    )
    subject_font = ImageFont.truetype(
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 20
    )

    name_text = full_name.upper()
    bbox = draw.textbbox((0, 0), name_text, font=name_font)
    w = bbox[2] - bbox[0]
    x = (template.width - w) / 2
    y = int(template.height * 0.44)
    draw.text((x, y), name_text, font=name_font, fill=(20, 20, 90))

    bbox2 = draw.textbbox((0, 0), subject, font=subject_font)
    w2 = bbox2[2] - bbox2[0]
    x2 = (template.width - w2) / 2
    y2 = int(template.height * 0.53)
    draw.text((x2, y2), subject, font=subject_font, fill=(20, 20, 90))

    out_path = f"diploma_{user_id}.png"
    template.save(out_path)
    return out_path  
def save_to_excel(data, user_id):
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "O'quvchilar"
        ws.append(["№", "Sana", "Ism Familiya", "Tel Raqam", "Ota-ona Tel", "Maktab", "Sinf", "Filial", "Smena", "Kurslar", "ID"])
        wb.save(EXCEL_FILE)

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    sana = datetime.now().strftime("%d.%m.%Y %H:%M")

    courses_list = data.get("selected_courses", [])
    courses_string = "\n".join(f"• {c.replace(chr(10), ' ')}" for c in courses_list)

    tartib_raqam = ws.max_row

    ws.append([
        tartib_raqam,
        sana,
        data.get("name"),
        data.get("phone"),
        data.get("parent_phone"),
        data.get("school"),
        data.get("grade"),
        data.get("filial"),
        data.get("time_pref"),
        courses_string,
        user_id
    ])

    try:
        col_widths = {1: 5, 2: 16, 3: 22, 4: 14, 5: 14, 6: 10, 7: 6, 8: 12, 9: 12, 10: 15, 11: 15}
        for col_num, width in col_widths.items():
            col_letter = openpyxl.utils.get_column_letter(col_num)
            ws.column_dimensions[col_letter].width = width

        for row in ws.iter_rows(min_row=1):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")

        for row in ws.iter_rows(min_row=2):
            max_lines = 1
            for cell in row:
                if cell.value:
                    lines_count = str(cell.value).count("\n") + 1
                    if lines_count > max_lines:
                        max_lines = lines_count
            ws.row_dimensions[row[0].row].height = max(max_lines * 15, 20)

        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75)

    except Exception as e:
        logging.warning(f"Excel formatlashda xato (malumot yozildi): {e}")

    wb.save(EXCEL_FILE)


def _write_to_google_sheets_sync(data, user_id):
    now = datetime.now()
    sana = now.strftime("%d.%m.%Y %H:%M")
    bugun = now.strftime("%d.%m.%Y")

    courses_list = data.get("selected_courses", [])
    courses_string = ", ".join([c.replace('\n', ' ') for c in courses_list])

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds_json = os.getenv("GOOGLE_CREDS")

    if creds_json:
        creds_data = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_data, scopes=scopes)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)

    client = gspread.authorize(creds)

    sheet_url = os.getenv("GOOGLE_SHEET_URL")
    spreadsheet_id = os.getenv("SPREADSHEET_ID")

    if sheet_url:
        spreadsheet = client.open_by_url(sheet_url)
    elif spreadsheet_id:
        spreadsheet = client.open_by_key(spreadsheet_id)
    else:
        spreadsheet = client.open_by_url(
            "https://docs.google.com/spreadsheets/d/1aXoL-TeP0Oh62u1kfgPyzyRsNjOdqGkovJmFutYlUn0/edit"
        )

    try:
        sheet = spreadsheet.worksheet(bugun)
    except Exception:
        sheet = spreadsheet.add_worksheet(title=bugun, rows=1000, cols=11)
        sheet.append_row([
            "№", "Sana", "Ism Familiya", "Tel Raqam",
            "Ota-ona Tel", "Maktab", "Sinf", "Filial", "Smena", "Kurslar", "ID"
        ])

    all_rows = sheet.get_all_values()
    tartib_raqam = len(all_rows)

    sheet.append_row([
        tartib_raqam,
        sana,
        data.get("name"),
        data.get("phone"),
        data.get("parent_phone"),
        data.get("school"),
        data.get("grade"),
        data.get("filial"),
        data.get("time_pref"),
        courses_string,
        user_id
    ])


async def save_to_google_sheets(data, user_id):
    try:
        await asyncio.to_thread(_write_to_google_sheets_sync, data, user_id)
        logging.info("Ma'lumotlar Google Sheets'ga muvaffaqiyatli yozildi!")
    except Exception as e:
        logging.exception("Google Sheets yozishda xato:")
        admin_id = os.getenv("ADMIN_ID")
        if admin_id:
            try:
                await bot.send_message(
                    int(admin_id),
                    f"Google Sheets xato: {type(e).__name__}: {e}"
                )
            except Exception:
                pass


class Registration(StatesGroup):
    name = State()
    phone = State()
    parent_phone = State()
    school = State()
    grade = State()
    filial = State()
    subjects = State()
    time_pref = State()


AVAILABLE_FILIALS = ["Angren", "Ohangaron"]
AVAILABLE_TIMES = ["Ertalabki", "Kunduzgi", "Kechki"]
AVAILABLE_SUBJECTS = [
    "Matematika - Milliy va xalqaro sertifikat",
    "Matematika - majburiy blok uchun",
    "Ingliz tili - IELTS",
    "Tibbiyot - shifokorlik kasblari uchun\nKimyo - Milliy va xalqaro sertifikat",
    "Prezident maktablariga tayyorlov",
    "Al-Xorazmiy maktablariga tayyorlov",
    "Tibbiyot - shifokorlik kasbini tanlaganlar uchun\nBiologiya - Milliy va xalqaro sertifikat",
    "Tarix - Milliy sertifikat",
    "Tarix - Majburiy blok uchun",
    "Huquq - Milliy sertifikat",
    "IT - Milliy va xalqaro sertifikat",
    "Ona tili va adabiyoti - Milliy sertifikat",
    "Ona tili va adabiyoti - Majburiy blok uchun",
    "Maktabga tayyorlov. Pochemuchka"
]
class TestQuiz(StatesGroup):
    subject = State()
    grade = State()
    question = State()

MATH_QUESTIONS = {
    1: [
        {"q": "2+3=?", "options": ["4","5","6","7"], "correct": 1},
        {"q": "7-2=?", "options": ["4","5","6","3"], "correct": 1},
        {"q": "1+1=?", "options": ["1","2","3","4"], "correct": 1},
        {"q": "9-4=?", "options": ["3","4","5","6"], "correct": 2},
        {"q": "5+4=?", "options": ["7","8","9","10"], "correct": 2},
        {"q": "10-6=?", "options": ["3","4","5","6"], "correct": 1},
        {"q": "6+3=?", "options": ["8","9","10","7"], "correct": 1},
        {"q": "8-5=?", "options": ["2","3","4","5"], "correct": 1},
        {"q": "4+4+1=?", "options": ["8","9","10","7"], "correct": 1},
        {"q": "10-3-2=?", "options": ["4","5","6","3"], "correct": 1},
    ],
    2: [
        {"q": "12+7=?", "options": ["18","19","20","17"], "correct": 1},
        {"q": "15-8=?", "options": ["6","7","8","9"], "correct": 1},
        {"q": "6×2=?", "options": ["10","12","14","8"], "correct": 1},
        {"q": "20-9=?", "options": ["10","11","12","9"], "correct": 1},
        {"q": "9+9=?", "options": ["16","17","18","19"], "correct": 2},
        {"q": "4×5=?", "options": ["18","20","22","16"], "correct": 1},
        {"q": "36-18=?", "options": ["16","17","18","19"], "correct": 2},
        {"q": "7×3=?", "options": ["19","20","21","22"], "correct": 2},
        {"q": "45+27=?", "options": ["70","71","72","73"], "correct": 2},
        {"q": "60-24=?", "options": ["34","35","36","37"], "correct": 2},
    ],
    3: [
        {"q": "8×7=?", "options": ["54","55","56","57"], "correct": 2},
        {"q": "63÷9=?", "options": ["6","7","8","9"], "correct": 1},
        {"q": "100-45=?", "options": ["54","55","56","57"], "correct": 1},
        {"q": "9×6=?", "options": ["52","53","54","55"], "correct": 2},
        {"q": "72÷8=?", "options": ["7","8","9","10"], "correct": 2},
        {"q": "125+237=?", "options": ["360","361","362","363"], "correct": 2},
        {"q": "84÷7=?", "options": ["11","12","13","14"], "correct": 1},
        {"q": "15×3=?", "options": ["43","44","45","46"], "correct": 2},
        {"q": "500-268=?", "options": ["230","231","232","233"], "correct": 2},
        {"q": "96÷8÷2=?", "options": ["5","6","7","8"], "correct": 1},
    ],
    4: [
        {"q": "234+567=?", "options": ["799","800","801","802"], "correct": 2},
        {"q": "12×12=?", "options": ["142","143","144","145"], "correct": 2},
        {"q": "900÷30=?", "options": ["28","29","30","31"], "correct": 2},
        {"q": "1000-456=?", "options": ["543","544","545","546"], "correct": 1},
        {"q": "15×20=?", "options": ["280","290","300","310"], "correct": 2},
        {"q": "3/4 + 1/4 = ?", "options": ["1/2","3/4","1","5/4"], "correct": 2},
        {"q": "6.5+3.2=?", "options": ["9.5","9.6","9.7","9.8"], "correct": 2},
        {"q": "144÷12=?", "options": ["11","12","13","14"], "correct": 1},
        {"q": "Tomoni 6 sm kvadrat perimetri?", "options": ["18","20","22","24"], "correct": 3},
        {"q": "2500÷25=?", "options": ["90","95","100","105"], "correct": 2},
    ],
    5: [
        {"q": "3/5 + 1/5 = ?", "options": ["2/5","4/5","3/10","1"], "correct": 1},
        {"q": "25% dan 200 necha?", "options": ["40","45","50","55"], "correct": 2},
        {"q": "7² = ?", "options": ["14","42","49","56"], "correct": 2},
        {"q": "0.75 ni foizga aylantiring", "options": ["65%","70%","75%","80%"], "correct": 2},
        {"q": "18×15=?", "options": ["260","270","280","290"], "correct": 1},
        {"q": "To'g'ri to'rtburchak yuzi: 8×5=?", "options": ["35","38","40","42"], "correct": 2},
        {"q": "-5+8=?", "options": ["2","3","4","13"], "correct": 1},
        {"q": "2/3 ning 1/2 qismi?", "options": ["1/6","1/3","2/6","3/5"], "correct": 1},
        {"q": "144 ning kvadrat ildizi?", "options": ["11","12","13","14"], "correct": 1},
        {"q": "3x=27, x=?", "options": ["7","8","9","10"], "correct": 2},
    ],
    6: [
        {"q": "-8+(-5)=?", "options": ["-13","-3","3","13"], "correct": 0},
        {"q": "2x+5=15, x=?", "options": ["4","5","6","7"], "correct": 1},
        {"q": "40% dan 350 necha?", "options": ["130","135","140","145"], "correct": 2},
        {"q": "(-3)×(-4)=?", "options": ["-12","-7","7","12"], "correct": 3},
        {"q": "15/20 ni qisqartiring", "options": ["3/4","1/2","4/5","2/3"], "correct": 0},
        {"q": "Radiusi 7 bo'lgan aylananing diametri?", "options": ["10","12","14","16"], "correct": 2},
        {"q": "5²-3²=?", "options": ["14","16","18","20"], "correct": 1},
        {"q": "3:4 = 15:x, x=?", "options": ["18","19","20","21"], "correct": 2},
        {"q": "-12÷4=?", "options": ["-4","-3","3","4"], "correct": 1},
        {"q": "2(x+3)=16, x=?", "options": ["4","5","6","7"], "correct": 1},
    ],
    7: [
        {"q": "3x-7=14, x=?", "options": ["5","6","7","8"], "correct": 2},
        {"q": "(x+2)(x-2)=?", "options": ["x²-4","x²+4","x²-2x","2x"], "correct": 0},
        {"q": "Ikki burchak yig'indisi 90°, biri 35°. Ikkinchisi?", "options": ["45","55","60","65"], "correct": 1},
        {"q": "2⁵=?", "options": ["16","32","64","10"], "correct": 1},
        {"q": "y=2x+1, x=3 bo'lsa y=?", "options": ["5","6","7","8"], "correct": 2},
        {"q": "√81=?", "options": ["7","8","9","10"], "correct": 2},
        {"q": "Uchburchak ichki burchaklar yig'indisi?", "options": ["90","180","270","360"], "correct": 1},
        {"q": "-3x=12, x=?", "options": ["-4","-3","3","4"], "correct": 0},
        {"q": "5(2x-1)=45, x=?", "options": ["4","5","6","7"], "correct": 1},
        {"q": "x/6=4/8, x=?", "options": ["2","3","4","5"], "correct": 1},
    ],
}

TEST_SUBJECTS = {
    "Matematika": MATH_QUESTIONS,
}

def get_main_menu():
    kb = ReplyKeyboardBuilder()
    kb.button(text="📝 Ro'yxatdan o'tish")
    kb.button(text="📈 Bilim darajasini tekshirish")
    kb.button(text="🚪 Davomat (Keldim/Ketdim)")
    kb.adjust(1, 2)
    kb.button(text="👤 Shaxsiy kabinet")
    return kb.as_markup(resize_keyboard=True)


@dp.message(F.text == "/start")
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "✨ Angren Akademiyasi rasmiy botiga xush kelibsiz!\n\n"
        "Kelajak akademiyasida o'z bilimingizni va farzandingiz kamolotini nazorat qiling.",
        reply_markup=get_main_menu()
    )


@dp.message(F.text == "/help")
async def cmd_help(message: types.Message):
    await message.answer("Sizga qanday yordam bera olaman?")


@dp.message(F.text == "/excel")
async def send_excel(message: types.Message):
    if os.path.exists(EXCEL_FILE):
        excel_doc = FSInputFile(EXCEL_FILE)
        await message.answer_document(document=excel_doc, caption="📊 Angren Akademiya o'quvchilar ro'yxati (Excel)")
    else:
        await message.answer("❌ Hozircha ro'yxat bo'sh! Hech kim ro'yxatdan o'tmadi.")


@dp.message(F.text == "📈 Bilim darajasini tekshirish")
async def check_knowledge(message: types.Message, state: FSMContext):
    kb = InlineKeyboardBuilder()
    for subject in TEST_SUBJECTS:
        kb.button(text=subject, callback_data=f"test_subj_{subject}")
    kb.adjust(1)
    await message.answer("📚 Qaysi fandan test topshirmoqchisiz?", reply_markup=kb.as_markup())
    await state.set_state(TestQuiz.subject)


@dp.callback_query(TestQuiz.subject, F.data.startswith("test_subj_"))
async def choose_grade(callback: types.CallbackQuery, state: FSMContext):
    subject = callback.data.replace("test_subj_", "")
    await state.update_data(subject=subject, score=0, q_index=0)

    kb = InlineKeyboardBuilder()
    for grade in range(1, 8):
        kb.button(text=f"{grade}-sinf", callback_data=f"test_grade_{grade}")
    kb.adjust(4, 3)
    await callback.message.edit_text("🎓 Sinfingizni tanlang:", reply_markup=kb.as_markup())
    await state.set_state(TestQuiz.grade)
    await callback.answer()


@dp.callback_query(TestQuiz.grade, F.data.startswith("test_grade_"))
async def start_questions(callback: types.CallbackQuery, state: FSMContext):
    grade = int(callback.data.replace("test_grade_", ""))
    data = await state.get_data()
    subject = data["subject"]
    questions = TEST_SUBJECTS[subject][grade]

    await state.update_data(grade=grade, questions=questions, q_index=0, score=0)
    await state.set_state(TestQuiz.question)
    await callback.answer()
    await send_question(callback.message, state)


async def send_question(message, state: FSMContext):
    data = await state.get_data()
    q_index = data["q_index"]
    questions = data["questions"]

    if q_index >= len(questions):
        await finish_test(message, state)
        return

    question = questions[q_index]
    kb = InlineKeyboardBuilder()
    for idx, option in enumerate(question["options"]):
        kb.button(text=option, callback_data=f"test_ans_{idx}")
    kb.adjust(2)

    await message.answer(
        f"❓ Savol {q_index + 1}/{len(questions)}:\n\n{question['q']}",
        reply_markup=kb.as_markup()
    )


@dp.callback_query(TestQuiz.question, F.data.startswith("test_ans_"))
async def process_answer(callback: types.CallbackQuery, state: FSMContext):
    chosen = int(callback.data.replace("test_ans_", ""))
    data = await state.get_data()
    q_index = data["q_index"]
    questions = data["questions"]
    score = data["score"]

    correct = questions[q_index]["correct"]
    if chosen == correct:
        score += 1
        await callback.answer("✅ To'g'ri!")
    else:
        correct_text = questions[q_index]["options"][correct]
        await callback.answer(f"❌ Xato! To'g'ri javob: {correct_text}", show_alert=True)

    await state.update_data(score=score, q_index=q_index + 1)
    await send_question(callback.message, state)


async def finish_test(message, state: FSMContext):
    data = await state.get_data()
    score = data["score"]
    questions = data["questions"]
    subject = data["subject"]
    grade = data["grade"]
    total = len(questions)
    percent = int((score / total) * 100)

    filename, title = get_diploma_tier(percent)

    result_text = (
        f"🎉 Test yakunlandi!\n\n"
        f"📊 Natija: {score}/{total} ({percent}%)\n"
    )

    if title:
        result_text += f"\n🏅 Siz \"{title}\" nominatsiyasiga munosib bo'ldingiz!"
    else:
        result_text += "\n💪 Yana urinib ko'ring, natijangizni yaxshilay olasiz!"

    await message.answer(result_text, reply_markup=get_main_menu())

    if filename:
        user_data = await state.get_data()
        full_name = message.chat.first_name or "O'quvchi"
        try:
            diploma_path = generate_diploma(full_name, f"{subject} ({grade}-sinf)", percent, message.chat.id)
            if diploma_path:
                await message.answer_photo(
                    photo=FSInputFile(diploma_path),
                    caption=f"🏅 \"{title}\" — Maqtov yorlig'ingiz tayyor!"
                )
                os.remove(diploma_path)
        except Exception:
            logging.exception("Diplom generatsiyasida xato:")

    await state.clear()

@dp.message(F.text == "🚪 Davomat (Keldim/Ketdim)")
async def attendance_menu(message: types.Message):
    kb = InlineKeyboardBuilder()
    kb.button(text="🔔 Keldim", callback_data="attendance_in")
    kb.button(text="🔕 Ketdim", callback_data="attendance_out")
    kb.button(text="💰 Oylik to'lov holati", callback_data="attendance_pay")
    kb.button(text="🚀 Uzoq muddatli imtiyozlar", callback_data="attendance_promo")
    kb.button(text="👤 Shaxsiy kabinet", callback_data="profile")
    kb.adjust(2, 2, 1)
    await message.answer(
        "🚪 Angren Akademiyasi — Davomat va Shaxsiy Balans\n\nKerakli tugmani bosing:",
        reply_markup=kb.as_markup()
    )


@dp.callback_query(F.data.startswith("attendance_"))
async def process_attendance(callback: types.CallbackQuery):
    action = callback.data.split("_")[1]
    current_time = datetime.now().strftime("%H:%M")

    if action == "in":
        await callback.message.answer(
            f"🔔 Angren Akademiya Xabarnomasi\n\nFarzandingiz soat {current_time} da markazimizga eson-omon yetib keldi. 🔬"
        )
    elif action == "out":
        await callback.message.answer(
            f"🔕 Angren Akademiya Xabarnomasi\n\nFarzandingiz soat {current_time} da darsdan chiqdi. Oq yo'l! ☀️"
        )
    elif action == "pay":
        pay_kb = InlineKeyboardBuilder()
        pay_kb.button(text="💳 Plastik (Click/Payme)", callback_data="pay_via_card")
        pay_kb.button(text="💵 Naqd pul (Qo'lda)", callback_data="pay_via_cash")
        pay_kb.adjust(1)
        await callback.message.answer("💰 To'lov usulini tanlang:", reply_markup=pay_kb.as_markup())
    elif action == "promo":
        await callback.message.answer(
            "🚀 \"Angren Akademiya\" Premium Imtiyozlar:\n\n"
            "🥈 3 Oylik: 10% chegirma + sovg'a daftar 🎁\n"
            "🥇 6 Oylik: 15% chegirma + futbolka va kepka 👕\n"
            "👑 1 Yillik: 20% chegirma + darsliklar bepul 📚"
        )
    await callback.answer()


@dp.callback_query(F.data.startswith("pay_via_"))
async def method_payment(callback: types.CallbackQuery):
    method = callback.data.split("_")[2]
    karta_raqam = os.getenv("KARTA_RAQAMI", "8600 0000 0000 0000")
    karta_egasi = os.getenv("KARTA_EGASI", "Angren Akademiya Mas'ul Xodimi")

    if method == "card":
        await callback.message.answer(
            f"💳 Karta raqami: {karta_raqam}\nEga: {karta_egasi}\n\nChekni adminga yuboring."
        )
    else:
        await callback.message.answer("💵 To'lovni administratorga topshiring. Rahmat!")
    await callback.answer()


@dp.message(F.text == "📝 Ro'yxatdan o'tish")
async def start_registration(message: types.Message, state: FSMContext):
    await message.answer("Ism va familiyangizni kiriting:")
    await state.set_state(Registration.name)


@dp.message(Registration.name)
async def process_name(message: types.Message, state: FSMContext):
    await state.update_data(name=message.text)
    await message.answer("📞 O'quvchining telefon raqamini kiriting:")
    await state.set_state(Registration.phone)


@dp.message(Registration.phone)
async def process_phone(message: types.Message, state: FSMContext):
    await state.update_data(phone=message.text)
    await message.answer("👨‍👩‍👦 Ota-onangizning telefon raqamini kiriting:")
    await state.set_state(Registration.parent_phone)


@dp.message(Registration.parent_phone)
async def process_parent_phone(message: types.Message, state: FSMContext):
    await state.update_data(parent_phone=message.text)
    await message.answer("🏫 Nechanchi maktabda o'qiysiz?")
    await state.set_state(Registration.school)


@dp.message(Registration.school)
async def process_school(message: types.Message, state: FSMContext):
    await state.update_data(school=message.text)
    await message.answer("🎓 Nechanchi sinfda o'qiysiz?")
    await state.set_state(Registration.grade)


@dp.message(Registration.grade)
async def process_grade(message: types.Message, state: FSMContext):
    await state.update_data(grade=message.text)
    kb = ReplyKeyboardBuilder()
    for filial in AVAILABLE_FILIALS:
        kb.button(text=filial)
    kb.adjust(2)
    await message.answer("📍 Filialni tanlang:", reply_markup=kb.as_markup(resize_keyboard=True))
    await state.set_state(Registration.filial)


@dp.message(Registration.filial)
async def process_filial(message: types.Message, state: FSMContext):
    if message.text not in AVAILABLE_FILIALS:
        await message.answer("Tugmalardan birini bosing!")
        return
    await state.update_data(filial=message.text, selected_courses=[])
    await show_subjects_keyboard(message, [])

    continue_kb = InlineKeyboardBuilder()
    continue_kb.button(text="✅ Tanlashni tugatdim, davom etish ➡️", callback_data="sub_done")
    await message.answer(
        "👆 Kerakli kurslarni yuqoridan belgilang.\n\n"
        "👇 Barchasini tanlab bo'lgach, shu tugmani bosing:",
        reply_markup=continue_kb.as_markup()
    )

    await state.set_state(Registration.subjects)


async def show_subjects_keyboard(message, selected_courses):
    kb = InlineKeyboardBuilder()
    for idx, subject in enumerate(AVAILABLE_SUBJECTS):
        status = "✅" if subject in selected_courses else ""
        kb.button(text=f"{subject} {status}", callback_data=f"sub_{idx}")
    kb.adjust(1)

    text = "📚 Kurslarni tanlang:\n\n"
    if selected_courses:
        text += "Tanlanganlar:\n" + "\n".join(
            f"- {c.replace(chr(10), ' ')}" for c in selected_courses
        )

    if isinstance(message, types.Message):
        await message.answer(text, reply_markup=kb.as_markup())
    else:
        await message.message.edit_text(text, reply_markup=kb.as_markup())


@dp.callback_query(Registration.subjects, F.data.startswith("sub_"))
async def process_subjects(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    selected_courses = data.get("selected_courses", [])
    action = callback.data.split("_")[1]

    if action == "done":
        if not selected_courses:
            await callback.answer("Kamida bitta fan tanlang!", show_alert=True)
            return
        kb = ReplyKeyboardBuilder()
        for t in AVAILABLE_TIMES:
            kb.button(text=t)
        kb.adjust(3)
        await callback.message.answer(
            "Dars vaqtini tanlang:", reply_markup=kb.as_markup(resize_keyboard=True)
        )
        await state.set_state(Registration.time_pref)
        await callback.answer()
        return

    subject_idx = int(action)
    subject_name = AVAILABLE_SUBJECTS[subject_idx]

    if subject_name in selected_courses:
        selected_courses.remove(subject_name)
    else:
        selected_courses.append(subject_name)

    await state.update_data(selected_courses=selected_courses)
    await show_subjects_keyboard(callback, selected_courses)
    await callback.answer()


def escape_markdown(text):
    if text is None:
        return ""
    text = str(text)
    for ch in ("_", "*", "`", "["):
        text = text.replace(ch, "\\" + ch)
    return text


@dp.message(Registration.time_pref)
async def process_time_pref(message: types.Message, state: FSMContext):
    if message.text not in AVAILABLE_TIMES:
        await message.answer("Smenani tugmalardan tanlang!")
        return
    await state.update_data(time_pref=message.text)
    user_data = await state.get_data()

    try:
        save_to_excel(user_data, message.from_user.id)
    except Exception:
        logging.exception("Excel'ga yozishda xato:")

    asyncio.create_task(save_to_google_sheets(user_data, message.from_user.id))
    admin_id = os.getenv("ADMIN_ID")
    if admin_id:
        admin_text = (
            f"🆕 Yangi o'quvchi!\n"
            f"👤 {user_data.get('name')}\n"
            f"📞 {user_data.get('phone')}\n"
            f"👨‍👩‍👦 Ota-ona: {user_data.get('parent_phone')}\n"
            f"🏫 Maktab: {user_data.get('school')}, Sinf: {user_data.get('grade')}\n"
            f"📍 Filial: {user_data.get('filial')} | Smena: {user_data.get('time_pref')}\n"
            f"🕐 {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        )
        try:
            await bot.send_message(int(admin_id), admin_text)
        except Exception:
            logging.exception("Admin xabarida xato:")
    selected_courses = user_data.get("selected_courses", [])
    courses_output = "📚 Tanlangan kurslar:\n" + "".join(
        f"• {c.replace(chr(10), ' ')}\n" for c in selected_courses
    )

    student_report = (
        f"Muvaffaqiyatli royxatdan o'tdingiz!\n\n"
        f"O'quvchi: {escape_markdown(user_data.get('name'))}\n"
        f"Maktab/Sinf: {escape_markdown(user_data.get('school'))}, {escape_markdown(user_data.get('grade'))}\n"
        f"Filial: {escape_markdown(user_data.get('filial'))} | Smena: {escape_markdown(user_data.get('time_pref'))}\n\n"
        f"{courses_output}\n"
        f"+998 94 041 42 55\n+998 93 101 58 70"
    )

    is_medical = any("Tibbiyot" in course for course in selected_courses)
    if is_medical:
        try:
            photo = FSInputFile("IMG_20260619_235730_628.jpg")
            await message.answer_photo(
                photo=photo,
                caption=student_report,
                parse_mode="Markdown",
                reply_markup=get_main_menu()
            )
        except Exception:
            logging.exception("Rasm yuborishda xato:")
            await message.answer(
                text=student_report,
                parse_mode="Markdown",
                reply_markup=get_main_menu()
            )
    else:
        await message.answer(
            text=student_report,
            parse_mode="Markdown",
            reply_markup=get_main_menu()
        )
    tabrik_kb = InlineKeyboardBuilder()
    tabrik_kb.button(text="👥 Do'stimni taklif qilish", url="https://t.me/share/url?url=https://t.me/AngrenAkademiyaBot&text=Angren+Akademiyaga+qo%27shiling!")
    tabrik_kb.button(text="📍 Angren Akademiya kanali", url="https://t.me/AngrenAkademiya")
    tabrik_kb.adjust(1)
    await message.answer(
        f"🎉 Tabriklaymiz, {escape_markdown(user_data.get('name'))}!\n\n"
        f"🎓 Sizga sertifikat berildi!\n"
        f"Kelib Angren Akademiyadan olib keting!\n\n"
        f"📚 Darslarimizga ishtirok etib voucherlarni qo'lga kiriting!\n\n"
        f"🚀 O'z karyerangiz tomon — biz bilan qadam bosing!\n\n"
        f"🎁 Do'stlaringizni taklif qiling:\n"
        f"👉 1 ta do'st — 30 000 so'mlik voucher yoki sovg'a!\n"
        f"👉 2 ta do'st — 40 000 so'mlik voucher yoki sovg'a!\n"
        f"👉 3 va undan ko'p — 50 000 so'mlik voucher yoki sovg'alarni ham qo'lga kiriting!\n\n"
        f"⚠️ Joylar cheklangan — imkoniyatlarni qo'ldan boy bermang!",
        reply_markup=tabrik_kb.as_markup()
    )
    try:
        cert_path = generate_certificate(user_data.get("name"), message.from_user.id)
        await message.answer_photo(
            photo=FSInputFile(cert_path),
            caption="🎓 Tabriklaymiz! Sizning shaxsiy sertifikatingiz tayyor."
        )
        os.remove(cert_path)
    except Exception:
        logging.exception("Sertifikat generatsiyasida xato:")

    channel_id = os.getenv("CHANNEL_ID")
    if channel_id:
        try:
            await bot.send_message(int(channel_id),
                f"🆕 Yangi o'quvchi!\n"
                f"👤 {user_data.get('name')}\n"
                f"📞 {user_data.get('phone')}\n"
                f"📍 {user_data.get('filial')} | {user_data.get('time_pref')}\n"
                f"🕐 {datetime.now().strftime('%d.%m.%Y %H:%M')}"
            )
        except Exception:
            logging.exception("Kanalga xabar yuborishda xato:")

    await state.clear()
@dp.message(F.text == "👤 Shaxsiy kabinet")
async def shaxsiy_kabinet(message: types.Message):
    await message.answer(
        f"👤 Sizning shaxsiy kabinetingiz\n\n"
        f"💳 Voucher balansiz: 30 000 so'm\n"
        f"⚡️ Darslarni boshlashingiz bilan faollashadi!\n\n"
        f"👥 Taklif qilgan do'stlar: 0 ta\n"
        f"🎓 Sertifikat: berildi ✅\n\n"
        f"🚀 O'z karyerangiz tomon — biz bilan qadam bosing!"
    )
@dp.callback_query(F.data == "profile")
async def profile_handler(callback: types.CallbackQuery):
    await callback.message.answer(
    f"👤 Sizning shaxsiy kabinetingiz\n\n"
    f"💳 Voucher balansiz: 30 000 so'm\n"
    f"⚡️ Darslarni boshlashingiz bilan faollashadi!\n\n"
    f"👥 Taklif qilgan do'stlar: 0 ta\n"
    f"🎓 Sertifikat: berildi ✅\n\n"
    f"🚀 O'z karyerangiz tomon — biz bilan qadam bosing!"
)



@dp.callback_query(F.data == "cert")
async def cert_handler(callback: types.CallbackQuery):
    await callback.message.answer("🎓 Sertifikat yuklash bo‘limi.")


async def handle_health(request):
    return web.Response(text="Angren Akademiya boti faol!")


async def start_web_server():
    app = web.Application()
    app.router.add_get("/", handle_health)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.getenv("PORT", 10000))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()


async def pinger_loop():
    app_url = os.getenv("APP_URL")
    if not app_url:
        return
    await asyncio.sleep(10)
    while True:
        try:
            async with ClientSession() as session:
                async with session.get(app_url) as response:
                    pass
        except Exception:
            pass
        await asyncio.sleep(300)


async def main():
    await start_web_server()
    asyncio.create_task(pinger_loop())
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())

import asyncio
import logging
import openpyxl
import os
import sqlite3
from datetime import datetime
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.filters import CommandStart
from aiogram.utils.keyboard import InlineKeyboardBuilder

API_TOKEN = os.getenv("BOT_TOKEN", "8583687270:AAFtlxuVDGnAsr_gWpPlc22nNcr4NEdD4Qg")
ADMIN_ID = 1243066883
EXCEL_FILE = "angren_akademiya.xlsx"
DB_FILE = "bot_data.db"

logging.basicConfig(level=logging.INFO)
bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT, phone TEXT, filial TEXT, courses TEXT, vaqt TEXT, sana TEXT, soat TEXT
        )
    ''')
    conn.commit()
    conn.close()

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Ro'yxat"
            ws.append(["№", "Ism", "Telefon", "Filial", "Kurslar", "Vaqt", "Sana", "Soat"])
            wb.save(EXCEL_FILE)
        except Exception as e:
            logging.error(f"Excel yaratishda xato: {e}")

def save_data(name, phone, filial, courses, vaqt):
    now = datetime.now()
    sana = now.strftime("%d.%m.%Y")
    soat = now.strftime("%H:%M")
    courses_str = ", ".join(courses)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO users (name, phone, filial, courses, vaqt, sana, soat)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (name, phone, filial, courses_str, vaqt, sana, soat))
    conn.commit()
    conn.close()
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        row_num = ws.max_row
        ws.append([row_num, name, phone, filial, courses_str, vaqt, sana, soat])
        wb.save(EXCEL_FILE)
    except Exception as e:
        logging.error(f"Excelga yozishda xato: {e}")

class RegState(StatesGroup):
    waiting_for_name = State()
    waiting_for_phone = State()
    waiting_for_filial = State()
    waiting_for_course = State()
    waiting_for_time = State()

VALID_COURSES = [
    "🏥 Kimyo va Biologiya",
    "🇬🇧 Ingliz tili — IELTS",
    "💻 IT — Sertifikat beriladi",
    "📐 Matematika",
    "👑 Prezident maktabiga tayyorlov",
    "📜 Tarix",
    "📖 Ona tili",
    "⚖️ Huquq",
    "🇰🇷 Koreys tili — TOPIK",
    "🎒 Pochemuchka",
    "✏️ Maktabga tayyorlov",
]

VALID_FILIALS = ["🏫 Angren filiali", "🏫 Ohangaron filiali"]
VALID_TIMES = ["🌅 Ertalabki guruh", "☀️ Kunduzgi guruh", "🌆 Kechqurungi guruh"]

def get_courses_inline(selected=None):
    if selected is None:
        selected = []
    builder = InlineKeyboardBuilder()
    for index, course in enumerate(VALID_COURSES):
        mark = "✅ " if course in selected else ""
        builder.button(text=f"{mark}{course}", callback_data=f"course_{index}")
    builder.button(text="✔️ Kurslarni Tasdiqlash", callback_data="confirm_courses")
    builder.adjust(1)
    return builder.as_markup()

def filial_keyboard():
    return types.ReplyKeyboardMarkup(keyboard=[
        [types.KeyboardButton(text="🏫 Angren filiali")],
        [types.KeyboardButton(text="🏫 Ohangaron filiali")]
    ], resize_keyboard=True, one_time_keyboard=True)

def time_keyboard():
    return types.ReplyKeyboardMarkup(keyboard=[
        [types.KeyboardButton(text="🌅 Ertalabki guruh")],
        [types.KeyboardButton(text="☀️ Kunduzgi guruh")],
        [types.KeyboardButton(text="🌆 Kechqurungi guruh")]
    ], resize_keyboard=True, one_time_keyboard=True)
    [22.06.2026 9:12] Yulduzoy Xudoyorova: import asyncio
import logging
import openpyxl
import os
import sqlite3
from datetime import datetime
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.filters import CommandStart
from aiogram.utils.keyboard import InlineKeyboardBuilder

API_TOKEN = os.getenv("BOT_TOKEN", "8583687270:AAFtlxuVDGnAsr_gWpPlc22nNcr4NEdD4Qg")
ADMIN_ID = 1243066883
EXCEL_FILE = "angren_akademiya.xlsx"
DB_FILE = "bot_data.db"

logging.basicConfig(level=logging.INFO)
bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT, phone TEXT, filial TEXT, courses TEXT, vaqt TEXT, sana TEXT, soat TEXT
        )
    ''')
    conn.commit()
    conn.close()

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Ro'yxat"
            ws.append(["№", "Ism", "Telefon", "Filial", "Kurslar", "Vaqt", "Sana", "Soat"])
            wb.save(EXCEL_FILE)
        except Exception as e:
            logging.error(f"Excel yaratishda xato: {e}")

def save_data(name, phone, filial, courses, vaqt):
    now = datetime.now()
    sana = now.strftime("%d.%m.%Y")
    soat = now.strftime("%H:%M")
    courses_str = ", ".join(courses)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO users (name, phone, filial, courses, vaqt, sana, soat)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (name, phone, filial, courses_str, vaqt, sana, soat))
    conn.commit()
    conn.close()
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        row_num = ws.max_row
        ws.append([row_num, name, phone, filial, courses_str, vaqt, sana, soat])
        wb.save(EXCEL_FILE)
    except Exception as e:
        logging.error(f"Excelga yozishda xato: {e}")

class RegState(StatesGroup):
    waiting_for_name = State()
    waiting_for_phone = State()
    waiting_for_filial = State()
    waiting_for_course = State()
    waiting_for_time = State()

VALID_COURSES = [
    "🏥 Kimyo va Biologiya",
    "🇬🇧 Ingliz tili — IELTS",
    "💻 IT — Sertifikat beriladi",
    "📐 Matematika",
    "👑 Prezident maktabiga tayyorlov",
    "📜 Tarix",
    "📖 Ona tili",
    "⚖️ Huquq",
    "🇰🇷 Koreys tili — TOPIK",
    "🎒 Pochemuchka",
    "✏️ Maktabga tayyorlov",
]

VALID_FILIALS = ["🏫 Angren filiali", "🏫 Ohangaron filiali"]
VALID_TIMES = ["🌅 Ertalabki guruh", "☀️ Kunduzgi guruh", "🌆 Kechqurungi guruh"]

def get_courses_inline(selected=None):
    if selected is None:
        selected = []
    builder = InlineKeyboardBuilder()
    for index, course in enumerate(VALID_COURSES):
        mark = "✅ " if course in selected else ""
        builder.button(text=f"{mark}{course}", callback_data=f"course_{index}")
    builder.button(text="✔️ Kurslarni Tasdiqlash", callback_data="confirm_courses")
    builder.adjust(1)
    return builder.as_markup()

def filial_keyboard():
    return types.ReplyKeyboardMarkup(keyboard=[
        [types.KeyboardButton(text="🏫 Angren filiali")],
        [types.KeyboardButton(text="🏫 Ohangaron filiali")]
    ], resize_keyboard=True, one_time_keyboard=True)

def time_keyboard():
    return types.ReplyKeyboardMarkup(keyboard=[
        [types.KeyboardButton(text="🌅 Ertalabki guruh")],
        [types.KeyboardButton(text="☀️ Kunduzgi guruh")],
        [types.KeyboardButton(text="🌆 Kechqurungi guruh")]
    ], resize_keyboard=True, one_time_keyboard=True)
[22.06.2026 9:12] Yulduzoy Xudoyorova: @dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    text = (
        "👋 Assalomu alaykum!\n"
        "«Angren Akademiya»ga xush kelibsiz!\n\n"
        "🎓 Biz sizning farzandingizni kelajakka tayyorlaymiz!\n"
        "🏆 Respublika va xalqaro olimpiadalar g'oliblari yetishtiramiz!\n\n"
        "🎁 Birinchi dars — BEPUL!\n\n"
        "✍️ Ro'yxatdan o'tish uchun Ism va Familiyangizni kiriting:"
    )
    await message.answer(text)
    await state.set_state(RegState.waiting_for_name)

@dp.message(RegState.waiting_for_name)
async def process_name(message: types.Message, state: FSMContext):
    name = message.text.strip()
    if len(name) < 3:
        await message.answer("❗️ Iltimos, to'liq ism va familiyangizni kiriting:")
        return
    await state.update_data(name=name)
    keyboard = types.ReplyKeyboardMarkup(
        keyboard=[[types.KeyboardButton(text="📱 Telefon raqamni yuborish", request_contact=True)]],
        resize_keyboard=True,
        one_time_keyboard=True
    )
    await message.answer(f"Rahmat, {name}! 😊\n\n📱 Telefon raqamingizni yuboring:", reply_markup=keyboard)
    await state.set_state(RegState.waiting_for_phone)

@dp.message(RegState.waiting_for_phone)
async def process_phone(message: types.Message, state: FSMContext):
    phone = message.contact.phone_number if message.contact else (message.text.strip() if message.text else None)
    if not phone:
        await message.answer("❗️ Iltimos telefon raqamingizni yuboring:")
        return
    await state.update_data(phone=phone)
    await message.answer("🏫 Qaysi filialni tanlaysiz?", reply_markup=filial_keyboard())
    await state.set_state(RegState.waiting_for_filial)

@dp.message(RegState.waiting_for_filial)
async def process_filial(message: types.Message, state: FSMContext):
    filial = message.text.strip()
    if filial not in VALID_FILIALS:
        await message.answer("❗️ Iltimos, quyidagi tugmalardan birini tanlang:", reply_markup=filial_keyboard())
        return
    await state.update_data(filial=filial, selected_courses=[])
    await message.answer(
        "📚 Qaysi kurslarga yozilmoqchisiz?\n\n"
        "✨ Bir yoki bir nechta kursni tanlashingiz mumkin.\n"
        "Tanlab bo'lgach, eng pastdagi ✔️ Kurslarni Tasdiqlash tugmasini bosing:",
        reply_markup=get_courses_inline([])
    )
    await state.set_state(RegState.waiting_for_course)

@dp.callback_query(F.data.startswith("course_"), RegState.waiting_for_course)
async def inline_course_click(callback: types.CallbackQuery, state: FSMContext):
    index = int(callback.data.split("_")[1])
    course_name = VALID_COURSES[index]
    user_data = await state.get_data()
    selected = user_data.get("selected_courses", [])
    if course_name in selected:
        selected.remove(course_name)
    else:
        selected.append(course_name)
    await state.update_data(selected_courses=selected)
    try:
        await callback.message.edit_reply_markup(reply_markup=get_courses_inline(selected))
    except Exception:
        pass
    await callback.answer()

@dp.callback_query(F.data == "confirm_courses", RegState.waiting_for_course)
async def inline_confirm_click(callback: types.CallbackQuery, state: FSMContext):
    user_data = await state.get_data()
    selected = user_data.get("selected_courses", [])
    if not selected:
        await callback.answer("❗️ Kamida bitta kurs tanlang!", show_alert=True)
        return
    try:
        await callback.message.delete()
    except Exception:
        pass
    await callback.message.answer("⏰ Qaysi vaqtda dars olishni xohlaysiz?", reply_markup=time_keyboard())
    await state.set_state(RegState.waiting_for_time)
    await callback.answer()
@dp.message(RegState.waiting_for_time)
async def process_time(message: types.Message, state: FSMContext):
    [22.06.2026 9:12] Yulduzoy Xudoyorova: import asyncio
import logging
import openpyxl
import os
import sqlite3
from datetime import datetime
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.filters import CommandStart
from aiogram.utils.keyboard import InlineKeyboardBuilder

API_TOKEN = os.getenv("BOT_TOKEN", "8583687270:AAFtlxuVDGnAsr_gWpPlc22nNcr4NEdD4Qg")
ADMIN_ID = 1243066883
EXCEL_FILE = "angren_akademiya.xlsx"
DB_FILE = "bot_data.db"

logging.basicConfig(level=logging.INFO)
bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT, phone TEXT, filial TEXT, courses TEXT, vaqt TEXT, sana TEXT, soat TEXT
        )
    ''')
    conn.commit()
    conn.close()

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Ro'yxat"
            ws.append(["№", "Ism", "Telefon", "Filial", "Kurslar", "Vaqt", "Sana", "Soat"])
            wb.save(EXCEL_FILE)
        except Exception as e:
            logging.error(f"Excel yaratishda xato: {e}")

def save_data(name, phone, filial, courses, vaqt):
    now = datetime.now()
    sana = now.strftime("%d.%m.%Y")
    soat = now.strftime("%H:%M")
    courses_str = ", ".join(courses)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO users (name, phone, filial, courses, vaqt, sana, soat)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (name, phone, filial, courses_str, vaqt, sana, soat))
    conn.commit()
    conn.close()
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        row_num = ws.max_row
        ws.append([row_num, name, phone, filial, courses_str, vaqt, sana, soat])
        wb.save(EXCEL_FILE)
    except Exception as e:
        logging.error(f"Excelga yozishda xato: {e}")

class RegState(StatesGroup):
    waiting_for_name = State()
    waiting_for_phone = State()
    waiting_for_filial = State()
    waiting_for_course = State()
    waiting_for_time = State()

VALID_COURSES = [
    "🏥 Kimyo va Biologiya",
    "🇬🇧 Ingliz tili — IELTS",
    "💻 IT — Sertifikat beriladi",
    "📐 Matematika",
    "👑 Prezident maktabiga tayyorlov",
    "📜 Tarix",
    "📖 Ona tili",
    "⚖️ Huquq",
    "🇰🇷 Koreys tili — TOPIK",
    "🎒 Pochemuchka",
    "✏️ Maktabga tayyorlov",
]

VALID_FILIALS = ["🏫 Angren filiali", "🏫 Ohangaron filiali"]
VALID_TIMES = ["🌅 Ertalabki guruh", "☀️ Kunduzgi guruh", "🌆 Kechqurungi guruh"]

def get_courses_inline(selected=None):
    if selected is None:
        selected = []
    builder = InlineKeyboardBuilder()
    for index, course in enumerate(VALID_COURSES):
        mark = "✅ " if course in selected else ""
        builder.button(text=f"{mark}{course}", callback_data=f"course_{index}")
    builder.button(text="✔️ Kurslarni Tasdiqlash", callback_data="confirm_courses")
    builder.adjust(1)
    return builder.as_markup()

def filial_keyboard():
    return types.ReplyKeyboardMarkup(keyboard=[
        [types.KeyboardButton(text="🏫 Angren filiali")],
        [types.KeyboardButton(text="🏫 Ohangaron filiali")]
    ], resize_keyboard=True, one_time_keyboard=True)

def time_keyboard():
    return types.ReplyKeyboardMarkup(keyboard=[
        [types.KeyboardButton(text="🌅 Ertalabki guruh")],
        [types.KeyboardButton(text="☀️ Kunduzgi guruh")],
        [types.KeyboardButton(text="🌆 Kechqurungi guruh")]
    ], resize_keyboard=True, one_time_keyboard=True)
[22.06.2026 9:12] Yulduzoy Xudoyorova: @dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    text = (
        "👋 Assalomu alaykum!\n"
        "«Angren Akademiya»ga xush kelibsiz!\n\n"
        "🎓 Biz sizning farzandingizni kelajakka tayyorlaymiz!\n"
        "🏆 Respublika va xalqaro olimpiadalar g'oliblari yetishtiramiz!\n\n"
        "🎁 Birinchi dars — BEPUL!\n\n"
        "✍️ Ro'yxatdan o'tish uchun Ism va Familiyangizni kiriting:"
    )
    await message.answer(text)
    await state.set_state(RegState.waiting_for_name)

@dp.message(RegState.waiting_for_name)
async def process_name(message: types.Message, state: FSMContext):
    name = message.text.strip()
    if len(name) < 3:
        await message.answer("❗️ Iltimos, to'liq ism va familiyangizni kiriting:")
        return
    await state.update_data(name=name)
    keyboard = types.ReplyKeyboardMarkup(
        keyboard=[[types.KeyboardButton(text="📱 Telefon raqamni yuborish", request_contact=True)]],
        resize_keyboard=True,
        one_time_keyboard=True
    )
    await message.answer(f"Rahmat, {name}! 😊\n\n📱 Telefon raqamingizni yuboring:", reply_markup=keyboard)
    await state.set_state(RegState.waiting_for_phone)

@dp.message(RegState.waiting_for_phone)
async def process_phone(message: types.Message, state: FSMContext):
    phone = message.contact.phone_number if message.contact else (message.text.strip() if message.text else None)
    if not phone:
        await message.answer("❗️ Iltimos telefon raqamingizni yuboring:")
        return
    await state.update_data(phone=phone)
    await message.answer("🏫 Qaysi filialni tanlaysiz?", reply_markup=filial_keyboard())
    await state.set_state(RegState.waiting_for_filial)

@dp.message(RegState.waiting_for_filial)
async def process_filial(message: types.Message, state: FSMContext):
    filial = message.text.strip()
    if filial not in VALID_FILIALS:
        await message.answer("❗️ Iltimos, quyidagi tugmalardan birini tanlang:", reply_markup=filial_keyboard())
        return
    await state.update_data(filial=filial, selected_courses=[])
    await message.answer(
        "📚 Qaysi kurslarga yozilmoqchisiz?\n\n"
        "✨ Bir yoki bir nechta kursni tanlashingiz mumkin.\n"
        "Tanlab bo'lgach, eng pastdagi ✔️ Kurslarni Tasdiqlash tugmasini bosing:",
        reply_markup=get_courses_inline([])
    )
    await state.set_state(RegState.waiting_for_course)

@dp.callback_query(F.data.startswith("course_"), RegState.waiting_for_course)
async def inline_course_click(callback: types.CallbackQuery, state: FSMContext):
    index = int(callback.data.split("_")[1])
    course_name = VALID_COURSES[index]
    user_data = await state.get_data()
    selected = user_data.get("selected_courses", [])
    if course_name in selected:
        selected.remove(course_name)
    else:
        selected.append(course_name)
    await state.update_data(selected_courses=selected)
    try:
        await callback.message.edit_reply_markup(reply_markup=get_courses_inline(selected))
    except Exception:
        pass
    await callback.answer()

@dp.callback_query(F.data == "confirm_courses", RegState.waiting_for_course)
async def inline_confirm_click(callback: types.CallbackQuery, state: FSMContext):
    user_data = await state.get_data()
    selected = user_data.get("selected_courses", [])
    if not selected:
        await callback.answer("❗️ Kamida bitta kurs tanlang!", show_alert=True)
        return
    try:
        await callback.message.delete()
    except Exception:
        pass
    await callback.message.answer("⏰ Qaysi vaqtda dars olishni xohlaysiz?", reply_markup=time_keyboard())
    await state.set_state(RegState.waiting_for_time)
    await callback.answer()

@dp.message(RegState.waiting_for_time)
async def process_time(message: types.Message, state: FSMContext):
[22.06.2026 9:12] Yulduzoy Xudoyorova: vaqt = message.text.strip()
    if vaqt not in VALID_TIMES:
        await message.answer("❗️ Iltimos, quyidagi tugmalardan birini tanlang:", reply_markup=time_keyboard())
        return
    data = await state.get_data()
    name = data.get("name")
    phone = data.get("phone")
    filial = data.get("filial")
    courses = data.get("selected_courses", [])
    save_data(name, phone, filial, courses, vaqt)
    courses_list = "\n".join([f"• {c}" for c in courses])
    await message.answer(
        f"🎉 Hurmatli {name}!\n\n"
        f"Bizni tanlaganingizdan mamnunmiz. 😊\n"
        f"Tez orada operatorlarimiz siz bilan bog'lanishadi! 🤝\n\n"
        f"📞 Yoki biz bilan bog'laning:\n"
        f"☎️ +998997925870\n"
        f"☎️ +998931015870\n\n"
        f"📋 Ma'lumotlaringiz:\n👤 Ism: {name}\n📱 Telefon: {phone}\n🏫 Filial: {filial}\n📚 Kurslar:\n{courses_list}\n⏰ Vaqt: {vaqt}",
        reply_markup=types.ReplyKeyboardRemove()
    )
    try:
        await bot.send_message(
            ADMIN_ID,
            f"🆕 Yangi ro'yxat!\n\n👤 Ism: {name}\n📱 Tel: {phone}\n🏫 Filial: {filial}\n📚 Kurslar:\n{courses_list}\n⏰ Vaqt: {vaqt}\n"
            f"🕐 Sana: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        )
    except Exception as e:
        logging.error(f"Adminga xabar ketmadi: {e}")
    await state.clear()

async def main():
    init_db()
    init_excel()
    await bot.delete_webhook(drop_pending_updates=True)
    print("✅ Bot professional Inline rejimda ishga tushdi!")
    await dp.start_polling(bot)

if__ name__ == "__main__":
    asyncio.run(main())

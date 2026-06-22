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

API_TOKEN = os.getenv("BOT_TOKEN")
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
            name TEXT, phone TEXT, filial TEXT,
            courses TEXT, vaqt TEXT, sana TEXT, soat TEXT
        )
    ''')
    conn.commit()
    conn.close()


def init_excel():
    if not os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Royxat"
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
    cursor.execute(
        "INSERT INTO users (name, phone, filial, courses, vaqt, sana, soat) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (name, phone, filial, courses_str, vaqt, sana, soat)
    )
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
    "Kimyo va Biologiya",
    "Ingliz tili IELTS",
    "IT Sertifikat beriladi",
    "Matematika",
    "Prezident maktabiga tayyorlov",
    "Tarix",
    "Ona tili",
    "Huquq",
    "Koreys tili TOPIK",
    "Pochemuchka",
    "Maktabga tayyorlov",
]

VALID_FILIALS = ["Angren filiali", "Ohangaron filiali"]
VALID_TIMES = ["Ertalabki guruh", "Kunduzgi guruh", "Kechqurungi guruh"]


def get_courses_inline(selected=None):
    if selected is None:
        selected = []
    builder = InlineKeyboardBuilder()
    for index, course in enumerate(VALID_COURSES):
        mark = "✅ " if course in selected else ""
        builder.button(text=f"{mark}{course}", callback_data=f"course_{index}")
    builder.button(text="✔️ Kurslarni Tasdiqlash", callback_data="confirm_courses")
    builder.button(text="🔙 Orqaga", callback_data="back_to_filial")
    builder.adjust(1)
    return builder.as_markup()


def phone_keyboard():
    return types.ReplyKeyboardMarkup(
        keyboard=[
            [types.KeyboardButton(text="📞 Telefon raqamni yuborish", request_contact=True)],
            [types.KeyboardButton(text="🔙 Orqaga")]
        ],
        resize_keyboard=True,
        one_time_keyboard=True
    )


def filial_keyboard():
    return types.ReplyKeyboardMarkup(
        keyboard=[
            [types.KeyboardButton(text="Angren filiali")],
            [types.KeyboardButton(text="Ohangaron filiali")],
            [types.KeyboardButton(text="🔙 Orqaga")]
        ],
        resize_keyboard=True,
        one_time_keyboard=True
    )


def time_keyboard():
    return types.ReplyKeyboardMarkup(
        keyboard=[
            [types.KeyboardButton(text="Ertalabki guruh")],
            [types.KeyboardButton(text="Kunduzgi guruh")],
            [types.KeyboardButton(text="Kechqurungi guruh")],
            [types.KeyboardButton(text="🔙 Orqaga")]
        ],
        resize_keyboard=True,
        one_time_keyboard=True
    )


# --- HANDLERLAR QISMI ---

@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await message.answer(
        "Assalomu alaykum! **Angren Akademiya** markazining rasmiy botiga xush kelibsiz.\n\n"
        "Ro‘yxatdan o‘tish uchun iltimos, **ismingiz va familiyangizni** kiriting:",
        parse_mode="Markdown"
    )
    await state.set_state(RegState.waiting_for_name)


@dp.message(RegState.waiting_for_name)
async def process_name(message: types.Message, state: FSMContext):
    if message.text == "🔙 Orqaga":
        await message.answer("Siz eng birinchi bosqichdasiz. Iltimos, ism-familiyangizni kiriting:")
        return

    await state.update_data(name=message.text)
    await message.answer(
        "Rahmat! Endi pastdagi tugma orqali telefon raqamingizni yuboring yoki qo‘lda kiriting (Masalan: +998901234567):",
        reply_markup=phone_keyboard()
    )
    await state.set_state(RegState.waiting_for_phone)


@dp.message(RegState.waiting_for_phone)
async def process_phone(message: types.Message, state: FSMContext):
    if message.text == "🔙 Orqaga":
        await message.answer("Ism va familiyangizni qaytadan kiriting:", reply_markup=types.ReplyKeyboardRemove())
        await state.set_state(RegState.waiting_for_name)
        return

    phone = message.contact.phone_number if message.contact else message.text
    await state.update_data(phone=phone)
    
    await message.answer("O‘zingizga qulay bo‘lgan filialni tanlang:", reply_markup=filial_keyboard())
    await state.set_state(RegState.waiting_for_filial)


@dp.message(RegState.waiting_for_filial)
async def process_filial(message: types.Message, state: FSMContext):
    if message.text == "🔙 Orqaga":
        await message.answer("Telefon raqamingizni qaytadan kiriting yoki yuboring:", reply_markup=phone_keyboard())
        await state.set_state(RegState.waiting_for_phone)
        return

    if message.text not in VALID_FILIALS:
        await message.answer("Iltimos, faqat pastdagi tugmalardan birini tanlang!", reply_markup=filial_keyboard())
        return

    await state.update_data(filial=message.text)
    await state.update_data(selected_courses=[])
    
    await message.answer(
        "Qaysi kurslarda o‘qishni xohlaysiz? Bir nechta variantni tanlashingiz mumkin.\n"
        "Tanlab bo‘lgach, **'✔️ Kurslarni Tasdiqlash'** tugmasini bosing:",
        reply_markup=get_courses_inline([]),
        parse_mode="Markdown"
    )
    await state.set_state(RegState.waiting_for_course)


@dp.callback_query(RegState.waiting_for_course, F.data.startswith("course_"))
async def process_course_selection(callback: types.CallbackQuery, state: FSMContext):
    course_index = int(callback.data.split("_")[1])
    course_name = VALID_COURSES[course_index]
    
    user_data = await state.get_data()
    selected_courses = user_data.get("selected_courses", [])
    
    if course_name in selected_courses:
        selected_courses.remove(course_name)
    else:
        selected_courses.append(course_name)
        
    await state.update_data(selected_courses=selected_courses)
    await callback.message.edit_reply_markup(reply_markup=get_courses_inline(selected_courses))
    await callback.answer()


@dp.callback_query(RegState.waiting_for_course, F.data == "back_to_filial")
async def back_to_filial(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.delete()
    await callback.message.answer("Filialni qaytadan tanlang:", reply_markup=filial_keyboard())
    await state.set_state(RegState.waiting_for_filial)
    await callback.answer()


@dp.callback_query(RegState.waiting_for_course, F.data == "confirm_courses")
async def confirm_courses(callback: types.CallbackQuery, state: FSMContext):
    user_data = await state.get_data()
    selected_courses = user_data.get("selected_courses", [])
    
    if not selected_courses:
        await callback.answer("Iltimos, kamida bitta kursni tanlang!", show_alert=True)
        return
        
    await callback.message.delete()
    await callback.message.answer("Darslar sizga qaysi vaqtda bo‘lishi qulay?", reply_markup=time_keyboard())
    await state.set_state(RegState.waiting_for_time)
    await callback.answer()


@dp.message(RegState.waiting_for_time)
async def process_time(message: types.Message, state: FSMContext):
    if message.text == "🔙 Orqaga":
        await message.answer(
            "Kurslarni qaytadan tanlang va tasdiqlang:",
            reply_markup=get_courses_inline([])
        )
        await state.set_state(RegState.waiting_for_course)
        return

    if message.text not in VALID_TIMES:
        await message.answer("Iltimos, quyidagi tugmalardan birini tanlang:", reply_markup=time_keyboard())
        return

    vaqt = message.text
    data = await state.get_data()
    name = data.get("name")
    phone = data.get("phone")
    filial = data.get("filial")
    courses = data.get("selected_courses", [])

    save_data(name, phone, filial, courses, vaqt)
    courses_list = "\n".join([f"- {c}" for c in courses])

    await message.answer(
        f"Hurmatli {name}!\n\n"
        f"Bizni tanlaganingizdan mamnunmiz.\n"
        f"Tez orada operatorlarimiz siz bilan boglanishadi!\n\n"
        f"Yoki biz bilan boglaning:\n"
        f"+998997925870\n"
        f"+998931015870\n\n"
        f"Malumotlaringiz:\n"
        f"Ism: {name}\n"
        f"Telefon: {phone}\n"
        f"Filial: {filial}\n"
        f"Kurslar:\n{courses_list}\n"
        f"Vaqt: {vaqt}",
        reply_markup=types.ReplyKeyboardRemove()
    )

    try:
        await bot.send_message(
            ADMIN_ID,
            f"Yangi royxat!\n\n"
            f"Ism: {name}\n"
            f"Tel: {phone}\n"
            f"Filial: {filial}\n"
            f"Kurslar:\n{courses_list}\n"
            f"Vaqt: {vaqt}\n"
            f"Sana: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        )
    except Exception as e:
        logging.error(f"Adminga xabar ketmadi: {e}")

    await state.clear()


async def main():
    init_db()
    init_excel()
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())

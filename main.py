import os
import asyncio
import logging
from datetime import datetime
from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.utils.keyboard import ReplyKeyboardBuilder

# Loglarni sozlash
logging.basicConfig(level=logging.INFO)

# Tokenni Render muhitidan o'qib olish
API_TOKEN = os.getenv("BOT_TOKEN")

if not API_TOKEN:
    raise ValueError("Xatolik: BOT_TOKEN topilmadi! Render sozlamalaridan uni tekshiring.")

bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# --- RO'YXATDAN O'TISH HOLATLARI (FSM) ---
class Registration(StatesGroup):
    name = State()         # Ism kutish
    phone = State()        # Telefon kutish
    filial = State()       # Filial kutish
    subjects = State()     # Fanlarni tanlash
    time_pref = State()    # O'qish vaqtini tanlash

# Markazingiz ma'lumotlari
AVAILABLE_FILIALS = ["Angren shahar", "Do'stlik Shaxarchasi", "Yangiobod"]
AVAILABLE_SUBJECTS = ["Kimyo", "Biologiya", "Matematika", "Fizika", "Ona tili", "Ingliz tili"]
AVAILABLE_TIMES = ["Ettalabgi", "Kunduzgi", "Kechki"]

# --- MA'LUMOTLAR BAZASI VA EXCEL BILAN ISHLASH ---
def init_db():
    import sqlite3
    conn = sqlite3.connect("bot_data.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            telegram_id INTEGER PRIMARY KEY,
            full_name TEXT,
            phone_number TEXT,
            filial TEXT,
            subjects TEXT,
            time_pref TEXT,
            reg_date TEXT,
            reg_time TEXT
        )
    """)
    conn.commit()
    conn.close()

def save_student_data(user_id, name, phone, filial, subjects_list, time_pref):
    import sqlite3
    import openpyxl
    from openpyxl import Workbook

    subjects_str = ", ".join(subjects_list)
    now = datetime.now()
    current_date = now.strftime("%Y-%m-%d")
    current_time = now.strftime("%H:%M:%S")

    # 1. SQLite bazaga saqlash
    conn = sqlite3.connect("bot_data.db")
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR REPLACE INTO users VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (user_id, name, phone, filial, subjects_str, time_pref, current_date, current_time))
    conn.commit()
    conn.close()

    # 2. Excel faylga saqlash (Siz ko'rsatgan tartibda)
    file_name = "angren_akademiya.xlsx"
    if os.path.exists(file_name):
        wb = openpyxl.load_workbook(file_name)
        sheet = wb.active
    else:
        wb = Workbook()
        sheet = wb.active
        sheet.title = "Ro'yxat"
        sheet.append(["№", "Ism", "Telefon", "Filial", "Kurslar", "Vaqt", "Sana", "Soat"])
    
    next_row = sheet.max_row
    sheet.append([next_row, name, phone, filial, subjects_str, time_pref, current_date, current_time])
    wb.save(file_name)
    logging.info(f"Excelga yozildi: {name}")

# --- TUGMALARNI SHAKLLANTIRISH ---
def get_keyboard(items):
    builder = ReplyKeyboardBuilder()
    for item in items:
        builder.button(text=item)
    builder.adjust(2)
    return builder.as_markup(resize_keyboard=True)

def get_subjects_keyboard(selected_subjects):
    builder = ReplyKeyboardBuilder()
    for subject in AVAILABLE_SUBJECTS:
        text = f"✅ {subject}" if subject in selected_subjects else subject
        builder.button(text=text)
    builder.button(text="📥 Davom etish (Vaqtni tanlash)")
    builder.adjust(2)
    return builder.as_markup(resize_keyboard=True)

# --- BOT HANDLERLARI ---

@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        f"Assalomu alaykum, {message.from_user.full_name}!\n"
        f"**'Angren Akademiya'** ta'lim markazi botiga xush kelibsiz.\n\n"
        f"Kurslarga ro'yxatdan o'tish uchun **To'liq ism-familiyangizni** kiriting:"
    )
    await state.set_state(Registration.name)

@dp.message(Registration.name)
async def process_name(message: types.Message, state: FSMContext):
    await state.update_data(name=message.text)
    await message.answer("Rahmat! Endi **Telefon raqamingizni** kiriting:")
    await state.set_state(Registration.phone)

@dp.message(Registration.phone)
async def process_phone(message: types.Message, state: FSMContext):
    await state.update_data(phone=message.text)
    await message.answer(
        "Iltimos, o'zingizga qulay **Filialni** tanlang:",
        reply_markup=get_keyboard(AVAILABLE_FILIALS)
    )
    await state.set_state(Registration.filial)

@dp.message(Registration.filial)
async def process_filial(message: types.Message, state: FSMContext):
    if message.text not in AVAILABLE_FILIALS:
        await message.answer("Iltimos, tugmalardan birini ishlating!")
        return
        
    await state.update_data(filial=message.text)
    await state.update_data(selected_subjects=[])
    
    await message.answer(
        "Ajoyib! Endi o'qimoqchi bo'lgan **Kurslarni (fanlarni)** tanlang (ketma-ket bir nechta tanlash mumkin):\n"
        "Tanlab bo'lgach, pastdagi **'📥 Davom etish (Vaqtni tanlash)'** tugmasini bosing:",
        reply_markup=get_subjects_keyboard([])
    )
    await state.set_state(Registration.subjects)

@dp.message(Registration.subjects)
async def process_subjects(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    selected_subjects = user_data.get("selected_subjects", [])
    text = message.text

    if text == "📥 Davom etish (Vaqtni tanlash)":
        if not selected_subjects:
            await message.answer("Iltimos, kamida bitta fanni tanlang!")
            return
        await message.answer(
            "O'qish uchun o'zingizga qulay **Vaqtni** belgilang:",
            reply_markup=get_keyboard(AVAILABLE_TIMES)
        )
        await state.set_state(Registration.time_pref)
        return

    clean_subject = text.replace("✅ ", "")
    if clean_subject in AVAILABLE_SUBJECTS:
        if clean_subject in selected_subjects:
            selected_subjects.remove(clean_subject)
        else:
            selected_subjects.append(clean_subject)
        
        await state.update_data(selected_subjects=selected_subjects)
        await message.answer("Yana fan tanlashingiz mumkin:", reply_markup=get_subjects_keyboard(selected_subjects))
    else:
        await message.answer("Iltimos, ro'yxatdagi tugmalardan foydalaning!")

@dp.message(Registration.time_pref)
async def process_time_pref(message: types.Message, state: FSMContext):
    if message.text not in AVAILABLE_TIMES:
        await message.answer("Iltimos, ro'yxatdagi vaqtlardan birini tanlang!")
        return

    user_data = await state.get_data()
    
    # Ma'lumotlarni saqlaymiz
    save_student_data(
        user_id=message.from_user.id,
        name=user_data['name'],
        phone=user_data['phone'],
        filial=user_data['filial'],
        subjects_list=user_data['selected_subjects'],
        time_pref=message.text
    )

    # Yakuniy xabar: Telefon raqamlaringiz bilan birga ko'rsatiladi
    await message.answer(
        f"Tabriklaymiz, ro'yxatdan muvaffaqiyatli o'tdingiz! 🎉\n"
        f"Sizning ma'lumotlaringiz 'Angren Akademiya' bazasiga to'liq joylashtirildi.\n\n"
        f"**Biz bilan bog'lanish uchun quyidagi telefon raqamlariga murojaat qilishingiz mumkin:**\n"
        f"📞 +998 94 041 42 55\n"
        f"📞 +998 93 101 58 70",
        reply_markup=types.ReplyKeyboardRemove()
    )
    await state.clear()

# --- SERVERNI ISHGA TUSHIRISH ---
async def main():
    init_db()
    await bot.delete_webhook(drop_pending_updates=True)
    
    import aiohttp.web
    async def dummy_handler(request):
        return aiohttp.web.Response(text="Angren Akademiya boti 24/7 faol!")
    
    app = aiohttp.web.Application()
    app.router.add_get("/", dummy_handler)
    runner = aiohttp.web.AppRunner(app)
    await runner.setup()
    
    port = int(os.getenv("PORT", 10000))
    site = aiohttp.web.TCPSite(runner, "0.0.0.0", port)
    
    await asyncio.gather(site.start(), dp.start_polling(bot))

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        logging.info("Bot to'xtatildi.")v

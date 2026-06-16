import asyncio
import logging
import openpyxl
import os
from datetime import datetime
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.filters import CommandStart

API_TOKEN = "8583687270:AAFnDh8oCD9i6Gn-R5zD-expOMuMfN8rNZ4"
ADMIN_ID = 1243066883
EXCEL_FILE = "angren_akademiya.xlsx"

logging.basicConfig(level=logging.INFO)
bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Ro'yxat"
        ws.append(["№", "Ism", "Telefon", "Kurs", "Sana", "Vaqt"])
        wb.save(EXCEL_FILE)

def save_to_excel(name, phone, course):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    row_num = ws.max_row
    now = datetime.now()
    ws.append([row_num, name, phone, course, now.strftime("%d.%m.%Y"), now.strftime("%H:%M")])
    wb.save(EXCEL_FILE)

class RegState(StatesGroup):
    waiting_for_name = State()
    waiting_for_phone = State()
    waiting_for_course = State()

@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await message.answer("👋 Assalomu alaykum!\n«Angren Akademiya» botiga xush kelibsiz.\n\nIsm va Familiyangizni kiriting:")
    await state.set_state(RegState.waiting_for_name)

@dp.message(RegState.waiting_for_name)
async def process_name(message: types.Message, state: FSMContext):
    await state.update_data(name=message.text)
    keyboard = types.ReplyKeyboardMarkup(keyboard=[[types.KeyboardButton(text="📱 Telefon raqamni yuborish", request_contact=True)]], resize_keyboard=True, one_time_keyboard=True)
    await message.answer("Telefon raqamingizni yuboring:", reply_markup=keyboard)
    await state.set_state(RegState.waiting_for_phone)

@dp.message(RegState.waiting_for_phone)
async def process_phone(message: types.Message, state: FSMContext):
    phone = message.contact.phone_number if message.contact else message.text
    await state.update_data(phone=phone)
    keyboard = types.ReplyKeyboardMarkup(keyboard=[
        [types.KeyboardButton(text="📐 Matematika"), types.KeyboardButton(text="🩺 Shifokorlik")],
        [types.KeyboardButton(text="💻 IT"), types.KeyboardButton(text="🇬🇧 Ingliz tili")],
        [types.KeyboardButton(text="🇷🇺 Rus tili"), types.KeyboardButton(text="📜 Tarix")],
        [types.KeyboardButton(text="👑 Prezident maktabi")],
        [types.KeyboardButton(text="🎒 Pochemuchka")]
    ], resize_keyboard=True)
    await message.answer("Kursni tanlang:", reply_markup=keyboard)
    await state.set_state(RegState.waiting_for_course)

@dp.message(RegState.waiting_for_course)
async def process_course(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    name = user_data['name']
    phone = user_data['phone']
    course = message.text
    save_to_excel(name, phone, course)
    await message.answer(f"🎉 Muvaffaqiyatli ro'yxatdan o'tdingiz!\n\n• Ism: {name}\n• Telefon: {phone}\n• Kurs: {course}\n\nTez orada bog'lanamiz!", reply_markup=types.ReplyKeyboardRemove())
    if ADMIN_ID:
        await bot.send_message(ADMIN_ID, f"🆕 Yangi ro'yxat!\n👤 {name}\n📱 {phone}\n📚 {course}\n🕐 {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    await state.clear()

async def main():
    init_excel()
    print("Bot ishga tushdi!")
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())

import os
import logging
from datetime import datetime
import asyncio
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.utils.keyboard import ReplyKeyboardBuilder, InlineKeyboardBuilder
from aiohttp import web, ClientSession
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment
from aiogram.types import FSInputFile

# Loglarni sozlash
logging.basicConfig(level=logging.INFO)

API_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_SHEET_URL = os.getenv("GOOGLE_SHEET_URL", "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID_HERE/edit")

if not API_TOKEN:
    raise ValueError("Xatolik: BOT_TOKEN topilmadi! Render sozlamalarini tekshiring.")

bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# EXCELGA YOZISH TIZIMI
EXCEL_FILE = "students.xlsx"


def save_to_excel(data):
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "O'quvchilar"
        ws.append(["Sana", "Ism Familiya", "Tel Raqam", "Ota-ona Tel", "Maktab", "Sinf", "Filial", "Smena", "Kurslar"])
        wb.save(EXCEL_FILE)
    
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    sana = datetime.now().strftime("%d.%m.%Y %H:%M")
    
    courses_list = data.get("selected_courses", [])
    courses_string = "\n".join([f"• {c.replace('\n', ' ')}" for c in courses_list])
    
    ws.append([
        sana,
        data.get("name"),
        data.get("phone"),
        data.get("parent_phone"),
        data.get("school"),
        data.get("grade"),
        data.get("filial"),
        data.get("time_pref"),
        courses_string
    ])
    
    for col in ws.columns:
        max_len = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        for cell in col:
            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
            if cell.value:
                lines = str(cell.value).split('\n')
                for line in lines:
                    if len(line) > max_len:
                        max_len = len(line)
        ws.column_dimensions[col_letter].width = max(min(max_len + 4, 38), 15)
        
    for row in ws.iter_rows(min_row=2):
        max_lines = 1
        for cell in row:
            if cell.value:
                lines_count = str(cell.value).count('\n') + 1
                if lines_count > max_lines:
                    max_lines = lines_count
        ws.row_dimensions[row[0].row].height = max_lines * 18 if max_lines > 1 else None

    wb.save(EXCEL_FILE)


# GOOGLE SHEETS TIZIMI
async def save_to_google_sheets(data):
    sana = datetime.now().strftime("%d.%m.%Y %H:%M")
    courses_list = data.get("selected_courses", [])
    courses_string = ", ".join([c.replace('\n', ' ') for c in courses_list])
    
    payload = {
        "sheet_url": GOOGLE_SHEET_URL,
        "sana": sana,
        "name": data.get("name"),
        "phone": data.get("phone"),
        "parent_phone": data.get("parent_phone"),
        "school": data.get("school"),
        "grade": data.get("grade"),
        "filial": data.get("filial"),
        "time_pref": data.get("time_pref"),
        "courses": courses_string
    }
    
    bridge_url = "https://script.google.com/macros/s/AKfycbzE3Zf_wF8rXOnfUu7f8vGg7hN-gQ246f_API/exec"
    try:
        async with ClientSession() as session:
            async with session.post(bridge_url, json=payload, timeout=10) as response:
                pass
    except Exception as e:
        logging.error(f"Google Sheets xatosi: {e}")


# --- RO'YXATDAN O'TISH HOLATLARI ---
class Registration(StatesGroup):
    name = State()
    phone = State()
    parent_phone = State()
    school = State()
    grade = State()
    filial = State()
    subjects = State()
    time_pref = State()


# --- AKADEMIYA ASOSIY MA'LUMOTLARI ---
AVAILABLE_FILIALS = ["Angren", "Ohangaron"]
AVAILABLE_TIMES = ["Ertalabki", "Kunduzgi", "Kechki"]
AVAILABLE_SUBJECTS = [
    "Matematika - Milliy va xalqaro sertifikat",
    "Matematika - majburiy blok ucun",
    "Ingliz tili - ILTES",
    "Tibbiyot - shifokorlik kasblari uchun\nKimyo - Milliy va xalqaro sertifikat",
    "Prezident maktablariga tayyorlov",
    "Al-Xorazmiy maktablariga tayyorlov",
    "Tibbiyot-shifokorlik kasbini tanlaganlar uchun-\nbiologiya - Milliy va xalqaro sertifikat",
    "Tarix- Milliy sertifikat",
    "Tarix -Majburiy blok uchun",
    "Huqu-Milliy sertifikat",
    "IT- Milliy v xalaro sertifikat",
    "Ona tili va adabiyoti -Milliy sertifikat",
    "Ona tili va adabiyoti -Majburiy blok uchun",
    "Maktabga tayyorlov. Pochemuchka"
]


def get_main_menu():
    kb = ReplyKeyboardBuilder()
    kb.button(text="📝 Ro'yxatdan o'tish")
    kb.button(text="📈 Bilim darajasini tekshirish")
    kb.button(text="🚪 Davomat (Keldim/Ketdim)")
    kb.adjust(1, 2)
    return kb.as_markup(resize_keyboard=True)


@dp.message(F.text == "/start")
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "✨ **Angren Akademiyasi** rasmiy botiga xush kelibsiz! \n\n"
        "Kelajak akademiyasida o'z bilimingizni va farzandingiz kamolotini nazorat qiling.",
        reply_markup=get_main_menu()
    )


@dp.message(F.text == "/excel")
async def send_excel(message: types.Message):
    if os.path.exists(EXCEL_FILE):
        excel_doc = FSInputFile(EXCEL_FILE)
        await message.answer_document(document=excel_doc, caption="📊 Angren Akademiya o'quvchilar ro'yxati (Excel)")
    else:
        await message.answer("❌ Hozircha ro'yxat bo'sh! Hech kim ro'yxatdan o'tmadi.")


@dp.message(F.text == "📈 Bilim darajasini tekshirish")
async def check_knowledge(message: types.Message):
    await message.answer(
        "📊 **\"Angren Akademiya\" — Bilim Nazorati Tizimi**\n\n"
        "✨ **Yaqin kunlarda hammasi yanada mukammal boʻladi!**\n\n"
        "Kelajakda farzandingiz bizning **\"Angren Akademiya\"** oʻquv markazimizni tanlaganda, "
        "ushbu tugma orqali har bir ota-ona aynan oʻz farzandining ismi, darsdagi ishtiroki va "
        "haqiqiy imtihon natijalari bilan muntazam tanishib borish imkoniyatiga ega boʻladi.\n\n"
        "🚀 *Biz kelajak texnologiyalarini taʼlimga olib kirmoqdamiz!*"
    )


@dp.message(F.text == "🚪 Davomat (Keldim/Ketdim)")
async def attendance_menu(message: types.Message):
    kb = InlineKeyboardBuilder()
    kb.button(text="🔔 Keldim", callback_data="attendance_in")
    kb.button(text="🔕 Ketdim", callback_data="attendance_out")
    kb.button(text="💰 Oylik to'lov holati", callback_data="attendance_pay") 
    kb.button(text="🚀 Uzoq muddatli imtiyozlar", callback_data="attendance_promo") 
    kb.adjust(2, 1, 1)
    await message.answer(
        "🚪 **Angren Akademiyasi — Davomat va Shaxsiy Balans**\n\nKerakli tugmani bosing:",
        reply_markup=kb.as_markup()
    )


@dp.callback_query(F.data.startswith("attendance_"))
async def process_attendance(callback: types.CallbackQuery):
    action = callback.data.split("_")[1]
    current_time = datetime.now().strftime("%H:%M")
    
    if action == "in":
        await callback.message.answer(f"🔔 **Angren Akademiya Xabarnomasi**\n\nFarzandingiz soat **{current_time}** da markazimizga eson-omon yetib keldi. 🔬")
    elif action == "out":
        await callback.message.answer(f"🔕 **Angren Akademiya Xabarnomasi**\n\nFarzandingiz soat **{current_time}** da darsdan chiqdi. Oq yo'l! ☀️")
    elif action == "pay":
        pay_kb = InlineKeyboardBuilder()
        pay_kb.button(text="💳 Plastik (Click/Payme)", callback_data="pay_via_card")
        pay_kb.button(text="💵 Naqd pul (Qo'lda)", callback_data="pay_via_cash")
        pay_kb.adjust(1)
        await callback.message.answer("💰 **To'lov usulini tanlang:**", reply_markup=pay_kb.as_markup())
    elif action == "promo":
        await callback.message.answer(
            "🚀 **\"Angren Akademiya\" Premium Imtiyozlar:**\n\n"
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
        await callback.message.answer(f"💳 **Karta raqami:** `{karta_raqam}`\n**Ega:** {karta_egasi}\n\nChekni adminga yuboring.")
    else:
        await callback.message.answer("💵 To'lovni administratorga topshiring. Rahmat!")
    await callback.answer()


# --- RO'YXATDAN O'TISH ---
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
    await state.set_state(Registration.subjects)


async def show_subjects_keyboard(message: types.Message, selected_courses: list):
    kb = InlineKeyboardBuilder()
    for idx, subject in enumerate(AVAILABLE_SUBJECTS):
        status = "✅" if subject in selected_courses else ""
        kb.button(text=f"{subject} {status}", callback_data=f"sub_{idx}")
    kb.button(text="➡️ Davom etish", callback_data="sub_done")
    kb.adjust(1)
    
    text = "📚 **Kurslarni tanlang:**\n\n"
    if selected_courses:
        text += "Tanlanganlar:\n" + "\n".join([f"- {c.replace('\n', ' ')}" for c in selected_courses])
        
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
        await callback.message.answer("Dars vaqtini tanlang:", reply_markup=kb.as_markup(resize_keyboard=True))
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


@dp.message(Registration.time_pref)
async def process_time_pref(message: types.Message, state: FSMContext):
    if message.text not in AVAILABLE_TIMES:
        await message.answer("Smenani tugmalardan tanlang!")
        return
    await state.update_data(time_pref=message.text)
    user_data = await state.get_data()
    
    save_to_excel(user_data)
    asyncio.create_task(save_to_google_sheets(user_data))
    
    selected_courses = user_data.get("selected_courses", [])
    courses_output = "📚 Tanlangan kurslar:\n" + "".join([f"• {c.replace('\n', ' ')}\n" for c in selected_courses])

    student_report = (
        f"🎉 **Muvaffaqiyatli ro'yxatdan o'tdingiz!**\n\n"
        f"👤 O'quvchi: {user_data.get('name')}\n"
        f"🏫 Maktab/Sinf: {user_data.get('school')}, {user_data.get('grade')}\n"
        f"📍 Filial: {user_data.get('filial')} | 🕒 Smena: {user_data.get('time_pref')}\n\n"
        f"{courses_output}\n"
        f"📞 +998 94 041 42 55\n📞 +998 93 101 58 70"
    )

    is_medical = any("Tibbiyot" in course for course in selected_courses)
    if is_medical:
        try:
            photo = FSInputFile("IMG_20260619_235730_628.jpg")
            await message.answer_photo(photo=photo, caption=student_report, parse_mode="Markdown", reply_markup=get_main_menu())
        except Exception as e:
            logging.error(f"Rasm yuborishda xato: {e}")
            await message.answer(text=student_report, parse_mode="Markdown", reply_markup=get_main_menu())
    else:
        await message.answer(text=student_report, parse_mode="Markdown", reply_markup=get_main_menu())
    await state.clear()


# --- RENDER PORTI ---
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

import os
import logging
from datetime import datetime
import asyncio
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.utils.keyboard import ReplyKeyboardBuilder, InlineKeyboardBuilder
from aiohttp import web, ClientSession  # Uyg'otuvchi ulanishi uchun ClientSession qo'shildi

# Loglarni sozlash
logging.basicConfig(level=logging.INFO)

API_TOKEN = os.getenv("BOT_TOKEN")
if not API_TOKEN:
    raise ValueError("Xatolik: BOT_TOKEN topilmadi! Render sozlamalarini tekshiring.")

bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# --- RO'YXATDAN O'TISH HOLATLARI (FSM ZANJIRI) ---
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
AVAILABLE_TIMES = ["Ettalabgi", "Kunduzgi", "Kechki"]
AVAILABLE_SUBJECTS = [
    "Matematika - Milliy va xalqaro sertifikat",
    "Matematika - majburiy blok ucun",
    "Ingliz tili - ILTES",
    "Tibbiyot - shifokorlik kasblari uchun Kimyo - Milliy va xalqaro sertifikat",
    "Prezident maktablariga tayyorlov",
    "Al-Xorazmiy maktablariga tayyorlov",
    "Tibbiyot-shifokorlik kasbini tanlaganlar uchun- biologiya - Milliy va xalqaro sertifikat",
    "Tarix- Milliy sertifikat",
    "Tarix -Majburiy blok uchun",
    "Huqu-Milliy sertifikat",
    "IT- Milliy v xalaro sertifikat",
    "Prezident maktablriga tayyorlov",
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

@dp.message(F.text == "📈 Bilim darajasini tekshirish")
async def check_knowledge(message: types.Message):
    today = datetime.now().day
    muhlat_sanasi = 15  
    tulov_qilingan = False 
    
    if not tulov_qilingan and today > muhlat_sanasi:
        await message.answer(
            "🔒 **\"Angren Akademiya\" — Tizim cheklangan**\n\n"
            f"Hurmatli ota-ona! Kelishilgan oylik to'lov muddati ({muhlat_sanasi}-sana) o'tib ketganligi sababli, "
            "farzandingizning bilim nazorati vaqtinchalik MUZLATILDI.\n\n"
            "Tizimni qayta faollashtirish uchun iltimos, oylik to'lovni amalga oshiring."
        )
    else:
        await message.answer(
            "📊 **\"Angren Akademiya\" Bilim Nazorati**\n\n"
            "👤 O'quvchi: **Kamolova Kamila**\n\n"
            "🔹 **Kimyo (Sertifikat kursi):** 26 / 30 ta to'g'ri (86%) ✅ *A'lo*\n"
            "🔹 **Matematika (Majburiy):** 9 / 10 ta to'g'ri (90%) ✅ *A'lo*\n\n"
            "💬 **Ustoz tavsiyasi:** O'zlashtirish juda yaxshi, shu zaylda davom eting!"
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
        "🚪 **Angren Akademiyasi — Davomat va Shaxsiy Balans**\n\n"
        "Kerakli tugmani bosing:",
        reply_markup=kb.as_markup()
    )

@dp.callback_query(F.data.startswith("attendance_"))
async def process_attendance(callback: types.CallbackQuery):
    action = callback.data.split("_")[1]
    current_time = datetime.now().strftime("%H:%M")
    darslar_soni = 2 
    
    if action == "in":
        darslar_soni += 1
        await callback.message.answer(
            f"🔔 **Angren Akademiya Xabarnomasi**\n\nFarzandingiz soat **{current_time}** da markazimizga eson-omon yetib keldi. 🔬"
        )
        if darslar_soni == 3:
            await callback.message.answer(
                "💳 **\"Angren Akademiya\" — To'lov Eslatmasi**\n\n"
                "Hurmatli ota-ona! Farzandingiz bugun markazimizda oʻzining **3-haqiqiy darsida** qatnashmoqda. Oylik to'lovni amalga oshirishingizni so'raymiz."
            )
        await callback.answer("Kelganingiz tasdiqlandi!", show_alert=True)
    elif action == "out":
        await callback.message.answer(f"🔕 **Angren Akademiya Xabarnomasi**\n\nFarzandingiz soat **{current_time}** da darsdan chiqdi. Oq yo'l! ☀️")
        await callback.answer("Ketganingiz tasdiqlandi!", show_alert=True)
    elif action == "pay":
        pay_kb = InlineKeyboardBuilder()
        pay_kb.button(text="💳 Plastik (Click/Payme)", callback_data="pay_via_card")
        pay_kb.button(text="💵 Naqd pul (Qo'lda)", callback_data="pay_via_cash")
        pay_kb.adjust(1)
        await callback.message.answer("💰 **To'lov usulini tanlang:**", reply_markup=pay_kb.as_markup())
        await callback.answer()
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
    await state.update_data(filial=message.text)
    await state.update_data(selected_courses=[])
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
        text += "Tanlanganlar:\n" + "\n".join([f"- {c}" for c in selected_courses])
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
    user_data = await state.get_data()
    selected_courses = user_data.get("selected_courses", [])
    courses_output = "📚 Tanlangan kurslar:\n" + "".join([f"• {c}\n" for c in selected_courses])

    student_report = (
        f"🎉 **Muvaffaqiyatli ro'yxatdan o'tdingiz!**\n\n"
        f"👤 O'quvchi: {user_data.get('name')}\n"
        f"🏫 Maktab/Sinf: {user_data.get('school')}, {user_data.get('grade')}\n"
        f"📍 Filial: {user_data.get('filial')} | 🕒 Smena: {message.text}\n\n"
        f"{courses_output}\n"
        f"📞 +998 94 041 42 55\n📞 +998 93 101 58 70"
    )

    is_medical = any("Tibbiyot" in course for course in selected_courses)
    if is_medical:
        from aiogram.types import FSInputFile
        try:
            photo = FSInputFile("IMG_20260619_235730_628.jpg")
            await message.answer_photo(photo=photo, caption=student_report, parse_mode="Markdown", reply_markup=get_main_menu())
        except Exception:
            await message.answer(text=student_report, parse_mode="Markdown", reply_markup=get_main_menu())
    else:
        await message.answer(text=student_report, parse_mode="Markdown", reply_markup=get_main_menu())
    await state.clear()

# --- RENDER PORTI (HEALTH CHECK) ---
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
    logging.info(f"Veb-server {port}-portda ishga tushdi.")

# --- 🚀 ENG MUHIM ZANJIR: UYG'OTUVCHI FUNKSIYA (SELF-PING) ---
async def pinger_loop():
    # Render sizga bergan shaxsiy URL manzilini Render boshqaruv panelidagi Environment Variables bo'limiga APP_URL kaliti bilan kiritasiz.
    # Masalan: APP_URL = https://angren-akademiya.onrender.com
    app_url = os.getenv("APP_URL")
    if not app_url:
        logging.warning("Diqqat: APP_URL muhit o'zgaruvchisi topilmadi. Bot o'z-oʻzini uyg'ota olmaydi!")
        return

    logging.info("O'z-o'zini uyg'otish (Self-Ping) tizimi ishga tushdi.")
    await asyncio.sleep(10) # Bot to'liq ishlab ketishi uchun ozgina kutish
    
    while True:
        try:
            async with ClientSession() as session:
                async with session.get(app_url) as response:
                    status = response.status
                    logging.info(f"Uyg'otish signali yuborildi. Status: {status}")
        except Exception as e:
            logging.error(f"Uyg'otishda xatolik yuz berdi: {e}")
        
        await asyncio.sleep(300) # Har 5 daqiqada (300 soniya) bir marta uyg'otib turadi

async def main():
    await start_web_server()
    # Uyg'otuvchi tsiklni orqa fonda parallel vazifa sifatida qo'shamiz
    asyncio.create_task(pinger_loop())
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())

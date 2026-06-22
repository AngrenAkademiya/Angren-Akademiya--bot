import os
import logging
from datetime import datetime
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.utils.keyboard import ReplyKeyboardBuilder, InlineKeyboardBuilder

# Loglarni sozlash (Render muhitida xatoliklarni kuzatish uchun)
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

# --- ASOSIY MENYU TUGMALARI ---
def get_main_menu():
    kb = ReplyKeyboardBuilder()
    kb.button(text="📝 Ro'yxatdan o'tish")
    kb.button(text="📈 Bilim darajasini tekshirish")
    kb.button(text="🚪 Davomat (Keldim/Ketdim)")
    kb.adjust(1, 2)
    return kb.as_markup(resize_keyboard=True)

# --- START BUYRUG'I ---
@dp.message(F.text == "/start")
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "✨ **Angren Akademiyasi** rasmiy botiga xush kelibsiz! \n\n"
        "Kelajak akademiyasida o'z bilimingizni va farzandingiz kamolotini nazorat qiling.",
        reply_markup=get_main_menu()
    )

# --- BILIM DARAJASINI TEKSHIRISH (BLOKLASH SANMSIYASI BILAN) ---
@dp.message(F.text == "📈 Bilim darajasini tekshirish")
async def check_knowledge(message: types.Message):
    # Muddat o'tib ketganda ota-onani bloklash mantiqi (Siz aytgan qat'iy intizom)
    # Kelajakda bazadan ota-onaning shaxsiy va'da qilgan sanasi (10, 12, 15-sana) tekshiriladi
    today = datetime.now().day
    muhlat_sanasi = 15  # Eng chekka muddat 15-sana qilib belgilangan
    tulov_qilingan = False # Agar to'lamagan bo'lsa va bugun 15 dan o'tgan bo'lsa
    
    if not tulov_qilingan and today > muhlat_sanasi:
        await message.answer(
            "🔒 **\"Angren Akademiya\" — Tizim cheklangan**\n\n"
            f"Hurmatli ota-ona! Kelishilgan oylik to'lov muddati ({muhlat_sanasi}-sana) o'tib ketganligi sababli, "
            "farzandingizning bilim nazorati, yangiliklar va davomat xabarnomalari tizimi vaqtincha MUZLATILDI.\n\n"
            "Tizimni qayta faollashtirish va cheklovlarni yechish uchun iltimos, oylik to'lovni amalga oshiring."
        )
    else:
        await message.answer(
            "📊 **\"Angren Akademiya\" Bilim Nazorati**\n\n"
            "👤 O'quvchi: **Kamolova Kamila**\n\n"
            "🔹 **Kimyo (Sertifikat kursi):** 26 / 30 ta to'g'ri (86%) ✅ *A'lo*\n"
            "🔹 **Matematika (Majburiy):** 9 / 10 ta to'g'ri (90%) ✅ *A'lo*\n"
            "🔹 **Tarix (Majburiy):** 6 / 10 ta to'g'ri (60%) ⚠️ *Yana ozroq harakat kerak*\n\n"
            "📉 **Umumiy reyting:** Guruhda 3-o'rinda (25 ta o'quvchi ichida).\n"
            "💬 **Ustoz tavsiyasi:** Kimyo va Matematikadan o'zlashtirish juda yaxshi. Tarix fani bo'yicha atamalar ustida ko'proq ishlash tavsiya etiladi."
        )

# --- DAVOMAT VA STRATEGIK YASHIRIN TO'LOV MENYUSI ---
@dp.message(F.text == "🚪 Davomat (Keldim/Ketdim)")
async def attendance_menu(message: types.Message):
    kb = InlineKeyboardBuilder()
    kb.button(text="🔔 Keldim", callback_data="attendance_in")
    kb.button(text="🔕 Ketdim", callback_data="attendance_out")
    kb.button(text="💰 Oylik to'lov holati", callback_data="attendance_pay") # Yangi kelgan ko'rmaydigan joy
    kb.button(text="🚀 Uzoq muddatli imtiyozlar", callback_data="attendance_promo") # Rag'batlantirish
    kb.adjust(2, 1, 1)
    
    await message.answer(
        "🚪 **Angren Akademiyasi — Davomat va Shaxsiy Balans**\n\n"
        "Darsga kelganingizni tasdiqlash yoki oylik hisobingizni tekshirish uchun kerakli tugmani bosing:",
        reply_markup=kb.as_markup()
    )

@dp.callback_query(F.data.startswith("attendance_"))
async def process_attendance(callback: types.CallbackQuery):
    action = callback.data.split("_")[1]
    current_time = datetime.now().strftime("%H:%M")
    
    # 3-dars mantiqi: Faqat oyog'i bilan darsga kelganda hisoblanadi (Baza simulyatsiyasi)
    # Har kelganda darslar_soni bittaga oshadi.
    darslar_soni = 2 # Namuna: O'quvchi bugun uchinchi marta eshikdan kirdi deb hisoblaymiz
    
    if action == "in":
        darslar_soni += 1
        
        # Ota-onaga standart xabar boradi (Agar foydalanuvchi bloklanmagan bo'lsa)
        await callback.message.answer(
            f"🔔 **Angren Akademiya Xabarnomasi**\n\n"
            f"Assalomu alaykum! Farzandingiz soat **{current_time}** da markazimizga eson-omon yetib keldi va darsga kirdi. 🔬"
        )
        
        # O'quvchi uydan turib emas, balki oyog'i bilan 3-darsga kelganida ota-onaga chiqadigan eslatma:
        if darslar_soni == 3:
            await callback.message.answer(
                "💳 **\"Angren Akademiya\" — To'lov Eslatmasi**\n\n"
                "Hurmatli ota-ona! Farzandingiz bugun markazimizda oʻzining **3-haqiqiy darsida** qatnashmoqda.\n\n"
                "Akademiyamiz tartib-qoidalariga koʻra, darslar davomiyligini taʼminlash uchun joriy oy toʻlovini amalga oshirishingizni soʻraymiz.\n"
                "To'lov turini tanlash uchun *'Oylik to'lov holati'* tugmasidan foydalanishingiz mumkin."
            )
        await callback.answer("Kelganingiz tasdiqlandi!", show_alert=True)
        
    elif action == "out":
        await callback.message.answer(
            f"🔕 **Angren Akademiya Xabarnomasi**\n\n"
            f"Farzandingiz soat **{current_time}** da bugungi darslarini tugatib, markazimizdan chiqdi. Oq yo'l! ☀️"
        )
        await callback.answer("Ketganingiz tasdiqlandi!", show_alert=True)
        
    elif action == "pay":
        # Naqd pul yoki plastik karta tanlovi
        pay_kb = InlineKeyboardBuilder()
        pay_kb.button(text="💳 Plastik karta (Click/Payme)", callback_data="pay_via_card")
        pay_kb.button(text="💵 Naqd pul (Qo'lda topshirish)", callback_data="pay_via_cash")
        pay_kb.adjust(1)
        
        await callback.message.answer(
            "💰 **\"Angren Akademiya\" To'lov Nazorati**\n\n"
            "Iltimos, o'zingizga qulay to'lov usulini tanlang:\n\n"
            "• Plastik orqali to'lov qilsangiz, chek rasmini botga yuklaysiz.\n"
            "• Naqd pul bo'lsa, administratorga topshirasiz va tizim faollashtiriladi.",
            reply_markup=pay_kb.as_markup()
        )
        await callback.answer()
        
    elif action == "promo":
        # Ko'p oylik to'lovlar uchun motivatsiya jadvali
        await callback.message.answer(
            "🚀 **\"Angren Akademiya\" Uzoq muddatli Premium Imtiyozlar va Motivatsiyalar:**\n\n"
            "🥈 **3 Oylik Oldindan (Silver Status):**\n"
            "• Umumiy summadan **10% chegirma** 💰\n"
            "• Akademiyaning maxsus brendlangan eksklyuziv daftari va ruchkasi sovg'a 🎁\n\n"
            "🥇 **6 Oylik Oldindan (Gold Status):**\n"
            "• Umumiy summadan **15% chegirma** 📉\n"
            "• Har oylik imtihonlarda professional ustozdan **shaxsiy konsultatsiya** 🧠\n"
            "• Markaz logotipi tushirilgan maxsus futbolka va kepka sovg'a 👕\n\n"
            "👑 **1 Yillik Oldindan (VIP Rezident):**\n"
            "• Umumiy summadan **20% chegirma** (2 oy bepul!) 🔥\n"
            "• Barcha darsliklar va kitoblar mutlaqo bepul 📚\n"
            "• Natija bermasa pulni qaytarish bo'yicha Kafolat shartnomasi 🤝"
        )
        await callback.answer()

# --- TO'LOV METODLARINI QABUL QILISH ---
@dp.callback_query(F.data.startswith("pay_via_"))
async def method_payment(callback: types.CallbackQuery):
    method = callback.data.split("_")[2]
    # Maxfiy ma'lumotlar Render muhitidan olinadi, yo'q bo'lsa standart namuna ko'rinadi
    karta_raqam = os.getenv("KARTA_RAQAMI", "8600 0000 0000 0000")
    karta_egasi = os.getenv("KARTA_EGASI", "Angren Akademiya Mas'ul Xodimi")
    
    if method == "card":
        await callback.message.answer(
            f"💳 **Plastik karta orqali oylik to'lov:**\n\n"
            f"• Karta raqami: `{karta_raqam}`\n"
            f"• Karta egasi: {karta_egasi}\n\n"
            f"📌 *To'lovni bajargach, iltimos, chek rasmini (skrinshot) shu yerga yuklang yoki administratorga yuboring. Administrator tasdiqlashi bilan tizimingiz to'liq faollashadi.*"
        )
    else:
        await callback.message.answer(
            "💵 **Naqd pul orqali to'lov shakli:**\n\n"
            "Siz naqd pul orqali hisob-kitob qilishni tanladingiz. Iltimos, darslar yakunlangunga qadar to'lovni "
            "markazimiz administratoriga topshiring. Xodimimiz to'lovni qabul qilib, tizimda profilingizni faollashtiradi. Rahmat!"
        )
    await callback.answer()

# --- RO'YXATDAN O'TISH ZANJIRI (ALGORITM) ---
@dp.message(F.text == "📝 Ro'yxatdan o'tish")
async def start_registration(message: types.Message, state: FSMContext):
    await message.answer("Ism va familiyangizni kiriting:")
    await state.set_state(Registration.name)

@dp.message(Registration.name)
async def process_name(message: types.Message, state: FSMContext):
    await state.update_data(name=message.text)
    await message.answer("📞 **O'quvchining shaxsiy telefon raqamini** kiriting:")
    await state.set_state(Registration.phone)

@dp.message(Registration.phone)
async def process_phone(message: types.Message, state: FSMContext):
    await state.update_data(phone=message.text)
    await message.answer("👨‍👩‍👦 **Ota-onangizning telefon raqamini** kiriting (Davomat xabarlari uchun):")
    await state.set_state(Registration.parent_phone)

@dp.message(Registration.parent_phone)
async def process_parent_phone(message: types.Message, state: FSMContext):
    await state.update_data(parent_phone=message.text)
    await message.answer("🏫 Hozirda **nechanchi maktabda** (yoki litsey/kollejda) o'qiysiz?")
    await state.set_state(Registration.school)

@dp.message(Registration.school)
async def process_school(message: types.Message, state: FSMContext):
    await state.update_data(school=message.text)
    await message.answer("🎓 **Nechanchi sinfda** o'qiysiz? (masalan: 9-sinf, 11-sinf):")
    await state.set_state(Registration.grade)

@dp.message(Registration.grade)
async def process_grade(message: types.Message, state: FSMContext):
    await state.update_data(grade=message.text)
    
    kb = ReplyKeyboardBuilder()
    for filial in AVAILABLE_FILIALS:
        kb.button(text=filial)
    kb.adjust(2)
    
    await message.answer("📍 O'zingizga yaqin **Filialni** tanlang:", reply_markup=kb.as_markup(resize_keyboard=True))
    await state.set_state(Registration.filial)

@dp.message(Registration.filial)
async def process_filial(message: types.Message, state: FSMContext):
    if message.text not in AVAILABLE_FILIALS:
        await message.answer("Iltimos, pastdagi tugmalardan birini bosing!")
        return
        
    await state.update_data(filial=message.text)
    await state.update_data(selected_courses=[])
    
    await show_subjects_keyboard(message, [])
    await state.set_state(Registration.subjects)

# --- MULTI-SELECT FANLAR PANELINI KO'RSATISH ---
async def show_subjects_keyboard(message: types.Message, selected_courses: list):
    kb = InlineKeyboardBuilder()
    for idx, subject in enumerate(AVAILABLE_SUBJECTS):
        status = "✅" if subject in selected_courses else ""
        kb.button(text=f"{subject} {status}", callback_data=f"sub_{idx}")
    
    kb.button(text="➡️ Davom etish (Smenani tanlash)", callback_data="sub_done")
    kb.adjust(1)
    
    text = "📚 **O'qimoqchi bo'lgan kurslaringizni tanlang (bir nechta tanlash mumkin):**\n\n"
    if selected_courses:
        text += "Tanlangan fanlar:\n" + "\n".join([f"- {c}" for c in selected_courses])
    
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
            await callback.answer("Iltimah, kamida bitta fan tanlang!", show_alert=True)
            return
            
        kb = ReplyKeyboardBuilder()
        for t in AVAILABLE_TIMES:
            kb.button(text=t)
        kb.adjust(3)
        
        await callback.message.answer("O'zingizga qulay dars **Vaqtini (Smenani)** tanlang:", reply_markup=kb.as_markup(resize_keyboard=True))
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

# --- YAKUNIY BOSQICH (AFISHA VA CONTACT MA'LUMOTLARI) ---
@dp.message(Registration.time_pref)
async def process_time_pref(message: types.Message, state: FSMContext):
    if message.text not in AVAILABLE_TIMES:
        await message.answer("Iltimos, smenani tugmalar orqali tanlang!")
        return
        
    user_data = await state.get_data()
    selected_courses = user_data.get("selected_courses", [])
    
    compulsory_text = "".join([f"• {c}\n" for c in selected_courses if "majburiy" in c.lower()])
    main_text = "".join([f"• {c}\n" for c in selected_courses if "majburiy" not in c.lower()])

    courses_output = ""
    if compulsory_text:
        courses_output += f"🔹 **Majburiy fanlar:**\n{compulsory_text}\n"
    if main_text:
        courses_output += f"🔸 **Tanlangan asosiy/sertifikat kurslari:**\n{main_text}"

    student_report = (
        f"🎉 **Siz muvaffaqiyatli ro'yxatdan o'tdingiz!**\n\n"
        f"👤 **O'quvchi:** {user_data.get('name')}\n"
        f"🏫 **Maktab/Sinf:** {user_data.get('school')}, {user_data.get('grade')}\n"
        f"📍 **Filial:** {user_data.get('filial')} | 🕒 **Smena:** {message.text}\n\n"
        f"{courses_output}\n\n"
        f"📌 **Bizning aloqa raqamlarimiz:**\n"
        f"📞 +998 94 041 42 55\n"
        f"📞 +998 93 101 58 70\n\n"
        f"Yaqin orada administratorlarimiz siz bilan bog'lanishadi. Akademiyamizga xush kelibsiz!"
    )

    is_medical = any("Tibbiyot" in course for course in selected_courses)

    if is_medical:
        from aiogram.types import FSInputFile
        try:
            # Tibbiyot afishasini yuborish
            photo = FSInputFile("IMG_20260619_235730_628.jpg")
            await message.answer_photo(photo=photo, caption=student_report, parse_mode="Markdown", reply_markup=get_main_menu())
        except Exception:
            # Agar rasm fayli topilmasa, kod to'xtab qolmasligi uchun fallback matn
            await message.answer(text=student_report, parse_mode="Markdown", reply_markup=get_main_menu())
    else:
        await message.answer(text=student_report, parse_mode="Markdown", reply_markup=get_main_menu())
        
    await state.clear()

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())

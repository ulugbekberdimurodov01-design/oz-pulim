<p align="center">
  <img src="rasmlar/sarlavha.svg" width="100%" alt="O‘z pulim — Surxondaryo budjetini oddiy tilda tushuntiradigan AI Telegram-bot">
</p>

<p align="center">
  <img src="https://img.shields.io/badge/LexHackathon-2026-E5534B?style=flat-square&labelColor=111827" alt="LexHackathon 2026">
  <img src="https://img.shields.io/badge/Python-3-F59E0B?style=flat-square&logo=python&logoColor=white&labelColor=111827" alt="Python 3">
  <img src="https://img.shields.io/badge/Telegram-bot-3B82F6?style=flat-square&logo=telegram&logoColor=white&labelColor=111827" alt="Telegram-bot">
  <img src="https://img.shields.io/badge/Gemini-AI-22C55E?style=flat-square&logo=googlegemini&logoColor=white&labelColor=111827" alt="Gemini AI">
</p>

<p align="center">
  <b>Har bir so‘m hisobda.</b> Kengash qarorlari ochiq, lekin tushunarsiz.<br>
  «O‘z pulim» ularni fuqaro tiliga o‘giradi va har bir javobni manbasi bilan beradi.
</p>

## ✨ Imkoniyatlar

| Bo‘lim | Nima beradi |
|:--|:--|
| 📥 **Kirim** | budjetga qancha pul tushgani |
| 📤 **Chiqim** | maktab, yo‘l, shifoxonaga qancha sarflangani |
| ➕ **Qo‘shimcha ajratmalar** | ortib qolgan pul kimga va nimaga berilgani |
| 📄 **Manba** | qaror raqami, sanasi va E-qaror havolasi — har bir javobda |
| ✍️ **Kengashga savol** | savolni rasmiy murojaat matniga aylantiradi |
| 🔔 **Obuna** | yangi qaror chiqqanda barcha foydalanuvchilarga xabar |
| 🌐 **3 til** | o‘zbek, ingliz, rus |

## 🧠 Qanday ishlaydi

<p align="center">
  <img src="rasmlar/jarayon.svg" width="100%" alt="PDF qarorlar → jadval → tekshiruv → AI javob">
</p>

Sun’iy intellekt (AI) internetdan emas, faqat Kengash qarorlaridan yig‘ilgan bazadan javob beradi: bazada yo‘q narsani aytmaydi, reja va haqiqiy sarfni aralashtirmaydi.

## 🚀 Ishga tushirish

Bot Google Colab’da — bulutli muhitda (cloud) ishlaydi:

1. Birinchi katakka (cell) kutubxonalarni (libraries) o‘rnating:
   ```python
   !pip install -q python-telegram-bot google-genai pandas openpyxl
   ```
2. Ikkinchi katakka [`O_z_pulim_bot.py`](O_z_pulim_bot.py) kodini qo‘ying va `BOT_TOKEN`, `GEMINI_API_KEY` ni yozing.
3. Katakni ishga tushiring: bot Excel fayllarni o‘zi topadi, topmasa — yuklashni so‘raydi.

> ⚠️ **Diqqat:** bot tokeni (token) va API kalitni (API key) GitHub’ga hech qachon yuklamang.

## ⚖️ Huquqiy asos

«O‘z pulim» yangi huquq yaratmaydi — mavjud huquqni amalda ishlatishni osonlashtiradi: O‘zbekiston Respublikasi Konstitutsiyasi, «Davlat hokimiyati va boshqaruvi organlari faoliyatining ochiqligi to‘g‘risida»gi (O‘RQ-369) va «Jamoatchilik nazorati to‘g‘risida»gi (O‘RQ-474) qonunlar.

---

<p align="center">
  <sub>Termiz davlat universiteti, yuridik fakultet talabalari · Ilmiy maslahatchi: Sayqalbek Allayorov<br>
  #LexHackathon2026 · #OchiqBudjet · Toshkent, 2026</sub>
</p>

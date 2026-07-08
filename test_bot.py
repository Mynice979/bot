import os
from dotenv import load_dotenv
from telegram.ext import Application

load_dotenv()
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
print("Token:", TOKEN[:5] + "..." if TOKEN else "Tidak ada")

app = Application.builder().token(TOKEN).build()
print("Bot berhasil dibuat")
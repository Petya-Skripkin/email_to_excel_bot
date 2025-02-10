import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from config import BOT_TOKEN, EXCEL_FILE
from email_reader import fetch_emails
import logging

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# Клавиатура
keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Скачать таблицу")]],
    resize_keyboard=True
)


# Команда /start и /download
@dp.message(Command(commands=["start", "download"]))
async def send_excel(message: types.Message):
    await message.answer("Нажмите кнопку, чтобы скачать таблицу с письмами", reply_markup=keyboard)

# Обработчик кнопки "Скачать таблицу"
@dp.message(lambda message: message.text == "Скачать таблицу")
async def send_file(message: types.Message):
    with open(EXCEL_FILE, "rb") as file:
        await bot.send_document(message.chat.id, file)

async def main():
    logging.basicConfig(level=logging.INFO)
    print("🚀 Бот запускается...")  # Добавили вывод
    fetch_emails()  # Читаем письма перед запуском бота
    await bot.delete_webhook(drop_pending_updates=True)
    print("✅ Бот запущен. Ожидаю сообщения...")  # Добавили вывод
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())

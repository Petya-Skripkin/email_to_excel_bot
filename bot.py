import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from config import BOT_TOKEN, EXCEL_FILE
from aiogram.types import FSInputFile
import logging
from email_reader import fetch_emails, monitor_emails

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
    file = FSInputFile("emails.xlsx")  # Указываем путь к файлу
    await bot.send_document(message.chat.id, document=file)

async def main():
    logging.basicConfig(level=logging.INFO)
    print("🚀 Бот запускается...")
    fetch_emails()  # Вызываем только один раз при старте
    await asyncio.gather(
        dp.start_polling(bot),  # Запускаем бот
        monitor_emails()  # Запускаем мониторинг почты
    )

if __name__ == "__main__":
    asyncio.run(main())

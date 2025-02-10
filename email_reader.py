import imaplib
import email
import pandas as pd
from email.header import decode_header
import asyncio
from email.utils import parsedate
from config import EMAIL, PASSWORD, IMAP_SERVER, FOLDER, CHECK_INTERVAL
from datetime import datetime, timedelta


async def fetch_emails():
    """Получает новые письма из почтового ящика."""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
        mail.login(EMAIL, PASSWORD)
        mail.select(FOLDER)

        _, messages = mail.search(None, "UNSEEN")  # Проверяем только новые (непрочитанные) письма
        messages = messages[0].split()

        if not messages:
            return []

        data_list = []
        for msg_id in messages:
            _, msg_data = mail.fetch(msg_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding or "utf-8")

                    sender = msg.get("From", "")
                    body = extract_body(msg)
                    
                                        # Извлекаем дату из заголовков
                    date = msg.get("Date")
                    if date:
                        # Парсим дату с использованием email.utils.parsedate
                        parsed_date = parsedate(date)
                        if parsed_date:
                            # Преобразуем в datetime
                            parsed_date = datetime(*parsed_date[:6])

                            # Добавляем 5 часов (хардкод)
                            parsed_date = parsed_date + timedelta(hours=5)

                            # Преобразуем в строку
                            date = parsed_date.strftime("%Y-%m-%d %H:%M:%S")
                        else:
                            date = "Не указана"
                    else:
                        date = "Не указана"

                    data = parse_email(body)
                    if data:
                        data["date"] = date
                        data_list.append(data)

        mail.logout()
        return data_list

    except imaplib.IMAP4.error as e:
        print(f"❌ Ошибка IMAP: {e}")
        return []


def extract_body(msg):
    """Извлекает текст из письма (только text/plain)."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    body = part.get_payload(decode=True).decode("utf-8")
                except UnicodeDecodeError:
                    body = part.get_payload(decode=True).decode("latin-1")
                break
    else:
        try:
            body = msg.get_payload(decode=True).decode("utf-8")
        except UnicodeDecodeError:
            body = msg.get_payload(decode=True).decode("latin-1")

    return body.strip()


def parse_email(body):
    """Парсит данные из тела письма"""
    data = {}
    fields = ["name", "company", "theme", "email", "message"]
    lines = body.split("\n")

    print("📩 Исходное тело письма:")
    print(body)  # Выведет тело письма для отладки

    for line in lines:
        if ":" in line:
            key, value = line.split(":", 1)
            key, value = key.strip(), value.strip()
            if key in fields:
                data[key] = value

    print("📊 Распознанные данные:", data)  # Выведет, какие поля удалось распарсить
    return {field: data.get(field, "") for field in fields}



async def save_to_excel(data, filename="emails.xlsx"):
    try:
        # Читаем существующий файл или создаём новый DataFrame с полями
        try:
            df = pd.read_excel(filename)
        except FileNotFoundError:
            # Если файла нет, создаём новый DataFrame с нужными полями
            df = pd.DataFrame(columns=["message", "name", "company", "email", "theme", "date"])

        # Добавляем новую строку с данными
        for entry in data:
            df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)

        # Сохраняем в файл
        df.to_excel(filename, index=False)
        print(f"✅ Данные успешно добавлены в {filename}")

    except Exception as e:
        print(f"❌ Ошибка при сохранении: {e}")


async def monitor_emails():
    """Мониторинг новых писем с периодическим опросом."""
    print("📩 Мониторинг писем запущен...")

    while True:
        new_emails = await fetch_emails()  # Сделаем fetch_emails асинхронным
        
        if new_emails:
            print(f"📬 Найдено {len(new_emails)} новых писем!")
            await save_to_excel(new_emails)  # Сделаем save_to_excel асинхронным
        else:
            print("📭 Новых писем нет...")

        await asyncio.sleep(CHECK_INTERVAL)  # Пауза перед следующим запросом (асинхронно)


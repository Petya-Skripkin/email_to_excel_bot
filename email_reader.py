import imaplib
import email
import pandas as pd
import time
from email.header import decode_header
from config import EMAIL, PASSWORD, IMAP_SERVER, FOLDER, CHECK_INTERVAL


def fetch_emails():
    """Получает новые письма из почтового ящика."""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
 #       mail.debug = 4
        mail.login(EMAIL, PASSWORD)
        mail.select(FOLDER)

        print(123)

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

                    data = parse_email(body)
                    if data:
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
    fields = ["message", "name", "company", "email", "theme"]
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



def save_to_excel(data_list):
    """Сохраняет данные в Excel файл (дописывает в существующий)."""
    if not data_list:
        return

    try:
        existing_df = pd.read_excel("emails.xlsx")
        df = pd.DataFrame(data_list)
        df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        df = pd.DataFrame(data_list)

    df.to_excel("emails.xlsx", index=False)
    print(f"✅ Добавлено {len(data_list)} новых писем в emails.xlsx")


def monitor_emails():
    """Мониторинг новых писем с периодическим опросом."""
    print("📩 Мониторинг писем запущен...")

    while True:
        new_emails = fetch_emails()
        
        if new_emails:
            print(f"📬 Найдено {len(new_emails)} новых писем!")
            save_to_excel(new_emails)
        else:
            print("📭 Новых писем нет...")

        time.sleep(CHECK_INTERVAL)  # Пауза перед следующим запросом

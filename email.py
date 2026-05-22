import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.header import Header
import os
from typing import List
from fastapi import HTTPException
import logging

from app.core.email_config import (
    SMTP_HOST, SMTP_PORT, SMTP_USER,
    SMTP_PASSWORD, SMTP_RIO_EMAIL,
    SMTP_FROM_EMAIL, SMTP_FROM_NAME
)

logger = logging.getLogger(__name__)


def send_email_with_attachments(
        to_email: str,
        subject: str,
        body: str,
        file_paths: List[str],
        file_names: List[str] = None
) -> bool:
    """
    Отправка email с вложениями
    """
    if not SMTP_USER or not SMTP_PASSWORD:
        raise HTTPException(
            status_code=500,
            detail="SMTP настройки не заданы. Проверьте .env файл"
        )

    try:
        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = to_email
        msg['Subject'] = Header(subject, 'utf-8').encode()

        full_body = f"От: {SMTP_FROM_NAME} <{SMTP_FROM_EMAIL}>\n\n{body}"
        msg.attach(MIMEText(full_body, 'plain', 'utf-8'))

        for idx, file_path in enumerate(file_paths):
            if not os.path.exists(file_path):
                logger.warning(f"Файл не найден: {file_path}")
                continue

            with open(file_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)

                if file_names and idx < len(file_names):
                    filename = file_names[idx]
                else:
                    filename = os.path.basename(file_path)

                part.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=Header(filename, 'utf-8').encode()
                )
                msg.attach(part)

        if SMTP_PORT == 465:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
                server.login(SMTP_USER, SMTP_PASSWORD)
                server.send_message(msg)
        else:
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASSWORD)
                server.send_message(msg)

        logger.info(f"Письмо успешно отправлено на {to_email}")
        return True

    except smtplib.SMTPAuthenticationError as e:
        logger.error(f"Ошибка авторизации: {e}")
        raise HTTPException(
            status_code=500,
            detail=f"Ошибка авторизации. Проверьте пароль приложения. {str(e)}"
        )
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        raise HTTPException(status_code=500, detail=f"Ошибка отправки: {str(e)}")


def send_to_rio(
        user_fio: str,
        user_email: str,
        comment: str,
        file_paths: List[str],
        file_names: List[str]
) -> bool:
    """
    Отправка файлов в РИО
    """
    subject = f"Новые методические материалы от {user_fio}"

    body = f"""
Уважаемые сотрудники РИО!

Пользователь {user_fio} ({user_email}) отправил на проверку следующие файлы:

{chr(10).join(f'- {name}' for name in file_names)}

Комментарий отправителя:
{comment if comment else 'Без комментария'}

---
Это автоматическое сообщение, пожалуйста, не отвечайте на него.

Сгенерировано системой проверки файлов.
    """

    return send_email_with_attachments(
        to_email=SMTP_RIO_EMAIL,
        subject=subject,
        body=body,
        file_paths=file_paths,
        file_names=file_names
    )
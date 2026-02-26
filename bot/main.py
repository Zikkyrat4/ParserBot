"""Telegram bot entry point: handlers for /start, /help, /template, text & file messages."""

from __future__ import annotations

import asyncio
import logging
import os
import re
from pathlib import Path

from dotenv import load_dotenv
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)

from bot.converter import parse_markdown, Metadata
from bot.docx_builder import build_docx
from bot.styles import WORK_TYPES
from bot.ai_processor import extract_text, generate_report

load_dotenv()
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

TEMPLATES_DIR = Path(__file__).parent / "templates"
MAX_FILE_SIZE = 1 * 1024 * 1024  # 1 MB

# Conversation states
EDIT_META, EDITING_FIELD, REPORT_INPUT = range(3)

# user_data keys
KEY_MD_TEXT = "md_text"
KEY_METADATA = "metadata"
KEY_BLOCKS = "blocks"
KEY_WORK_TYPE = "work_type"
KEY_EDITING_FIELD = "editing_field"

# Editable metadata fields: key → (display label, input prompt)
META_FIELDS = {
    "title": ("Название", "Введите название работы:"),
    "author": ("Студент", "Введите ФИО студента (например: Иванов Иван Иванович):"),
    "group": ("Группа", "Введите номер группы (например: ИЗ-41):"),
    "teacher": ("Преподаватель", "Введите ФИО преподавателя (например: Петров Пётр Петрович):"),
    "subject": ("Предмет", "Введите название предмета:"),
    "work_number": ("Номер работы", "Введите номер работы:"),
}

# Work type display labels
WORK_TYPE_LABELS = {
    "lab": "Лабораторная работа",
    "coursework": "Курсовая работа",
    "practice": "Отчёт по практике",
    "report": "Отчёт",
}


# --- Dashboard ---

def _build_dashboard_text(meta: Metadata, work_type: str) -> str:
    """Build the metadata dashboard message text."""
    wt_label = WORK_TYPE_LABELS.get(work_type, "—")
    return (
        "Данные титульного листа:\n\n"
        f"Тип работы: {wt_label}\n"
        f"Название: {meta.title or '—'}\n"
        f"Студент: {meta.author or '—'}\n"
        f"Группа: {meta.group or '—'}\n"
        f"Преподаватель: {meta.teacher or '—'}\n"
        f"Предмет: {meta.subject or '—'}\n"
        f"Номер работы: {meta.work_number or '—'}\n"
        "\nНажмите на поле, чтобы изменить значение."
    )


def _build_dashboard_keyboard() -> InlineKeyboardMarkup:
    """Build the inline keyboard for the metadata dashboard."""
    buttons = [
        [
            InlineKeyboardButton("Тип работы", callback_data="field:work_type"),
            InlineKeyboardButton("Название", callback_data="field:title"),
        ],
        [
            InlineKeyboardButton("Студент", callback_data="field:author"),
            InlineKeyboardButton("Группа", callback_data="field:group"),
        ],
        [
            InlineKeyboardButton("Преподаватель", callback_data="field:teacher"),
            InlineKeyboardButton("Предмет", callback_data="field:subject"),
        ],
        [
            InlineKeyboardButton("Номер работы", callback_data="field:work_number"),
        ],
        [
            InlineKeyboardButton("Сгенерировать DOCX", callback_data="meta:generate"),
        ],
        [
            InlineKeyboardButton("Отмена", callback_data="meta:cancel"),
        ],
    ]
    return InlineKeyboardMarkup(buttons)


def _work_type_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(label, callback_data=f"wt:{key}")]
        for key, label in WORK_TYPE_LABELS.items()
    ]
    buttons.append([InlineKeyboardButton("Назад", callback_data="wt:back")])
    return InlineKeyboardMarkup(buttons)


# --- Standalone commands (outside conversation) ---

async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Скачать шаблон", callback_data="get_template")],
    ])
    await update.message.reply_text(
        "Привет! Я конвертирую Markdown в DOCX по ГОСТ 7.32-2017.\n\n"
        "Отправьте мне:\n"
        "- текст в формате Markdown\n"
        "- файл .md или .txt\n\n"
        "Команды:\n"
        "/report — генерация отчёта через AI\n"
        "/help — поддерживаемые возможности\n"
        "/template — пример шаблона отчёта\n"
        "/cancel — отмена текущей операции",
        reply_markup=keyboard,
    )


async def help_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Поддерживаемые возможности Markdown:\n\n"
        "- Заголовки: # H1, ## H2, ### H3\n"
        "- Жирный, курсив, код\n"
        "- Нумерованные и маркированные списки\n"
        "- Блоки кода (с подсветкой языка)\n"
        "- Таблицы\n"
        "- Изображения (по URL)\n"
        "- Цитаты (> текст)\n\n"
        "YAML-заголовок (frontmatter):\n"
        "---\n"
        "title: Название работы\n"
        "author: Иванов И.И.\n"
        "group: ИТ-21\n"
        "teacher: Петров П.П.\n"
        "subject: Программирование\n"
        "work_number: 1\n"
        "---\n\n"
        "Поля, не заданные в frontmatter, можно заполнить через кнопки в боте.\n\n"
        "Генерация отчёта через AI:\n"
        "/report — отправьте любой материал (DOCX, PDF, TXT или текст),\n"
        "и AI сгенерирует структурированный отчёт."
    )


async def template_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    template_path = TEMPLATES_DIR / "example.md"
    if template_path.exists():
        with open(template_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename="example.md",
                caption="Пример шаблона лабораторной работы. Отредактируйте и отправьте обратно.",
            )
    else:
        await update.message.reply_text("Шаблон не найден.")


# --- Conversation: text input → dashboard ---

async def text_entry(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle plain text messages as Markdown input."""
    md_text = update.message.text
    if not md_text or len(md_text.strip()) < 3:
        await update.message.reply_text("Сообщение слишком короткое. Отправьте Markdown-текст.")
        return ConversationHandler.END

    try:
        metadata, blocks = parse_markdown(md_text)
    except ValueError as e:
        await update.message.reply_text(str(e))
        return ConversationHandler.END

    if not blocks:
        await update.message.reply_text(
            "Документ пуст — не найдено ни одного блока содержимого.\n"
            "Проверьте формат Markdown."
        )
        return ConversationHandler.END

    context.user_data[KEY_MD_TEXT] = md_text
    context.user_data[KEY_METADATA] = metadata
    context.user_data[KEY_BLOCKS] = blocks
    context.user_data.setdefault(KEY_WORK_TYPE, "lab")

    return await _send_dashboard(update.message, context)


# --- Conversation: file input → dashboard ---

async def document_entry(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle .md and .txt file uploads."""
    doc = update.message.document
    if not doc:
        return ConversationHandler.END

    if doc.file_size and doc.file_size > MAX_FILE_SIZE:
        await update.message.reply_text(
            "Файл слишком большой (максимум 1 МБ). Отправьте файл меньшего размера."
        )
        return ConversationHandler.END

    name = (doc.file_name or "").lower()
    if not (name.endswith(".md") or name.endswith(".txt") or name.endswith(".markdown")):
        await update.message.reply_text(
            "Поддерживаются файлы .md, .markdown и .txt.\n"
            "Отправьте файл в одном из этих форматов."
        )
        return ConversationHandler.END

    file = await doc.get_file()
    data = await file.download_as_bytearray()
    md_text = data.decode("utf-8", errors="replace")

    try:
        metadata, blocks = parse_markdown(md_text)
    except ValueError as e:
        await update.message.reply_text(str(e))
        return ConversationHandler.END

    if not blocks:
        await update.message.reply_text(
            "Документ пуст — не найдено ни одного блока содержимого.\n"
            "Проверьте формат Markdown."
        )
        return ConversationHandler.END

    context.user_data[KEY_MD_TEXT] = md_text
    context.user_data[KEY_METADATA] = metadata
    context.user_data[KEY_BLOCKS] = blocks
    context.user_data.setdefault(KEY_WORK_TYPE, "lab")

    await update.message.reply_text("Файл получен!")
    return await _send_dashboard(update.message, context)


# --- Dashboard helpers ---

async def _send_dashboard(message, context: ContextTypes.DEFAULT_TYPE, edit: bool = False) -> int:
    """Send or edit the metadata dashboard message."""
    meta: Metadata = context.user_data.get(KEY_METADATA, Metadata())
    work_type = context.user_data.get(KEY_WORK_TYPE, "lab")
    text = _build_dashboard_text(meta, work_type)
    keyboard = _build_dashboard_keyboard()

    if edit:
        await message.edit_text(text, reply_markup=keyboard)
    else:
        await message.reply_text(text, reply_markup=keyboard)
    return EDIT_META


# --- Dashboard button handlers ---

async def dashboard_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle all button presses on the dashboard."""
    query = update.callback_query
    await query.answer()
    data = query.data or ""

    # Field edit button
    if data.startswith("field:"):
        field_key = data[6:]

        # Work type gets its own keyboard
        if field_key == "work_type":
            await query.edit_message_text(
                "Выберите тип работы:",
                reply_markup=_work_type_keyboard(),
            )
            return EDIT_META

        # Text field — ask for input
        if field_key in META_FIELDS:
            _, prompt = META_FIELDS[field_key]
            context.user_data[KEY_EDITING_FIELD] = field_key
            await query.edit_message_text(f"{prompt}\n\n/cancel — отмена")
            return EDITING_FIELD

    # Work type selection
    if data.startswith("wt:"):
        wt = data[3:]
        if wt == "back":
            return await _send_dashboard(query.message, context, edit=True)
        if wt in WORK_TYPES:
            context.user_data[KEY_WORK_TYPE] = wt
        return await _send_dashboard(query.message, context, edit=True)

    # Generate
    if data == "meta:generate":
        meta: Metadata = context.user_data.get(KEY_METADATA, Metadata())
        if not meta.author:
            await query.answer("Заполните ФИО студента!", show_alert=True)
            return EDIT_META
        if not meta.teacher:
            await query.answer("Заполните ФИО преподавателя!", show_alert=True)
            return EDIT_META
        await query.edit_message_text("Генерирую документ...")
        return await _generate_and_send(query.message, context)

    # Cancel
    if data == "meta:cancel":
        await query.edit_message_text("Отменено. Отправьте Markdown заново.")
        context.user_data.clear()
        return ConversationHandler.END

    return EDIT_META


async def field_text_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Receive text input for a metadata field and return to dashboard."""
    field_key = context.user_data.pop(KEY_EDITING_FIELD, None)
    value = update.message.text.strip()

    if field_key and value:
        meta: Metadata = context.user_data.get(KEY_METADATA, Metadata())
        if hasattr(meta, field_key):
            setattr(meta, field_key, value)
            context.user_data[KEY_METADATA] = meta

    return await _send_dashboard(update.message, context)


# --- /report flow ---

async def report_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Entry point for /report command."""
    if not os.getenv("GEMINI_API_KEY"):
        await update.message.reply_text(
            "API ключ Gemini не настроен. Добавьте GEMINI_API_KEY в .env файл."
        )
        return ConversationHandler.END

    await update.message.reply_text(
        "Отправьте файл (DOCX, PDF, TXT) или текст для генерации отчёта.\n\n"
        "/cancel — отмена"
    )
    return REPORT_INPUT


async def _process_report(text: str, message, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Common logic: validate text, call AI, parse result, show dashboard."""
    if len(text.strip()) < 10:
        await message.reply_text("Недостаточно материала для генерации отчёта (минимум 10 символов).")
        return ConversationHandler.END

    status_msg = await message.reply_text("Генерирую отчёт...")

    try:
        work_type = context.user_data.get(KEY_WORK_TYPE, "lab")
        ai_markdown = await generate_report(text, work_type)
    except RuntimeError as e:
        await status_msg.edit_text(f"Ошибка: {e}")
        return ConversationHandler.END
    except Exception:
        logger.exception("Gemini API error")
        await status_msg.edit_text("Ошибка генерации, попробуйте позже.")
        return ConversationHandler.END

    if not ai_markdown.strip():
        await status_msg.edit_text("Не удалось сгенерировать отчёт.")
        return ConversationHandler.END

    try:
        metadata, blocks = parse_markdown(ai_markdown)
    except ValueError as e:
        await status_msg.edit_text(f"Ошибка разбора результата: {e}")
        return ConversationHandler.END

    if not blocks:
        await status_msg.edit_text("Не удалось сгенерировать отчёт.")
        return ConversationHandler.END

    context.user_data[KEY_MD_TEXT] = ai_markdown
    context.user_data[KEY_METADATA] = metadata
    context.user_data[KEY_BLOCKS] = blocks
    context.user_data.setdefault(KEY_WORK_TYPE, "lab")

    await status_msg.edit_text("Отчёт сгенерирован!")
    return await _send_dashboard(message, context)


async def report_text_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle plain text sent after /report."""
    return await _process_report(update.message.text, update.message, context)


async def report_file_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle file sent after /report."""
    doc = update.message.document
    if not doc:
        await update.message.reply_text("Отправьте файл или текст.")
        return REPORT_INPUT

    if doc.file_size and doc.file_size > MAX_FILE_SIZE:
        await update.message.reply_text("Файл слишком большой (максимум 1 МБ).")
        return REPORT_INPUT

    name = doc.file_name or ""
    name_lower = name.lower()
    if not (name_lower.endswith(".docx") or name_lower.endswith(".pdf") or name_lower.endswith(".txt")):
        await update.message.reply_text("Поддерживаются файлы .docx, .pdf, .txt.")
        return REPORT_INPUT

    file = await doc.get_file()
    data = bytes(await file.download_as_bytearray())

    try:
        text = extract_text(data, name)
    except Exception:
        logger.exception("Text extraction error")
        await update.message.reply_text("Не удалось извлечь текст из файла.")
        return ConversationHandler.END

    if not text.strip():
        await update.message.reply_text("Не удалось извлечь текст из файла (возможно, это скан).")
        return ConversationHandler.END

    return await _process_report(text, update.message, context)


# --- Generate ---

async def _generate_and_send(message, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Build DOCX and send it back."""
    meta: Metadata = context.user_data.get(KEY_METADATA, Metadata())
    blocks = context.user_data.get(KEY_BLOCKS, [])
    work_type = context.user_data.get(KEY_WORK_TYPE, "lab")

    try:
        docx_buf = build_docx(blocks, meta, work_type)

        title = meta.title or "document"
        safe_title = re.sub(r'[^\w \-]', '_', title).strip() or "document"
        filename = f"{safe_title}.docx"

        await message.reply_document(
            document=docx_buf,
            filename=filename,
            caption="Готово!",
        )
    except Exception:
        logger.exception("Error generating DOCX")
        await message.reply_text("Ошибка при генерации документа. Проверьте формат Markdown.")
    finally:
        context.user_data.pop(KEY_MD_TEXT, None)
        context.user_data.pop(KEY_METADATA, None)
        context.user_data.pop(KEY_BLOCKS, None)
        context.user_data.pop(KEY_WORK_TYPE, None)
        context.user_data.pop(KEY_EDITING_FIELD, None)

    return ConversationHandler.END


async def cancel_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Cancel the conversation."""
    await update.message.reply_text("Отменено. Отправьте Markdown заново.")
    context.user_data.clear()
    return ConversationHandler.END


async def template_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle inline button for template download."""
    query = update.callback_query
    await query.answer()
    template_path = TEMPLATES_DIR / "example.md"
    if template_path.exists():
        with open(template_path, "rb") as f:
            await query.message.reply_document(
                document=f,
                filename="example.md",
                caption="Пример шаблона. Отредактируйте и отправьте обратно.",
            )
    else:
        await query.message.reply_text("Шаблон не найден.")


def main() -> None:
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("BOT_TOKEN is not set. Create a .env file with BOT_TOKEN=...")

    app = Application.builder().token(token).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("report", report_start),
            MessageHandler(filters.Document.ALL, document_entry),
            MessageHandler(filters.TEXT & ~filters.COMMAND, text_entry),
        ],
        states={
            EDIT_META: [
                CallbackQueryHandler(dashboard_callback),
            ],
            EDITING_FIELD: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, field_text_received),
            ],
            REPORT_INPUT: [
                MessageHandler(filters.Document.ALL, report_file_received),
                MessageHandler(filters.TEXT & ~filters.COMMAND, report_text_received),
            ],
        },
        fallbacks=[
            CommandHandler("cancel", cancel_handler),
            CommandHandler("start", start_handler),
        ],
    )

    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("help", help_handler))
    app.add_handler(CommandHandler("template", template_handler))
    app.add_handler(CallbackQueryHandler(template_callback, pattern=r"^get_template$"))
    app.add_handler(conv)

    logger.info("Bot started")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        asyncio.set_event_loop(asyncio.new_event_loop())
    main()

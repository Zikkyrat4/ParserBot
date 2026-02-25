"""Telegram bot entry point: handlers for /start, /help, /template, text & file messages."""

from __future__ import annotations

import asyncio
import logging
import os
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

load_dotenv()
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

TEMPLATES_DIR = Path(__file__).parent / "templates"

# Conversation states
ASK_WORK_TYPE, ASK_AUTHOR, ASK_TEACHER, GENERATE = range(4)

# user_data keys
KEY_MD_TEXT = "md_text"
KEY_METADATA = "metadata"
KEY_BLOCKS = "blocks"
KEY_WORK_TYPE = "work_type"


# --- Keyboards ---

def _work_type_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(label, callback_data=f"wt:{key}")]
        for key, label in {
            "lab": "Лабораторная работа",
            "coursework": "Курсовая работа",
            "practice": "Отчёт по практике",
            "report": "Отчёт",
        }.items()
    ]
    return InlineKeyboardMarkup(buttons)


# --- Standalone commands (outside conversation) ---

async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Привет! Я конвертирую Markdown в DOCX по ГОСТ 7.32-2017.\n\n"
        "Отправьте мне:\n"
        "• текст в формате Markdown\n"
        "• файл .md или .txt\n\n"
        "Команды:\n"
        "/help — поддерживаемые возможности\n"
        "/template — пример шаблона отчёта"
    )


async def help_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Поддерживаемые возможности Markdown:\n\n"
        "• Заголовки: # H1, ## H2, ### H3\n"
        "• Жирный, курсив, код\n"
        "• Нумерованные и маркированные списки\n"
        "• Блоки кода\n"
        "• Таблицы\n\n"
        "YAML-заголовок (frontmatter):\n"
        "---\n"
        "title: Название работы\n"
        "author: Иванов И.И.\n"
        "group: ИТ-21\n"
        "teacher: Петров П.П.\n"
        "subject: Программирование\n"
        "work_number: 1\n"
        "---\n\n"
        "Если author/teacher не указаны в frontmatter, бот спросит их отдельно."
    )


async def template_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    template_path = TEMPLATES_DIR / "example.md"
    if template_path.exists():
        await update.message.reply_document(
            document=open(template_path, "rb"),
            filename="example.md",
            caption="Пример шаблона лабораторной работы. Отредактируйте и отправьте обратно.",
        )
    else:
        await update.message.reply_text("Шаблон не найден.")


# --- Conversation: text input ---

async def text_entry(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle plain text messages as Markdown input."""
    md_text = update.message.text
    if not md_text or len(md_text.strip()) < 5:
        await update.message.reply_text("Сообщение слишком короткое. Отправьте Markdown-текст.")
        return ConversationHandler.END

    metadata, blocks = parse_markdown(md_text)
    context.user_data[KEY_MD_TEXT] = md_text
    context.user_data[KEY_METADATA] = metadata
    context.user_data[KEY_BLOCKS] = blocks

    await update.message.reply_text(
        "Выберите тип работы:",
        reply_markup=_work_type_keyboard(),
    )
    return ASK_WORK_TYPE


# --- Conversation: file input ---

async def document_entry(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle .md and .txt file uploads."""
    doc = update.message.document
    if not doc:
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

    metadata, blocks = parse_markdown(md_text)
    context.user_data[KEY_MD_TEXT] = md_text
    context.user_data[KEY_METADATA] = metadata
    context.user_data[KEY_BLOCKS] = blocks

    await update.message.reply_text(
        "Файл получен. Выберите тип работы:",
        reply_markup=_work_type_keyboard(),
    )
    return ASK_WORK_TYPE


# --- Conversation: work type → maybe ask author → maybe ask teacher → generate ---

async def work_type_chosen(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle work type selection, then ask for missing metadata."""
    query = update.callback_query
    await query.answer()

    data = query.data or ""
    if not data.startswith("wt:"):
        return ASK_WORK_TYPE

    work_type = data[3:]
    context.user_data[KEY_WORK_TYPE] = work_type

    meta: Metadata = context.user_data.get(KEY_METADATA, Metadata())

    # Ask for author if missing
    if not meta.author:
        await query.edit_message_text("Введите ФИО студента (например: Иванов Иван Иванович):")
        return ASK_AUTHOR

    # Ask for teacher if missing
    if not meta.teacher:
        await query.edit_message_text("Введите ФИО преподавателя:")
        return ASK_TEACHER

    # All metadata present — generate
    await query.edit_message_text("⏳ Генерирую документ...")
    return await _generate_and_send(query.message, context)


async def author_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Receive student name from user."""
    meta: Metadata = context.user_data.get(KEY_METADATA, Metadata())
    meta.author = update.message.text.strip()
    context.user_data[KEY_METADATA] = meta

    if not meta.teacher:
        await update.message.reply_text("Введите ФИО преподавателя:")
        return ASK_TEACHER

    await update.message.reply_text("⏳ Генерирую документ...")
    return await _generate_and_send(update.message, context)


async def teacher_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Receive teacher name from user."""
    meta: Metadata = context.user_data.get(KEY_METADATA, Metadata())
    meta.teacher = update.message.text.strip()
    context.user_data[KEY_METADATA] = meta

    await update.message.reply_text("⏳ Генерирую документ...")
    return await _generate_and_send(update.message, context)


async def _generate_and_send(message, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Build DOCX and send it back."""
    meta: Metadata = context.user_data.get(KEY_METADATA, Metadata())
    blocks = context.user_data.get(KEY_BLOCKS, [])
    work_type = context.user_data.get(KEY_WORK_TYPE, "lab")

    try:
        docx_buf = build_docx(blocks, meta, work_type)

        title = meta.title or "document"
        safe_title = "".join(c if c.isalnum() or c in " _-" else "_" for c in title)
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

    return ConversationHandler.END


async def cancel_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Cancel the conversation."""
    await update.message.reply_text("Отменено. Отправьте Markdown заново.")
    context.user_data.clear()
    return ConversationHandler.END


def main() -> None:
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("BOT_TOKEN is not set. Create a .env file with BOT_TOKEN=...")

    app = Application.builder().token(token).build()

    conv = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Document.ALL, document_entry),
            MessageHandler(filters.TEXT & ~filters.COMMAND, text_entry),
        ],
        states={
            ASK_WORK_TYPE: [
                CallbackQueryHandler(work_type_chosen, pattern=r"^wt:"),
            ],
            ASK_AUTHOR: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, author_received),
            ],
            ASK_TEACHER: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, teacher_received),
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
    app.add_handler(conv)

    logger.info("Bot started")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        asyncio.set_event_loop(asyncio.new_event_loop())
    main()

"""AI report generation via Gemini 2.0 Flash and text extraction from various formats."""

from __future__ import annotations

import io
import logging
import os

from google import genai

logger = logging.getLogger(__name__)

REPORT_PROMPT = """\
Ты — помощник для написания академических отчётов.
На основе предоставленного материала напиши структурированный отчёт
в формате Markdown.

Требования:
- Используй заголовки: # для основных разделов, ## для подразделов
- Включи: Введение, основную часть, Заключение
- Пиши на русском языке, академический стиль
- Тип работы: {work_type_label}
- НЕ добавляй YAML frontmatter

Исходный материал:
{text}"""

WORK_TYPE_LABELS = {
    "lab": "Лабораторная работа",
    "coursework": "Курсовая работа",
    "practice": "Отчёт по практике",
    "report": "Отчёт",
}


def extract_text_from_docx(data: bytes) -> str:
    """Extract plain text from a DOCX file."""
    from docx import Document

    doc = Document(io.BytesIO(data))
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def extract_text_from_pdf(data: bytes) -> str:
    """Extract plain text from a PDF file."""
    from PyPDF2 import PdfReader

    reader = PdfReader(io.BytesIO(data))
    pages = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            pages.append(text)
    return "\n".join(pages)


def extract_text(data: bytes, filename: str) -> str:
    """Dispatch text extraction by file extension."""
    name = filename.lower()
    if name.endswith(".docx"):
        return extract_text_from_docx(data)
    if name.endswith(".pdf"):
        return extract_text_from_pdf(data)
    if name.endswith(".txt"):
        return data.decode("utf-8", errors="replace")
    raise ValueError(f"Unsupported format: {filename}")


async def generate_report(text: str, work_type: str) -> str:
    """Call Gemini 2.0 Flash to generate a structured Markdown report."""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY is not set")

    client = genai.Client(api_key=api_key)

    work_type_label = WORK_TYPE_LABELS.get(work_type, "Отчёт")
    prompt = REPORT_PROMPT.format(work_type_label=work_type_label, text=text)

    response = await client.aio.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
    )

    if not response.text:
        raise ValueError("AI returned empty result")

    return response.text

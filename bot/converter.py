"""Markdown parser: extracts frontmatter and converts MD to typed blocks."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional

import yaml
import mistune


# --- Block dataclasses ---

@dataclass
class Run:
    """A piece of text with formatting."""
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False


@dataclass
class HeadingBlock:
    level: int
    text: str


@dataclass
class ParagraphBlock:
    runs: List[Run] = field(default_factory=list)


@dataclass
class ListBlock:
    ordered: bool
    items: List[List[Run]] = field(default_factory=list)


@dataclass
class CodeBlock:
    language: str
    code: str


@dataclass
class TableBlock:
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)


@dataclass
class ImageBlock:
    alt: str
    url: str


@dataclass
class BlockquoteBlock:
    runs: List[Run] = field(default_factory=list)


Block = HeadingBlock | ParagraphBlock | ListBlock | CodeBlock | TableBlock | ImageBlock | BlockquoteBlock


# --- Frontmatter extraction ---

_FRONTMATTER_RE = re.compile(r"\A---\s*\n(.*?)\n---\s*\n", re.DOTALL)


@dataclass
class Metadata:
    title: str = ""
    author: str = ""
    group: str = ""
    teacher: str = ""
    subject: str = ""
    university: str = ""
    year: str = ""
    work_number: str = ""
    institute: str = ""
    department: str = ""
    city: str = ""


def extract_frontmatter(text: str) -> tuple[Metadata, str]:
    """Extract YAML frontmatter from Markdown text. Returns (metadata, remaining_md)."""
    m = _FRONTMATTER_RE.match(text)
    if not m:
        return Metadata(), text
    try:
        raw = yaml.safe_load(m.group(1)) or {}
    except yaml.YAMLError as e:
        raise ValueError(
            f"Ошибка в YAML-заголовке документа: {e}\n"
            "Проверьте формат YAML между маркерами ---."
        ) from e
    meta = Metadata(
        title=str(raw.get("title", "")),
        author=str(raw.get("author", "")),
        group=str(raw.get("group", "")),
        teacher=str(raw.get("teacher", "")),
        subject=str(raw.get("subject", "")),
        university=str(raw.get("university", "")),
        year=str(raw.get("year", "")),
        work_number=str(raw.get("work_number", "")),
        institute=str(raw.get("institute", "")),
        department=str(raw.get("department", "")),
        city=str(raw.get("city", "")),
    )
    body = text[m.end():]
    return meta, body


# --- AST walking ---

def _inline_to_runs(children: list) -> List[Run]:
    """Convert mistune inline AST nodes to a list of Runs."""
    runs: List[Run] = []
    if children is None:
        return runs
    for child in children:
        tp = child.get("type", "")
        if tp == "text":
            runs.append(Run(text=child.get("raw", child.get("text", ""))))
        elif tp == "codespan":
            runs.append(Run(text=child.get("raw", child.get("text", "")), code=True))
        elif tp == "strong":
            for r in _inline_to_runs(child.get("children", [])):
                r.bold = True
                runs.append(r)
        elif tp == "emphasis":
            for r in _inline_to_runs(child.get("children", [])):
                r.italic = True
                runs.append(r)
        elif tp == "link":
            url = child.get("attrs", {}).get("url", "")
            for r in _inline_to_runs(child.get("children", [])):
                runs.append(r)
            if url:
                runs.append(Run(text=f" ({url})"))
        elif tp == "image":
            alt = _extract_plain_text(child.get("children", []))
            if alt:
                runs.append(Run(text=f"[{alt}]"))
        elif tp == "softbreak" or tp == "linebreak":
            runs.append(Run(text="\n"))
        else:
            # Fallback: try to extract text
            raw = child.get("raw", child.get("text", ""))
            if raw:
                runs.append(Run(text=raw))
            elif "children" in child:
                runs.extend(_inline_to_runs(child["children"]))
    return runs


def _extract_plain_text(children: list) -> str:
    """Recursively extract plain text from inline AST nodes."""
    parts = []
    if children is None:
        return ""
    for child in children:
        tp = child.get("type", "")
        if tp == "text":
            parts.append(child.get("raw", child.get("text", "")))
        elif tp == "codespan":
            parts.append(child.get("raw", child.get("text", "")))
        elif "children" in child:
            parts.append(_extract_plain_text(child["children"]))
        elif "raw" in child:
            parts.append(child["raw"])
    return "".join(parts)


def _list_items_to_runs(items: list) -> List[List[Run]]:
    """Convert list item AST nodes to lists of Runs."""
    result = []
    for item in items:
        children = item.get("children", [])
        item_runs: List[Run] = []
        for child in children:
            tp = child.get("type", "")
            if tp == "paragraph":
                item_runs.extend(_inline_to_runs(child.get("children", [])))
            elif tp == "text":
                item_runs.append(Run(text=child.get("raw", child.get("text", ""))))
            elif "children" in child:
                item_runs.extend(_inline_to_runs(child["children"]))
        result.append(item_runs)
    return result


def _walk_ast(tokens: list) -> List[Block]:
    """Walk mistune AST tokens and produce typed blocks."""
    blocks: List[Block] = []
    for token in tokens:
        tp = token.get("type", "")

        if tp == "heading":
            text = _extract_plain_text(token.get("children", []))
            level = token.get("attrs", {}).get("level", 1)
            blocks.append(HeadingBlock(level=level, text=text))

        elif tp == "paragraph":
            children = token.get("children", [])
            # Standalone image → ImageBlock
            if len(children) == 1 and children[0].get("type") == "image":
                img = children[0]
                alt = _extract_plain_text(img.get("children", []))
                url = img.get("attrs", {}).get("url", "")
                blocks.append(ImageBlock(alt=alt, url=url))
            else:
                runs = _inline_to_runs(children)
                if runs:
                    blocks.append(ParagraphBlock(runs=runs))

        elif tp in ("code_block", "block_code"):
            info = token.get("attrs", {}).get("info", "") or ""
            raw = token.get("raw", token.get("text", ""))
            blocks.append(CodeBlock(language=info, code=raw.rstrip("\n")))

        elif tp == "list":
            ordered = token.get("attrs", {}).get("ordered", False)
            items = _list_items_to_runs(token.get("children", []))
            blocks.append(ListBlock(ordered=ordered, items=items))

        elif tp == "table":
            tbl = _parse_table(token)
            if tbl:
                blocks.append(tbl)

        elif tp == "block_quote":
            quote_children = token.get("children", [])
            runs: List[Run] = []
            for child in quote_children:
                child_tp = child.get("type", "")
                if child_tp == "paragraph":
                    if runs:
                        runs.append(Run(text="\n"))
                    runs.extend(_inline_to_runs(child.get("children", [])))
                elif "children" in child:
                    if runs:
                        runs.append(Run(text="\n"))
                    runs.extend(_inline_to_runs(child["children"]))
            if runs:
                blocks.append(BlockquoteBlock(runs=runs))

        elif tp == "thematic_break":
            pass  # skip horizontal rules

        elif tp == "blank_line":
            pass

        elif "children" in token:
            blocks.extend(_walk_ast(token["children"]))

    return blocks


def _parse_table(token: dict) -> Optional[TableBlock]:
    """Parse a mistune table token into a TableBlock."""
    children = token.get("children", [])
    headers: List[str] = []
    rows: List[List[str]] = []
    for child in children:
        tp = child.get("type", "")
        if tp == "table_head":
            # mistune v3: table_head contains table_cell directly (no row wrapper)
            for cell in child.get("children", []):
                if cell.get("type") == "table_cell":
                    headers.append(_extract_plain_text(cell.get("children", [])))
        elif tp == "table_body":
            for row in child.get("children", []):
                row_data = []
                for cell in row.get("children", []):
                    if cell.get("type") == "table_cell":
                        row_data.append(_extract_plain_text(cell.get("children", [])))
                if row_data:
                    rows.append(row_data)
    if headers or rows:
        return TableBlock(headers=headers, rows=rows)
    return None


# --- Public API ---

def parse_markdown(text: str) -> tuple[Metadata, List[Block]]:
    """Parse Markdown text into metadata and a list of typed blocks."""
    meta, body = extract_frontmatter(text)
    md = mistune.create_markdown(renderer="ast", plugins=["table", "strikethrough"])
    ast = md(body)
    blocks = _walk_ast(ast)
    return meta, blocks

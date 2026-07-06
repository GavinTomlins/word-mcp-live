"""Render a .docx as Markdown for content verification (ADR 0004).

The goal is a faithful *structural* read-back — headings, lists, tables,
inline emphasis, hyperlinks and (optionally) tracked changes — so an agent
can confirm that what it authored actually landed in the document. It is not
a lossless OOXML converter.

Tracked changes render in CriticMarkup, matching the pandoc
``--track-changes=all`` idiom: insertions as ``{++text++}``, deletions as
``{--text--}``.
"""

import re

from docx import Document
from docx.oxml.ns import qn


def _run_markdown(r_el, show_revisions: bool, in_deletion: bool) -> str:
    """Render one <w:r> as markdown text with emphasis markers."""
    text_tag = qn("w:delText") if in_deletion else qn("w:t")
    parts = []
    for child in r_el:
        if child.tag == text_tag:
            parts.append(child.text or "")
        elif child.tag == qn("w:tab"):
            parts.append("\t")
        elif child.tag in (qn("w:br"), qn("w:cr")):
            parts.append("\n")
    text = "".join(parts)
    if not text:
        return ""
    rpr = r_el.find(qn("w:rPr"))
    if rpr is not None and text.strip():
        bold = rpr.find(qn("w:b")) is not None
        italic = rpr.find(qn("w:i")) is not None
        lead, core, trail = _split_ws(text)
        if bold and italic:
            text = f"{lead}***{core}***{trail}"
        elif bold:
            text = f"{lead}**{core}**{trail}"
        elif italic:
            text = f"{lead}*{core}*{trail}"
    return text


def _split_ws(text: str) -> tuple[str, str, str]:
    core = text.strip()
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()) :]
    return lead, core, trail


def _hyperlink_target(h_el, rels) -> str | None:
    rid = h_el.get(qn("r:id"))
    if rid and rid in rels:
        return rels[rid].target_ref
    return None


def _paragraph_inline(p_el, rels, show_revisions: bool) -> str:
    """Render the inline content of a <w:p>, honouring ins/del/hyperlinks."""
    out: list[str] = []

    def emit_children(parent, in_ins: bool, in_del: bool) -> None:
        for child in parent:
            if child.tag == qn("w:r"):
                text = _run_markdown(child, show_revisions, in_del)
                if not text:
                    continue
                if in_del:
                    out.append(f"{{--{text}--}}" if show_revisions else "")
                elif in_ins:
                    out.append(f"{{++{text}++}}" if show_revisions else text)
                else:
                    out.append(text)
            elif child.tag == qn("w:hyperlink"):
                target = _hyperlink_target(child, rels)
                before = len(out)
                emit_children(child, in_ins, in_del)
                if target is not None:
                    label = "".join(out[before:])
                    del out[before:]
                    out.append(f"[{label}]({target})")
            elif child.tag == qn("w:ins"):
                emit_children(child, True, in_del)
            elif child.tag == qn("w:del"):
                emit_children(child, in_ins, True)

    emit_children(p_el, False, False)
    return "".join(out)


_HEADING_RE = re.compile(r"^heading (\d)$", re.IGNORECASE)


def _paragraph_prefix(paragraph) -> str:
    """Markdown prefix for a paragraph: heading hashes or list bullet."""
    style = (paragraph.style.name or "") if paragraph.style is not None else ""
    m = _HEADING_RE.match(style)
    if m:
        return "#" * int(m.group(1)) + " "
    if style.lower() == "title":
        return "# "
    ppr = paragraph._p.find(qn("w:pPr"))
    if ppr is not None:
        numpr = ppr.find(qn("w:numPr"))
        if numpr is not None:
            ilvl_el = numpr.find(qn("w:ilvl"))
            level = int(ilvl_el.get(qn("w:val")) or 0) if ilvl_el is not None else 0
            return "  " * level + "- "
    if "list" in style.lower():
        marker = "1. " if "number" in style.lower() else "- "
        return marker
    return ""


def _table_markdown(table, rels, show_revisions: bool) -> str:
    rows = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            cell_text = " ".join(
                _paragraph_inline(p._p, rels, show_revisions)
                for p in cell.paragraphs
            ).strip()
            cells.append(cell_text.replace("|", "\\|").replace("\n", " "))
        rows.append(cells)
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    lines = ["| " + " | ".join(r + [""] * (width - len(r))) + " |" for r in rows]
    lines.insert(1, "|" + "|".join([" --- "] * width) + "|")
    return "\n".join(lines)


def document_to_markdown(filename: str, show_revisions: bool = False) -> str:
    """Render the document body as Markdown."""
    doc = Document(filename)
    rels = doc.part.rels
    blocks: list[str] = []
    body = doc.element.body
    for child in body:
        if child.tag == qn("w:p"):
            # find the Paragraph proxy for style access
            from docx.text.paragraph import Paragraph

            paragraph = Paragraph(child, doc)
            inline = _paragraph_inline(child, rels, show_revisions)
            if not inline.strip():
                continue
            blocks.append(_paragraph_prefix(paragraph) + inline)
        elif child.tag == qn("w:tbl"):
            from docx.table import Table

            md = _table_markdown(Table(child, doc), rels, show_revisions)
            if md:
                blocks.append(md)
    return "\n\n".join(blocks) + ("\n" if blocks else "")

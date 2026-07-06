"""Business-rule validation for .docx packages (ADR 0003).

Checks the OOXML package directly for the problems that most often make a
generated document render wrong or fail to open, and reports each finding
with an actionable message (location + measured values + consequence).

This is the business-rule layer only; full XSD validation is out of scope
(see ADR 0003).
"""

import struct
import zipfile

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
XML_NS = "http://www.w3.org/XML/1998/namespace"

NS = {"w": W_NS, "wp": WP_NS, "a": A_NS, "r": R_NS, "w15": W15_NS}


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


# Ratio mismatch tolerance for table widths and image aspect ratios.
TOLERANCE = 0.05

# OOXML schema order for the child elements we check (subset: only elements
# named here are order-checked; unknown elements are ignored).
PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "numPr", "pBdr", "shd", "tabs", "spacing", "ind", "jc",
    "outlineLvl", "rPr",
]
RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "strike",
    "color", "spacing", "sz", "szCs", "highlight", "u", "vertAlign",
    "rtl",
]


def _image_dimensions(data: bytes) -> tuple[int, int] | None:
    """Read pixel dimensions from PNG/JPEG/GIF bytes without Pillow."""
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        width, height = struct.unpack(">II", data[16:24])
        return width, height
    if data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        width, height = struct.unpack("<HH", data[6:10])
        return width, height
    if data[:2] == b"\xff\xd8":  # JPEG: scan for a SOF marker
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                height, width = struct.unpack(">HH", data[i + 5 : i + 9])
                return width, height
            if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
                i += 2
                continue
            (length,) = struct.unpack(">H", data[i + 2 : i + 4])
            i += 2 + length
    return None


class _Report:
    def __init__(self) -> None:
        self.errors: list[str] = []
        self.warnings: list[str] = []

    def error(self, msg: str) -> None:
        self.errors.append(msg)

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)

    def as_dict(self) -> dict:
        return {
            "valid": not self.errors,
            "errors": self.errors,
            "warnings": self.warnings,
            "summary": (
                f"{len(self.errors)} error(s), {len(self.warnings)} warning(s)"
                if (self.errors or self.warnings)
                else "no issues found"
            ),
        }


def _check_tables(body: etree._Element, report: _Report) -> None:
    for t_idx, tbl in enumerate(body.iter(_w("tbl"))):
        children = [etree.QName(c).localname for c in tbl]
        if "tblGrid" not in children:
            report.error(
                f"TABLE[{t_idx}]: missing <w:tblGrid> — Word may refuse to "
                "open the document"
            )
            continue
        grid_cols = [
            int(gc.get(_w("w")) or 0)
            for gc in tbl.find(_w("tblGrid")).findall(_w("gridCol"))
        ]
        for r_idx, tr in enumerate(tbl.findall(_w("tr"))):
            cells = tr.findall(_w("tc"))
            col_cursor = 0
            spanned_cells = 0
            for c_idx, tc in enumerate(cells):
                span = 1
                tc_w = None
                tc_type = None
                tc_pr = tc.find(_w("tcPr"))
                if tc_pr is not None:
                    gs = tc_pr.find(_w("gridSpan"))
                    if gs is not None:
                        span = int(gs.get(_w("val")) or 1)
                    w_el = tc_pr.find(_w("tcW"))
                    if w_el is not None:
                        tc_type = w_el.get(_w("type"))
                        try:
                            tc_w = int(w_el.get(_w("w")) or 0)
                        except ValueError:
                            tc_w = None
                spanned_cells += span
                if (
                    tc_w
                    and tc_type in (None, "dxa")
                    and col_cursor + span <= len(grid_cols)
                ):
                    grid_w = sum(grid_cols[col_cursor : col_cursor + span])
                    if grid_w and abs(tc_w - grid_w) / grid_w > TOLERANCE:
                        report.warn(
                            f"TABLE[{t_idx}] row {r_idx} cell {c_idx}: "
                            f"tcW={tc_w} != gridCol sum={grid_w} "
                            "(column widths will skew)"
                        )
                col_cursor += span
            if grid_cols and spanned_cells != len(grid_cols):
                report.warn(
                    f"TABLE[{t_idx}] row {r_idx}: spans {spanned_cells} grid "
                    f"column(s) but tblGrid defines {len(grid_cols)} "
                    "(row will misalign)"
                )


def _check_images(
    doc_tree: etree._ElementTree, zf: zipfile.ZipFile, report: _Report
) -> None:
    try:
        rels_root = etree.fromstring(zf.read("word/_rels/document.xml.rels"))
    except KeyError:
        return
    rel_targets = {
        rel.get("Id"): rel.get("Target")
        for rel in rels_root.findall(f"{{{REL_NS}}}Relationship")
    }
    for i_idx, drawing in enumerate(doc_tree.iter(_w("drawing"))):
        extent = drawing.find(f".//{{{WP_NS}}}extent")
        blip = drawing.find(f".//{{{A_NS}}}blip")
        if extent is None or blip is None:
            continue
        cx, cy = int(extent.get("cx") or 0), int(extent.get("cy") or 0)
        target = rel_targets.get(blip.get(f"{{{R_NS}}}embed"))
        if not (cx and cy and target):
            continue
        member = "word/" + target.lstrip("/") if not target.startswith("word/") else target
        try:
            dims = _image_dimensions(zf.read(member))
        except KeyError:
            report.error(f"IMAGE[{i_idx}]: relationship target {target!r} missing from package")
            continue
        if not dims or not dims[1]:
            continue
        display_ratio = cx / cy
        actual_ratio = dims[0] / dims[1]
        if abs(display_ratio - actual_ratio) / actual_ratio > TOLERANCE:
            report.warn(
                f"IMAGE[{i_idx}] ({member}): display ratio "
                f"{display_ratio:.2f} != actual {actual_ratio:.2f} "
                "(image will look distorted)"
            )


def _check_comments(
    doc_tree: etree._ElementTree, zf: zipfile.ZipFile, report: _Report
) -> None:
    names = set(zf.namelist())
    anchor_ids = {
        el.get(_w("id"))
        for el in doc_tree.iter(_w("commentReference"))
        if el.get(_w("id")) is not None
    }
    defined_ids: set[str] = set()
    if "word/comments.xml" in names:
        comments_root = etree.fromstring(zf.read("word/comments.xml"))
        defined_ids = {
            c.get(_w("id"))
            for c in comments_root.findall(_w("comment"))
            if c.get(_w("id")) is not None
        }
    for cid in sorted(anchor_ids - defined_ids):
        report.error(
            f"COMMENTS: document references comment id={cid} that is not "
            "defined in comments.xml"
        )
    for cid in sorted(defined_ids - anchor_ids):
        report.warn(
            f"COMMENTS: comment id={cid} defined in comments.xml has no "
            "anchor in document.xml (comment will be invisible)"
        )
    if "word/commentsExtended.xml" in names:
        ext_root = etree.fromstring(zf.read("word/commentsExtended.xml"))
        has_replies = any(
            el.get(f"{{{W15_NS}}}paraIdParent")
            for el in ext_root.iter(f"{{{W15_NS}}}commentEx")
        )
        if has_replies and "word/commentsIds.xml" not in names:
            report.warn(
                "COMMENTS: threaded replies present but commentsIds.xml is "
                "missing (replies may not round-trip)"
            )
    elif defined_ids and "word/commentsExtended.xml" not in names:
        report.warn(
            "COMMENTS: comments.xml present without commentsExtended.xml "
            "(modern Word clients may not thread or resolve them)"
        )


def _check_fields(
    doc_tree: etree._ElementTree, zf: zipfile.ZipFile, report: _Report
) -> None:
    has_toc = any(
        "TOC" in (el.text or "") for el in doc_tree.iter(_w("instrText"))
    )
    if not has_toc:
        return
    try:
        settings_root = etree.fromstring(zf.read("word/settings.xml"))
    except KeyError:
        settings_root = None
    update_fields = (
        settings_root is not None
        and settings_root.find(_w("updateFields")) is not None
    )
    if not update_fields:
        report.warn(
            "FIELDS: document contains a TOC field but settings.xml lacks "
            "<w:updateFields/> — page numbers will be stale until manual "
            "refresh (use set_update_fields_on_open)"
        )


def _check_text_whitespace(doc_tree: etree._ElementTree, report: _Report) -> None:
    count = 0
    for t in doc_tree.iter(_w("t")):
        text = t.text or ""
        if text != text.strip() and t.get(f"{{{XML_NS}}}space") != "preserve":
            count += 1
    if count:
        report.warn(
            f"TEXT: {count} <w:t> element(s) with leading/trailing "
            'whitespace lack xml:space="preserve" (spaces will be dropped)'
        )


def _check_element_order(doc_tree: etree._ElementTree, report: _Report) -> None:
    for tag, order in ((_w("pPr"), PPR_ORDER), (_w("rPr"), RPR_ORDER)):
        bad = 0
        for props in doc_tree.iter(tag):
            positions = [
                order.index(etree.QName(c).localname)
                for c in props
                if etree.QName(c).localname in order
            ]
            if positions != sorted(positions):
                bad += 1
        if bad:
            report.warn(
                f"ORDER: {bad} <{etree.QName(tag).localname}> element(s) have "
                "children out of OOXML schema order (some Word versions "
                "ignore the misplaced properties)"
            )


def validate_docx(path: str) -> dict:
    """Run all business-rule checks against a .docx file.

    Returns a dict with ``valid`` (no errors), ``errors``, ``warnings`` and
    a one-line ``summary``.
    """
    report = _Report()
    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as e:
        report.error(f"PACKAGE: cannot open as a .docx package: {e}")
        return report.as_dict()
    with zf:
        try:
            doc_tree = etree.ElementTree(etree.fromstring(zf.read("word/document.xml")))
        except (KeyError, etree.XMLSyntaxError) as e:
            report.error(f"PACKAGE: word/document.xml unreadable: {e}")
            return report.as_dict()
        body = doc_tree.getroot().find(_w("body"))
        if body is None:
            report.error("PACKAGE: document.xml has no <w:body>")
            return report.as_dict()

        _check_tables(body, report)
        _check_images(doc_tree, zf, report)
        _check_comments(doc_tree, zf, report)
        _check_fields(doc_tree, zf, report)
        _check_text_whitespace(doc_tree, report)
        _check_element_order(doc_tree, report)
    return report.as_dict()

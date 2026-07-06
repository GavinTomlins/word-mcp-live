"""File-based comment threading: reply, resolve, delete (ADR 0007).

Word models threaded/resolvable comments across four parts:

- ``word/comments.xml``          — comment text; each comment paragraph
  carries a ``w14:paraId``
- ``word/commentsExtended.xml``  — per-paraId state: ``w15:done`` and the
  ``w15:paraIdParent`` that forms reply threads
- ``word/commentsIds.xml``       — paraId → durableId mapping
- ``word/commentsExtensible.xml``— durableId metadata

macOS Word's scripting dictionary exposes none of this, so these
operations are implemented on the closed file (zip + lxml), reusing the
anchor/relationship machinery from ``comment_writer``.
"""

import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

from word_mcp_live_gavintomlins.core.comment_writer import (
    CT_NS,
    REL_NS,
    W,
    W14,
    _generate_hex_id,
    _get_max_comment_id,
    _get_max_id_in_doc,
    _get_next_rid,
    _load_zip_part,
    _now_iso,
)
from word_mcp_live_gavintomlins.defaults import DEFAULT_AUTHOR, DEFAULT_INITIALS

W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID_NS = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX_NS = "http://schemas.microsoft.com/office/word/2018/wordml/cex"

W15 = lambda tag: f"{{{W15_NS}}}{tag}"
W16CID = lambda tag: f"{{{W16CID_NS}}}{tag}"
W16CEX = lambda tag: f"{{{W16CEX_NS}}}{tag}"

_PARTS = {
    "word/commentsExtended.xml": (
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w15:commentsEx xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        b'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
        b'mc:Ignorable="w15"/>',
    ),
    "word/commentsIds.xml": (
        "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w16cid:commentsIds xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        b'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" '
        b'mc:Ignorable="w16cid"/>',
    ),
    "word/commentsExtensible.xml": (
        "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w16cex:commentsExtensible xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        b'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" '
        b'mc:Ignorable="w16cex"/>',
    ),
}


class _Package:
    """Loaded .docx with lazily-parsed parts and a modified-part registry."""

    def __init__(self, filepath: str):
        self.path = Path(filepath)
        self.zip_bytes = self.path.read_bytes()
        self._trees: dict[str, etree._Element] = {}
        self.modified: set[str] = set()
        self.created: set[str] = set()

    def part(self, name: str, create: bool = False) -> etree._Element | None:
        if name in self._trees:
            return self._trees[name]
        raw = _load_zip_part(self.zip_bytes, name)
        if raw is None:
            if not create:
                return None
            raw = _PARTS[name][2]
            self.created.add(name)
            self.modified.add(name)
        self._trees[name] = etree.fromstring(raw)
        return self._trees[name]

    def touch(self, name: str) -> None:
        self.modified.add(name)

    def ensure_part_registered(self, name: str) -> None:
        """Make sure rels + content types declare a comments side-part."""
        rel_type, content_type, _ = _PARTS[name]
        rels = self.part("word/_rels/document.xml.rels")
        if rels is not None:
            target = name.split("/", 1)[1]
            if not any(
                rel.get("Type") == rel_type
                for rel in rels.iter(f"{{{REL_NS}}}Relationship")
            ):
                rel = etree.SubElement(rels, f"{{{REL_NS}}}Relationship")
                rel.set("Id", f"rId{_get_next_rid(rels)}")
                rel.set("Type", rel_type)
                rel.set("Target", target)
                self.touch("word/_rels/document.xml.rels")
        ct = self.part("[Content_Types].xml")
        if ct is not None:
            part_name = f"/{name}"
            if not any(
                o.get("PartName") == part_name
                for o in ct.iter(f"{{{CT_NS}}}Override")
            ):
                override = etree.SubElement(ct, f"{{{CT_NS}}}Override")
                override.set("PartName", part_name)
                override.set("ContentType", content_type)
                self.touch("[Content_Types].xml")

    def save(self) -> None:
        serialized = {
            name: etree.tostring(
                tree, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            for name, tree in self._trees.items()
            if name in self.modified
        }
        buffer = BytesIO()
        with zipfile.ZipFile(BytesIO(self.zip_bytes), "r") as zf_in:
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf_out:
                written = set()
                for item in zf_in.infolist():
                    if item.filename in serialized:
                        zf_out.writestr(item, serialized[item.filename])
                        written.add(item.filename)
                    else:
                        zf_out.writestr(item, zf_in.read(item.filename))
                for name, data in serialized.items():
                    if name not in written:
                        zf_out.writestr(name, data)
        self.path.write_bytes(buffer.getvalue())


def _find_comment(pkg: _Package, comment_id: int) -> etree._Element | None:
    comments = pkg.part("word/comments.xml")
    if comments is None:
        return None
    for comment in comments.iter(W("comment")):
        if comment.get(W("id")) == str(comment_id):
            return comment
    return None


def _comment_para_id(pkg: _Package, comment: etree._Element) -> str:
    """paraId of the comment's last paragraph, generating one if absent."""
    paras = comment.findall(W("p"))
    if not paras:
        para = etree.SubElement(comment, W("p"))
        paras = [para]
    last = paras[-1]
    para_id = last.get(W14("paraId"))
    if not para_id:
        para_id = _generate_hex_id()
        last.set(W14("paraId"), para_id)
        pkg.touch("word/comments.xml")
    return para_id


def _ensure_extended_entry(
    pkg: _Package, para_id: str, parent_para_id: str | None = None
) -> etree._Element:
    ext = pkg.part("word/commentsExtended.xml", create=True)
    for entry in ext.iter(W15("commentEx")):
        if entry.get(W15("paraId")) == para_id:
            return entry
    entry = etree.SubElement(ext, W15("commentEx"))
    entry.set(W15("paraId"), para_id)
    if parent_para_id:
        entry.set(W15("paraIdParent"), parent_para_id)
    entry.set(W15("done"), "0")
    pkg.touch("word/commentsExtended.xml")
    return entry


def _ensure_durable_ids(pkg: _Package, para_id: str) -> str:
    ids = pkg.part("word/commentsIds.xml", create=True)
    for entry in ids.iter(W16CID("commentId")):
        if entry.get(W16CID("paraId")) == para_id:
            return entry.get(W16CID("durableId"))
    durable = _generate_hex_id()
    entry = etree.SubElement(ids, W16CID("commentId"))
    entry.set(W16CID("paraId"), para_id)
    entry.set(W16CID("durableId"), durable)
    pkg.touch("word/commentsIds.xml")

    ext = pkg.part("word/commentsExtensible.xml", create=True)
    x_entry = etree.SubElement(ext, W16CEX("commentExtensible"))
    x_entry.set(W16CEX("durableId"), durable)
    x_entry.set(W16CEX("dateUtc"), _now_iso())
    pkg.touch("word/commentsExtensible.xml")
    return durable


def _register_side_parts(pkg: _Package) -> None:
    for name in _PARTS:
        pkg.ensure_part_registered(name)


def reply_to_comment(
    filepath: str,
    comment_id: int,
    reply_text: str,
    author: str = DEFAULT_AUTHOR,
    initials: str = DEFAULT_INITIALS,
) -> dict:
    """Add a threaded reply to an existing comment (file must be closed)."""
    pkg = _Package(filepath)
    parent = _find_comment(pkg, comment_id)
    if parent is None:
        return {"success": False, "error": f"Comment id={comment_id} not found"}
    parent_para_id = _comment_para_id(pkg, parent)

    comments = pkg.part("word/comments.xml")
    doc = pkg.part("word/document.xml")
    if doc is None:
        return {"success": False, "error": "Cannot read word/document.xml"}

    reply_id = max(_get_max_comment_id(comments), _get_max_id_in_doc(doc)) + 1
    reply_para_id = _generate_hex_id()

    reply = etree.SubElement(comments, W("comment"))
    reply.set(W("id"), str(reply_id))
    reply.set(W("author"), author)
    reply.set(W("date"), _now_iso())
    reply.set(W("initials"), initials)
    rp = etree.SubElement(reply, W("p"))
    rp.set(W14("paraId"), reply_para_id)
    rp.set(W14("textId"), "77777777")
    ann_run = etree.SubElement(rp, W("r"))
    ann_rpr = etree.SubElement(ann_run, W("rPr"))
    ann_style = etree.SubElement(ann_rpr, W("rStyle"))
    ann_style.set(W("val"), "CommentReference")
    etree.SubElement(ann_run, W("annotationRef"))
    text_run = etree.SubElement(rp, W("r"))
    t = etree.SubElement(text_run, W("t"))
    t.text = reply_text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    pkg.touch("word/comments.xml")

    # Anchor the reply at the same range as the parent.
    parent_id = str(comment_id)
    anchors = {"start": None, "end": None, "ref": None}
    for el in doc.iter():
        if not isinstance(el.tag, str):
            continue
        tag = etree.QName(el).localname
        if tag == "commentRangeStart" and el.get(W("id")) == parent_id:
            anchors["start"] = el
        elif tag == "commentRangeEnd" and el.get(W("id")) == parent_id:
            anchors["end"] = el
        elif tag == "commentReference" and el.get(W("id")) == parent_id:
            anchors["ref"] = el.getparent()  # the run
    if any(v is None for v in anchors.values()):
        return {
            "success": False,
            "error": f"Comment id={comment_id} has incomplete anchors in "
                     "document.xml; cannot attach a reply",
        }
    for key, tag in (("start", "commentRangeStart"), ("end", "commentRangeEnd")):
        el = anchors[key]
        new_el = etree.Element(W(tag))
        new_el.set(W("id"), str(reply_id))
        el.addnext(new_el)
    ref_run = etree.Element(W("r"))
    ref_rpr = etree.SubElement(ref_run, W("rPr"))
    ref_style = etree.SubElement(ref_rpr, W("rStyle"))
    ref_style.set(W("val"), "CommentReference")
    ref = etree.SubElement(ref_run, W("commentReference"))
    ref.set(W("id"), str(reply_id))
    anchors["ref"].addnext(ref_run)
    pkg.touch("word/document.xml")

    _ensure_extended_entry(pkg, parent_para_id)
    _ensure_durable_ids(pkg, parent_para_id)
    _ensure_extended_entry(pkg, reply_para_id, parent_para_id=parent_para_id)
    _ensure_durable_ids(pkg, reply_para_id)
    _register_side_parts(pkg)
    pkg.save()
    return {
        "success": True,
        "reply_id": reply_id,
        "parent_id": comment_id,
        "author": author,
        "message": f"Added reply #{reply_id} to comment #{comment_id}",
    }


def resolve_comment(filepath: str, comment_id: int, resolved: bool = True) -> dict:
    """Mark a comment thread resolved (done) or reopen it (file closed)."""
    pkg = _Package(filepath)
    comment = _find_comment(pkg, comment_id)
    if comment is None:
        return {"success": False, "error": f"Comment id={comment_id} not found"}
    para_id = _comment_para_id(pkg, comment)
    entry = _ensure_extended_entry(pkg, para_id)
    entry.set(W15("done"), "1" if resolved else "0")
    pkg.touch("word/commentsExtended.xml")
    _ensure_durable_ids(pkg, para_id)
    _register_side_parts(pkg)
    pkg.save()
    state = "resolved" if resolved else "reopened"
    return {
        "success": True,
        "comment_id": comment_id,
        "resolved": resolved,
        "message": f"Comment #{comment_id} {state}",
    }


def delete_comment(filepath: str, comment_id: int) -> dict:
    """Delete a comment and its replies from every part (file closed)."""
    pkg = _Package(filepath)
    comment = _find_comment(pkg, comment_id)
    if comment is None:
        return {"success": False, "error": f"Comment id={comment_id} not found"}
    comments = pkg.part("word/comments.xml")
    doc = pkg.part("word/document.xml")

    target_para_id = comment.findall(W("p"))[-1].get(W14("paraId")) if comment.findall(W("p")) else None

    # Collect reply paraIds via commentsExtended parentage.
    doomed_para_ids = {target_para_id} if target_para_id else set()
    ext = pkg.part("word/commentsExtended.xml")
    if ext is not None and target_para_id:
        changed = True
        while changed:
            changed = False
            for entry in ext.iter(W15("commentEx")):
                pid = entry.get(W15("paraId"))
                parent = entry.get(W15("paraIdParent"))
                if parent in doomed_para_ids and pid not in doomed_para_ids:
                    doomed_para_ids.add(pid)
                    changed = True

    # Map doomed paraIds back to comment ids (plus the target id itself).
    doomed_ids = {str(comment_id)}
    for c in comments.iter(W("comment")):
        paras = c.findall(W("p"))
        pid = paras[-1].get(W14("paraId")) if paras else None
        if pid and pid in doomed_para_ids:
            doomed_ids.add(c.get(W("id")))

    removed = 0
    for c in list(comments.iter(W("comment"))):
        if c.get(W("id")) in doomed_ids:
            c.getparent().remove(c)
            removed += 1
    pkg.touch("word/comments.xml")

    if doc is not None:
        for el in list(doc.iter()):
            if not isinstance(el.tag, str):
                continue
            tag = etree.QName(el).localname
            if tag in ("commentRangeStart", "commentRangeEnd") and el.get(W("id")) in doomed_ids:
                el.getparent().remove(el)
            elif tag == "commentReference" and el.get(W("id")) in doomed_ids:
                run = el.getparent()
                run.getparent().remove(run)
        pkg.touch("word/document.xml")

    durable_ids = set()
    ids_part = pkg.part("word/commentsIds.xml")
    if ids_part is not None:
        for entry in list(ids_part.iter(W16CID("commentId"))):
            if entry.get(W16CID("paraId")) in doomed_para_ids:
                durable_ids.add(entry.get(W16CID("durableId")))
                entry.getparent().remove(entry)
        pkg.touch("word/commentsIds.xml")
    if ext is not None:
        for entry in list(ext.iter(W15("commentEx"))):
            if entry.get(W15("paraId")) in doomed_para_ids:
                entry.getparent().remove(entry)
        pkg.touch("word/commentsExtended.xml")
    extensible = pkg.part("word/commentsExtensible.xml")
    if extensible is not None and durable_ids:
        for entry in list(extensible.iter(W16CEX("commentExtensible"))):
            if entry.get(W16CEX("durableId")) in durable_ids:
                entry.getparent().remove(entry)
        pkg.touch("word/commentsExtensible.xml")

    pkg.save()
    return {
        "success": True,
        "deleted_ids": sorted(int(i) for i in doomed_ids),
        "message": f"Deleted comment #{comment_id} and its replies "
                   f"({removed} comment element(s) removed)",
    }


def list_comments(filepath: str) -> list[dict]:
    """List all comments with thread structure and resolution state.

    Returns entries shaped like the legacy reader (id, comment_id, author,
    initials, date, text, …) plus ``resolved``, ``parent_id`` and
    ``reply_count`` derived from commentsExtended.xml.
    """
    pkg = _Package(filepath)
    comments = pkg.part("word/comments.xml")
    if comments is None:
        return []

    ext_by_para: dict[str, dict] = {}
    ext = pkg.part("word/commentsExtended.xml")
    if ext is not None:
        for entry in ext.iter(W15("commentEx")):
            ext_by_para[entry.get(W15("paraId"))] = {
                "done": entry.get(W15("done")) == "1",
                "parent_para": entry.get(W15("paraIdParent")),
            }

    para_to_cid: dict[str, str] = {}
    records = []
    for idx, comment in enumerate(comments.iter(W("comment"))):
        cid = comment.get(W("id"), str(idx))
        paras = comment.findall(W("p"))
        para_id = paras[-1].get(W14("paraId")) if paras else None
        if para_id:
            para_to_cid[para_id] = cid
        text = "".join(
            t.text or ""
            for t in comment.iter(W("t"))
        )
        state = ext_by_para.get(para_id, {})
        records.append({
            "id": f"comment_{idx + 1}",
            "comment_id": cid,
            "author": comment.get(W("author"), "Unknown"),
            "initials": comment.get(W("initials"), ""),
            "date": comment.get(W("date")),
            "text": text.strip(),
            "resolved": state.get("done", False),
            "parent_para": state.get("parent_para"),
            "paragraph_index": None,
            "in_table": False,
            "reference_text": "",
        })

    for record in records:
        parent_para = record.pop("parent_para")
        record["parent_id"] = para_to_cid.get(parent_para) if parent_para else None
    for record in records:
        record["reply_count"] = sum(
            1 for r in records if r["parent_id"] == record["comment_id"]
        )
    return records

"""Small, local DOCX inspection and exact-edit MCP server."""

from __future__ import annotations

import base64
import hashlib
import json
import os
import posixpath
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any
from urllib.parse import unquote

from mcp.server.fastmcp import FastMCP


mcp = FastMCP("word-enhanced")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
W = "{" + NS["w"] + "}"
WP = "{" + NS["wp"] + "}"
A = "{" + NS["a"] + "}"
MAX_PART_BYTES = 250 * 1024 * 1024
MAX_PACKAGE_BYTES = 1024 * 1024 * 1024


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _qname(name: str) -> str:
    if name.startswith("{"):
        uri, local = name[1:].split("}", 1)
        for prefix, known in NS.items():
            if uri == known:
                return f"{prefix}:{local}"
        return f"{{{uri}}}{local}"
    return name


def _xml_part(name: str) -> bool:
    return name.endswith((".xml", ".rels"))


def _package(path: str) -> dict[str, bytes]:
    source = Path(path).expanduser()
    if not source.is_file():
        raise ValueError(f"DOCX does not exist: {source}")
    if source.suffix.lower() != ".docx":
        raise ValueError("Expected a .docx file")
    try:
        with zipfile.ZipFile(source) as archive:
            infos = archive.infolist()
            names = [item.filename for item in infos]
            if len(names) != len(set(names)):
                raise ValueError("Duplicate ZIP part names")
            if sum(item.file_size for item in infos) > MAX_PACKAGE_BYTES:
                raise ValueError("DOCX is too large for this tool")
            if any(item.file_size > MAX_PART_BYTES for item in infos):
                raise ValueError("A DOCX part is too large for this tool")
            if archive.testzip() is not None:
                raise ValueError("DOCX ZIP CRC check failed")
            parts = {name: archive.read(name) for name in names}
    except (zipfile.BadZipFile, OSError) as exc:
        raise ValueError(f"Invalid DOCX ZIP: {exc}") from exc
    if "[Content_Types].xml" not in parts or "word/document.xml" not in parts:
        raise ValueError("Missing required DOCX parts")
    for name, data in parts.items():
        if _xml_part(name):
            try:
                ET.fromstring(data)
            except ET.ParseError as exc:
                raise ValueError(f"Invalid XML in {name}: {exc}") from exc
    return parts


def _opc_errors(parts: dict[str, bytes]) -> list[str]:
    """Check the package relationships and content-type coverage we can prove."""
    errors: list[str] = []
    if "_rels/.rels" not in parts:
        errors.append("Missing root relationships")
    for name in parts:
        if name.startswith("/") or name == ".." or "/../" in f"/{name}/":
            errors.append(f"Invalid part name: {name}")
    content = ET.fromstring(parts["[Content_Types].xml"])
    if content.tag != "{" + NS["ct"] + "}Types":
        errors.append("Invalid content types root")
    defaults = {item.get("Extension", "").lower() for item in content.findall("ct:Default", NS)}
    overrides = {item.get("PartName", "").lstrip("/") for item in content.findall("ct:Override", NS)}
    if "word/document.xml" not in overrides:
        errors.append("Missing main document content type override")
    if "_rels/.rels" in parts:
        root_rels = ET.fromstring(parts["_rels/.rels"])
        if not any((item.get("Type") or "").endswith("/officeDocument") for item in root_rels):
            errors.append("Missing root officeDocument relationship")
    for name in parts:
        if name == "[Content_Types].xml":
            continue
        extension = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if name not in overrides and extension not in defaults:
            errors.append(f"Missing content type: {name}")
    for rel_name, data in parts.items():
        if not rel_name.endswith(".rels"):
            continue
        root = ET.fromstring(data)
        if root.tag != "{" + NS["pr"] + "}Relationships":
            errors.append(f"Invalid relationships root: {rel_name}")
        base = "" if rel_name == "_rels/.rels" else rel_name.rsplit("/_rels/", 1)[0] + "/"
        if rel_name != "_rels/.rels" and "/_rels/" not in rel_name:
            errors.append(f"Unexpected relationships path: {rel_name}")
        for rel in root:
            if rel.get("TargetMode") == "External":
                continue
            target = unquote((rel.get("Target") or "").split("#", 1)[0])
            resolved = posixpath.normpath(("" if target.startswith("/") else base) + target.lstrip("/"))
            if resolved not in parts:
                errors.append(f"Broken relationship: {rel_name} {rel.get('Id')} -> {target}")
    return errors


def _tree(node: ET.Element, unordered: bool = False) -> Any:
    children = [_tree(child) for child in node]
    if unordered:
        children.sort(key=lambda item: json.dumps(item, ensure_ascii=False, sort_keys=True))
    text = node.text or ""
    if len(node) and not text.strip():
        text = ""
    tail = node.tail or ""
    if not tail.strip():
        tail = ""
    return (_qname(node.tag), tuple(sorted((_qname(k), v) for k, v in node.attrib.items())), text, tail, children)


def _normalized(data: bytes, name: str) -> Any:
    root = ET.fromstring(data)
    return _tree(root, name == "[Content_Types].xml" or name.endswith(".rels"))


def _digest_tree(data: bytes, name: str) -> str:
    return _sha(json.dumps(_normalized(data, name), ensure_ascii=False, sort_keys=True).encode())


def _xml_diffs(left: ET.Element, right: ET.Element, path: str | None = None) -> list[dict[str, Any]]:
    here = path or f"/{_qname(left.tag)}[1]"
    changes: list[dict[str, Any]] = []
    if left.tag != right.tag:
        return [{"path": here, "before": _qname(left.tag), "after": _qname(right.tag)}]
    for key in sorted(set(left.attrib) | set(right.attrib)):
        a, b = left.get(key), right.get(key)
        if a != b:
            changes.append({"path": f"{here}/@{_qname(key)}", "before": a, "after": b})
    if (left.text or "") != (right.text or ""):
        if not (len(left) and len(right) and not (left.text or "").strip() and not (right.text or "").strip()):
            changes.append({"path": f"{here}/#text", "before": left.text, "after": right.text})

    def indexed(node: ET.Element) -> dict[str, ET.Element]:
        counts: dict[str, int] = {}
        result: dict[str, ET.Element] = {}
        for child in node:
            tag = _qname(child.tag)
            counts[tag] = counts.get(tag, 0) + 1
            result[f"{tag}[{counts[tag]}]"] = child
        return result

    old, new = indexed(left), indexed(right)
    for key in dict.fromkeys([*old, *new]):
        child_path = f"{here}/{key}"
        if key not in old:
            changes.append({"path": child_path, "before": None, "after": "<element added>"})
        elif key not in new:
            changes.append({"path": child_path, "before": "<element removed>", "after": None})
        else:
            changes.extend(_xml_diffs(old[key], new[key], child_path))
    # Child order is meaningful for WordprocessingML even if children match by name.
    if list(old) != list(new):
        changes.append({"path": f"{here}/#child-order", "before": list(old), "after": list(new)})
    return changes


def _semantic_diffs(before: bytes, after: bytes, name: str) -> list[dict[str, Any]]:
    if _normalized(before, name) == _normalized(after, name):
        return []
    if name == "[Content_Types].xml" or name.endswith(".rels"):
        # Declaration and relationship order has no meaning. Normalize the
        # collection first, then report changed entries rather than order.
        old = _normalized(before, name)
        new = _normalized(after, name)
        return [{"path": f"/{_qname(ET.fromstring(before).tag)}[1]", "before": old, "after": new}]
    return _xml_diffs(ET.fromstring(before), ET.fromstring(after))


def _word_objects(parts: dict[str, bytes]) -> dict[str, Any]:
    document = ET.fromstring(parts["word/document.xml"])
    body = document.find("w:body", NS)
    paragraphs = len(body.findall("w:p", NS)) if body is not None else 0
    tables = len(body.findall("w:tbl", NS)) if body is not None else 0
    bookmarks: list[dict[str, str | None]] = []
    fields: list[dict[str, str | None]] = []
    hyperlinks: list[dict[str, str | None]] = []
    drawings: list[dict[str, Any]] = []
    for part_name, data in parts.items():
        if not (part_name == "word/document.xml" or part_name.startswith(("word/header", "word/footer", "word/footnotes", "word/endnotes")) and part_name.endswith(".xml")):
            continue
        root = ET.fromstring(data)
        for element in root.iter():
            if element.tag in (W + "bookmarkStart", W + "bookmarkEnd"):
                bookmarks.append({"part": part_name, "kind": _qname(element.tag), "name": element.get(W + "name"), "id": element.get(W + "id")})
            elif element.tag in (W + "instrText", W + "fldChar", W + "fldSimple"):
                fields.append({"part": part_name, "kind": _qname(element.tag), "value": element.text if element.tag == W + "instrText" else element.get(W + "fldCharType") or element.get(W + "instr")})
            elif element.tag == W + "hyperlink":
                hyperlinks.append({"part": part_name, "anchor": element.get(W + "anchor"), "id": element.get("{" + NS["r"] + "}id")})
        for kind in ("anchor", "inline"):
            for index, drawing in enumerate(root.findall(f".//wp:{kind}", NS), 1):
                ph = drawing.find("wp:positionH/wp:posOffset", NS)
                pv = drawing.find("wp:positionV/wp:posOffset", NS)
                extent = drawing.find("wp:extent", NS)
                effect = drawing.find("wp:effectExtent", NS)
                blip = drawing.find(".//a:blip", NS)
                wrap = next((child for child in drawing if child.tag.startswith(WP + "wrap")), None)
                drawings.append({
                    "part": part_name, "kind": kind, "index": index,
                    "positionH": ph.text if ph is not None else None,
                    "positionV": pv.text if pv is not None else None,
                    "extent": dict(extent.attrib) if extent is not None else None,
                    "effectExtent": dict(effect.attrib) if effect is not None else None,
                    "editId": drawing.get("{" + NS["wp14"] + "}editId"),
                    "anchorId": drawing.get("{" + NS["wp14"] + "}anchorId"),
                    "wrap": _qname(wrap.tag) if wrap is not None else None,
                    "embed": blip.get("{" + NS["r"] + "}embed") if blip is not None else None,
                    "semantic_hash": _digest_tree(ET.tostring(drawing), "drawing.xml"),
                })
    relationships = {}
    for name, data in parts.items():
        if name.endswith(".rels"):
            root = ET.fromstring(data)
            relationships[name] = sorted((item.get("Id"), item.get("Type"), item.get("Target"), item.get("TargetMode")) for item in root)
    styles = ET.fromstring(parts["word/styles.xml"]) if "word/styles.xml" in parts else None
    sections = document.findall(".//w:sectPr", NS)
    return {
        "paragraphs": paragraphs, "tables": tables,
        "bookmarks": bookmarks, "fields": fields, "hyperlinks": hyperlinks,
        "drawings": drawings, "relationships": relationships,
        "style_count": len(styles.findall("w:style", NS)) if styles is not None else 0,
        "styles_hash": _digest_tree(parts["word/styles.xml"], "word/styles.xml") if styles is not None else None,
        "numbering_hash": _digest_tree(parts["word/numbering.xml"], "word/numbering.xml") if "word/numbering.xml" in parts else None,
        "section_count": len(sections),
        "media_hashes": {n: _sha(v) for n, v in parts.items() if n.startswith("word/media/")},
        "header_footer_hashes": {n: _digest_tree(v, n) for n, v in parts.items() if n.startswith(("word/header", "word/footer")) and n.endswith(".xml")},
    }


def inspect_docx_impl(path: str) -> dict[str, Any]:
    parts = _package(path)
    objects = _word_objects(parts)
    opc_errors = _opc_errors(parts)
    return {
        "valid_zip": True, "basic_opc_valid": not opc_errors, "opc_errors": opc_errors,
        "part_count": len(parts), "parts": sorted(parts),
        "document_xml_exists": True, "paragraph_count": objects["paragraphs"],
        "table_count": objects["tables"], "style_count": objects["style_count"],
        "has_styles": "word/styles.xml" in parts,
        "has_numbering": "word/numbering.xml" in parts,
        "relationship_parts": sorted(objects["relationships"]),
        "relationship_count": sum(map(len, objects["relationships"].values())),
        "relationships": {name: [{"id": rid, "type": kind, "target": target, "mode": mode} for rid, kind, target, mode in items] for name, items in objects["relationships"].items()},
        "media_parts": sorted(objects["media_hashes"]),
        "header_footer_parts": sorted(objects["header_footer_hashes"]),
        "bookmark_start_count": sum(b["kind"] == "w:bookmarkStart" for b in objects["bookmarks"]),
        "bookmark_end_count": sum(b["kind"] == "w:bookmarkEnd" for b in objects["bookmarks"]),
        "bookmark_names": sorted({b["name"] for b in objects["bookmarks"] if b["name"]}),
        "field_node_count": len(objects["fields"]),
        "field_instructions": [f["value"] for f in objects["fields"] if f["kind"] == "w:instrText"],
        "hyperlink_count": len(objects["hyperlinks"]),
        "anchor_count": sum(d["kind"] == "anchor" for d in objects["drawings"]),
        "inline_count": sum(d["kind"] == "inline" for d in objects["drawings"]),
        "section_count": objects["section_count"],
        "drawings": objects["drawings"],
    }


def _compare_parts(before: dict[str, bytes], after: dict[str, bytes]) -> tuple[dict[str, Any], dict[str, list[dict[str, Any]]]]:
    added = sorted(after.keys() - before.keys())
    removed = sorted(before.keys() - after.keys())
    changed = sorted(n for n in before.keys() & after.keys() if before[n] != after[n])
    xml_diffs = {n: _semantic_diffs(before[n], after[n], n) for n in changed if _xml_part(n)}
    normalized_changed = sorted(n for n, changes in xml_diffs.items() if changes)
    return ({
        "added_parts": added, "removed_parts": removed, "raw_changed_parts": changed,
        "normalized_xml_changed_parts": normalized_changed,
        "raw_only_xml_parts": sorted(n for n in xml_diffs if not xml_diffs[n]),
        "binary_changed_parts": sorted(n for n in changed if not _xml_part(n)),
    }, xml_diffs)


def compare_docx_impl(before_path: str, after_path: str) -> dict[str, Any]:
    before, after = _package(before_path), _package(after_path)
    package, xml_diffs = _compare_parts(before, after)
    old, new = _word_objects(before), _word_objects(after)
    old_names = {b["name"] for b in old["bookmarks"] if b["kind"] == "w:bookmarkStart"}
    new_names = {b["name"] for b in new["bookmarks"] if b["kind"] == "w:bookmarkStart"}
    old_bookmark_ids = {b["name"]: b["id"] for b in old["bookmarks"] if b["kind"] == "w:bookmarkStart" and b["name"]}
    new_bookmark_ids = {b["name"]: b["id"] for b in new["bookmarks"] if b["kind"] == "w:bookmarkStart" and b["name"]}
    rel_changes = {name: {"before": old["relationships"].get(name), "after": new["relationships"].get(name)} for name in old["relationships"].keys() | new["relationships"].keys() if old["relationships"].get(name) != new["relationships"].get(name)}
    drawing_changes = []
    for a, b in zip(old["drawings"], new["drawings"]):
        diff = {key: {"before": a.get(key), "after": b.get(key)} for key in a.keys() | b.keys() if a.get(key) != b.get(key)}
        if diff:
            drawing_changes.append({"part": a["part"], "kind": a["kind"], "index": a["index"], "changes": diff})
    return {
        "valid_before": True, "valid_after": True,
        "basic_opc_before": not _opc_errors(before), "basic_opc_after": not _opc_errors(after),
        "package": package,
        "xml": {"changes_by_part": {n: {"count": len(v), "sample": v[:80]} for n, v in xml_diffs.items() if v}},
        "objects": {
            "bookmarks": {"before_starts": sum(b["kind"] == "w:bookmarkStart" for b in old["bookmarks"]), "after_starts": sum(b["kind"] == "w:bookmarkStart" for b in new["bookmarks"]), "removed_names": sorted(old_names - new_names), "added_names": sorted(new_names - old_names), "changed_ids": {name: {"before": old_bookmark_ids[name], "after": new_bookmark_ids[name]} for name in old_bookmark_ids.keys() & new_bookmark_ids.keys() if old_bookmark_ids[name] != new_bookmark_ids[name]}, "sequence_unchanged": old["bookmarks"] == new["bookmarks"]},
            "fields": {"before_nodes": len(old["fields"]), "after_nodes": len(new["fields"]), "sequence_unchanged": old["fields"] == new["fields"]},
            "hyperlinks": {"before": len(old["hyperlinks"]), "after": len(new["hyperlinks"]), "sequence_unchanged": old["hyperlinks"] == new["hyperlinks"]},
            "relationships": {"unchanged": old["relationships"] == new["relationships"], "before_count": sum(map(len, old["relationships"].values())), "after_count": sum(map(len, new["relationships"].values())), "changed_parts": rel_changes},
            "styles": {"before_count": old["style_count"], "after_count": new["style_count"], "unchanged": old["styles_hash"] == new["styles_hash"]},
            "numbering": {"unchanged": old["numbering_hash"] == new["numbering_hash"]},
            "headers_footers": {"unchanged": old["header_footer_hashes"] == new["header_footer_hashes"], "removed": sorted(old["header_footer_hashes"].keys() - new["header_footer_hashes"].keys()), "added": sorted(new["header_footer_hashes"].keys() - old["header_footer_hashes"].keys()), "changed": sorted(n for n in old["header_footer_hashes"].keys() & new["header_footer_hashes"].keys() if old["header_footer_hashes"][n] != new["header_footer_hashes"][n])},
            "drawings": {"before": len(old["drawings"]), "after": len(new["drawings"]), "changes": drawing_changes},
            "media": {"unchanged": old["media_hashes"] == new["media_hashes"], "changed_or_removed": sorted(n for n, h in old["media_hashes"].items() if new["media_hashes"].get(n) != h), "before_hashes": old["media_hashes"], "after_hashes": new["media_hashes"]},
            "sections": {"before": old["section_count"], "after": new["section_count"]},
        },
        "note": "Raw part bytes, normalized XML structure, and visual appearance are separate checks. This tool does not render or open Word.",
    }


def _fragment(text: str | None, b64: str | None, label: str) -> bytes:
    if (text is None) == (b64 is None):
        raise ValueError(f"Specify exactly one of {label} or {label}_b64")
    if b64 is not None:
        try:
            return base64.b64decode(b64, validate=True)
        except ValueError as exc:
            raise ValueError(f"Invalid base64 for {label}") from exc
    return text.encode("utf-8")


def patch_docx_part_exact_impl(input_path: str, output_path: str, part_name: str, old_fragment: str | None = None, new_fragment: str | None = None, expected_count: int = 1, old_fragment_b64: str | None = None, new_fragment_b64: str | None = None) -> dict[str, Any]:
    source, target = Path(input_path).expanduser().resolve(), Path(output_path).expanduser().absolute()
    if source == target or target.exists():
        raise ValueError("Output must be a new path distinct from input")
    if target.suffix.lower() != ".docx":
        raise ValueError("Output must have .docx extension")
    if not _xml_part(part_name):
        raise ValueError("Exact patch is limited to XML and relationship parts")
    if not isinstance(expected_count, int) or expected_count < 1:
        raise ValueError("expected_count must be a positive integer")
    old = _fragment(old_fragment, old_fragment_b64, "old_fragment")
    new = _fragment(new_fragment, new_fragment_b64, "new_fragment")
    if not old or old == new:
        raise ValueError("Old fragment must be nonempty and differ from new fragment")
    parts = _package(str(source))
    source_hash = _sha(source.read_bytes())
    opc_errors = _opc_errors(parts)
    if opc_errors:
        raise ValueError(f"Input has OPC errors: {opc_errors[:5]}")
    if part_name not in parts:
        raise ValueError(f"Part not found: {part_name}")
    count = parts[part_name].count(old)
    if count != expected_count:
        raise ValueError(f"Expected {expected_count} matches in {part_name}; found {count}")
    patched = parts[part_name].replace(old, new)
    ET.fromstring(patched)  # Fail before writing if XML would be malformed.
    target.parent.mkdir(parents=True, exist_ok=True)
    temp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(dir=target.parent, prefix=".word-enhanced-", suffix=".docx", delete=False) as temp:
            temp_path = temp.name
        with zipfile.ZipFile(source) as original, zipfile.ZipFile(temp_path, "w") as output:
            output.comment = original.comment
            for info in original.infolist():
                output.writestr(info, patched if info.filename == part_name else parts[info.filename])
        result = _package(temp_path)
        output_opc_errors = _opc_errors(result)
        if output_opc_errors:
            raise ValueError(f"Output has OPC errors: {output_opc_errors[:5]}")
        if any(result[n] != data for n, data in parts.items() if n != part_name):
            raise ValueError("A non-target uncompressed part changed")
        if result[part_name] != patched:
            raise ValueError("Target part verification failed")
        if _sha(source.read_bytes()) != source_hash:
            raise ValueError("Input changed during patch; output was not published")
        os.link(temp_path, target)  # Fails if another process created the target.
    finally:
        if temp_path is not None:
            Path(temp_path).unlink(missing_ok=True)
    return {"output_path": str(target), "part_name": part_name, "replacements": count, "zip_valid": True, "non_target_uncompressed_parts_unchanged": True, "input_sha256": source_hash, "output_sha256": _sha(target.read_bytes())}


def _permitted_path(path: str, allowed: list[str]) -> bool:
    return any(path == prefix or path.startswith(prefix + "/") for prefix in allowed)


def validate_docx_edit_impl(before_path: str, after_path: str, allowlist: dict[str, Any]) -> dict[str, Any]:
    before, after = _package(before_path), _package(after_path)
    package, xml_diffs = _compare_parts(before, after)
    reasons: list[str] = []
    for error in _opc_errors(after):
        reasons.append(f"Output OPC error: {error}")
    expected_exact = allowlist.get("exact_replacements", [])
    expected_values = allowlist.get("expected_xml_values", [])
    allowed_paths = allowlist.get("allowed_xml_paths", {})
    if not expected_exact and not expected_values:
        reasons.append("No target assertion supplied")
    for name in package["added_parts"]:
        if name not in allowlist.get("allowed_added_parts", []):
            reasons.append(f"Unexpected added part: {name}")
    for name in package["removed_parts"]:
        if name not in allowlist.get("allowed_removed_parts", []):
            reasons.append(f"Unexpected removed part: {name}")
    exact_parts: set[str] = set()
    for rule in expected_exact:
        name = rule["part"]
        if name not in before or name not in after:
            reasons.append(f"Exact replacement part missing: {name}")
            continue
        old = _fragment(rule.get("old_fragment"), rule.get("old_fragment_b64"), "old_fragment")
        new = _fragment(rule.get("new_fragment"), rule.get("new_fragment_b64"), "new_fragment")
        count = rule.get("expected_count", 1)
        if not old or old == new or not isinstance(count, int) or count < 1 or before[name].count(old) != count:
            reasons.append(f"Original fragment count differs from expected in {name}")
        elif after[name] != before[name].replace(old, new):
            reasons.append(f"Actual output differs from exact replacement in {name}")
        exact_parts.add(name)
    all_xml_changes = [(name, change) for name, changes in xml_diffs.items() for change in changes]
    unexpected_xml: dict[str, list[str]] = {}
    for name, change in all_xml_changes:
        if name not in exact_parts and not _permitted_path(change["path"], allowed_paths.get(name, [])):
            unexpected_xml.setdefault(name, []).append(change["path"])
    for name, paths in unexpected_xml.items():
        reasons.append(f"Unexpected XML changes: {name} ({len(paths)} paths; first: {paths[0]})")
    for rule in expected_values:
        found = any(name == rule.get("part") and change["path"] == rule.get("path") and change["before"] == rule.get("before") and change["after"] == rule.get("after") for name, change in all_xml_changes)
        if not found:
            reasons.append(f"Expected XML value change missing: {rule.get('part')} {rule.get('path')}")
    for name in package["binary_changed_parts"]:
        if name not in allowlist.get("allowed_binary_changed_parts", []):
            reasons.append(f"Unexpected binary change: {name}")
    if not allowlist.get("allow_raw_reserialization", False):
        for name in package["raw_only_xml_parts"]:
            if name not in allowlist.get("allowed_raw_changed_parts", []):
                reasons.append(f"Unexpected raw-only XML change: {name}")
    return {
        "status": "PASS" if not reasons else "FAIL", "pass": not reasons,
        "target_completed": bool(expected_exact or expected_values) and not any("Expected" in item or "Original fragment" in item or "No target" in item or "Actual output differs" in item for item in reasons),
        "non_target_changes": reasons[:100], "non_target_change_count": len(reasons),
        "unexpected_xml_changes": {name: {"count": len(paths), "sample_paths": paths[:20]} for name, paths in unexpected_xml.items()},
        "package": package,
        "risk_notes": ["Structural PASS does not verify visual layout or Word-native behavior."] + (["Raw XML was reserialized; normalized structure was compared."] if package["raw_only_xml_parts"] else []),
    }


@mcp.tool()
def inspect_docx(path: str) -> dict[str, Any]:
    """Read-only ZIP/OPC and key Word object summary for an existing DOCX."""
    return inspect_docx_impl(path)


@mcp.tool()
def compare_docx(before_path: str, after_path: str) -> dict[str, Any]:
    """Compare raw parts, normalized XML, and critical Word objects read-only."""
    return compare_docx_impl(before_path, after_path)


@mcp.tool()
def patch_docx_part_exact(input_path: str, output_path: str, part_name: str, old_fragment: str | None = None, new_fragment: str | None = None, expected_count: int = 1, old_fragment_b64: str | None = None, new_fragment_b64: str | None = None) -> dict[str, Any]:
    """Replace exact bytes in one XML part of a NEW DOCX; UTF-8 or base64 fragments."""
    return patch_docx_part_exact_impl(input_path, output_path, part_name, old_fragment, new_fragment, expected_count, old_fragment_b64, new_fragment_b64)


@mcp.tool()
def validate_docx_edit(before_path: str, after_path: str, allowlist: dict[str, Any]) -> dict[str, Any]:
    """Fail unless a target assertion is met and every raw/semantic change is allowed. Allowlist keys: exact_replacements, expected_xml_values, allowed_xml_paths, allowed_added_parts, allowed_removed_parts, allowed_binary_changed_parts, allowed_raw_changed_parts, allow_raw_reserialization."""
    return validate_docx_edit_impl(before_path, after_path, allowlist)


if __name__ == "__main__":
    mcp.run(transport="stdio")

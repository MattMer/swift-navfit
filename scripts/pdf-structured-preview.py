"""Read-only structured preview for the three supplied XFA form revisions.

No PDF/XFA JavaScript runs. Flattened, revised or ambiguous forms are rejected.
This development adapter does not yet commit drafts in Access.
"""
import argparse
import hashlib
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import pikepdf

ROOT = Path(__file__).resolve().parents[1]
SOURCES = {
    "EVAL": "E1-E6 EVAL BLANK NAVPERS 1616-26 05-2025_Final.pdf",
    "CHIEFEVAL": "CHIEF EVAL BLANK 1616-27 05-2025_Final.pdf",
    "FITREP": "W2-O6 FITREP NAVPERS 1610-2 05-2025_Final.pdf",
}

def packets(pdf):
    acro = pdf.Root.get("/AcroForm")
    if acro is None or "/XFA" not in acro:
        raise ValueError("No structured XFA data; flattened PDFs require a separate workflow")
    xfa = acro["/XFA"]
    if not isinstance(xfa, pikepdf.Array) or len(xfa) % 2:
        raise ValueError("Unsupported XFA packet structure")
    result = {}
    for index in range(0, len(xfa), 2):
        name = str(xfa[index])
        if name in result:
            raise ValueError("Duplicate XFA packet")
        raw = xfa[index + 1].read_bytes()
        if len(raw) > 4_000_000:
            raise ValueError("XFA packet exceeds preview limit")
        result[name] = raw
    return result

def preview(path):
    path = Path(path).resolve()
    if path.stat().st_size > 30_000_000:
        raise ValueError("PDF exceeds preview limit")
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    with pikepdf.open(path) as pdf:
        if not pdf.allow.extract:
            raise ValueError("PDF does not permit extraction")
        data = packets(pdf)
    template_hash = hashlib.sha256(data["template"]).hexdigest()
    matches = []
    for kind, filename in SOURCES.items():
        with pikepdf.open(ROOT / filename) as source:
            if hashlib.sha256(packets(source)["template"]).hexdigest() == template_hash:
                matches.append(kind)
    if len(matches) != 1:
        raise ValueError("Unsupported or ambiguous form revision")
    kind = matches[0]
    raw = data["datasets"]
    if b"<!DOCTYPE" in raw.upper() or b"<!ENTITY" in raw.upper():
        raise ValueError("XML declarations are not permitted")
    root = ET.fromstring(raw)
    mapping = json.loads((ROOT / "src/pdf-field-map.json").read_text(encoding="utf-8"))[kind]
    values = {}
    for row in mapping:
        source_key, target_key = row["source"], row["target"]
        fields = [node for node in root.iter() if node.tag.split("}")[-1] == source_key]
        if len(fields) > 1:
            raise ValueError(f"Ambiguous repeated field: {source_key}")
        # Missing data is explicitly distinguished from an existing empty field.
        values[target_key] = {"source_field": source_key, "conversion": row["conversion"], "present": bool(fields),
                              "value": (fields[0].text or "") if fields else None}
    retained = []
    def retain(node, prefix):
        current = prefix + "/" + node.tag.split("}")[-1]
        if not len(node):
            retained.append({"path": current, "value": node.text or ""})
        for child in node:
            retain(child, current)
    retain(root, "")
    if before != hashlib.sha256(path.read_bytes()).hexdigest():
        raise ValueError("Source changed during preview")
    return {"status": "PreviewOnly", "kind": kind, "source_sha256": before,
            "template_sha256": template_hash, "fields": values, "retained_raw": retained,
            "requires_review": ["Dates retain source text and require explicit century interpretation",
                                "Identifiers use SSN by default; legacy DoDID requires explicit selection",
                                "Blank/missing source values do not erase existing draft values",
                                "No draft values have been committed"]}

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    result = preview(args.source)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        with args.output.open("x", encoding="utf-8") as output:
            output.write(text + "\n")
    else:
        print(text)

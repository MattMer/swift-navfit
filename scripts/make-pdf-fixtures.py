"""Fill supplied XFA datasets with synthetic values; preserve interactive form packets."""
from pathlib import Path
import hashlib
import json
import xml.etree.ElementTree as ET
import pikepdf

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output" / "pdf"
SOURCES = {
    "EVAL": "E1-E6 EVAL BLANK NAVPERS 1616-26 05-2025_Final.pdf",
    "CHIEFEVAL": "CHIEF EVAL BLANK 1616-27 05-2025_Final.pdf",
    "FITREP": "W2-O6 FITREP NAVPERS 1610-2 05-2025_Final.pdf",
}

def packets(pdf):
    xfa = pdf.Root["/AcroForm"]["/XFA"]
    return {str(xfa[i]): (i + 1, xfa[i + 1].read_bytes()) for i in range(0, len(xfa), 2)}

def create():
    OUTPUT.mkdir(parents=True, exist_ok=True)
    manifest = []
    for kind, filename in SOURCES.items():
        source = ROOT / filename
        before = hashlib.sha256(source.read_bytes()).hexdigest()
        target = OUTPUT / f"synthetic-{kind.lower()}.pdf"
        if target.exists():
            raise FileExistsError(target)
        with pikepdf.open(source) as pdf:
            if not pdf.allow.modify_form:
                raise ValueError("Source does not permit form filling")
            data = packets(pdf)
            template = ET.fromstring(data["template"][1])
            available = {f.attrib.get("name") for f in template.iter() if f.tag.endswith("}field")}
            values = {"a": f"EXAMPLE, SYNTHETIC {kind}", "b": {"EVAL": "YN2", "CHIEFEVAL": "YNC", "FITREP": "LCDR"}[kind],
                      "c": "1110" if kind == "FITREP" else "", "d": "000000000",
                      "g": "SYNTHETIC UNIT", "k": "260101", "l": "260930",
                      "H": f"SYNTHETIC {kind} TEST COMMENTS. No real person or command. Test punctuation: O'Brien & example."}
            if kind == "CHIEFEVAL":
                values = {"NAMEmaster": "EXAMPLE, SYNTHETIC CHIEF", "RATINGmaster": "YNC", "DESIGNATIONmaster": "",
                          "SSNDODIDmaster": "000000000", "SHIPSTATION": "SYNTHETIC UNIT", "FromDate": "260101",
                          "ToDate": "260930", "REPORTINGSENIORCOMMENTS": "SYNTHETIC CHIEF TEST COMMENTS. No real person or command. O'Brien & example."}
            if not values.keys() <= available:
                raise ValueError(f"Unexpected template names: {values.keys() - available}")
            datasets = ET.fromstring(data["datasets"][1])
            container = datasets.find(".//topmostSubform")
            if kind == "CHIEFEVAL":
                container = datasets.find(".//form1")
            if container is None:
                raise ValueError("Unsupported XFA data root")
            for name, value in values.items():
                field = container.find(".//" + name)
                if field is None:
                    field = ET.SubElement(container, name)
                field.text = value
            pdf.Root["/AcroForm"]["/XFA"][data["datasets"][0]] = pdf.make_stream(ET.tostring(datasets, encoding="utf-8"))
            pdf.save(target, encryption=True if pdf.is_encrypted else None)
        with pikepdf.open(target) as written:
            reopened = ET.fromstring(packets(written)["datasets"][1])
            for name, value in values.items():
                assert (reopened.find(".//" + name).text or "") == value
            assert "template" in packets(written)
        assert before == hashlib.sha256(source.read_bytes()).hexdigest()
        manifest.append({"kind": kind, "source": filename, "source_sha256": before,
                         "output": target.name, "output_sha256": hashlib.sha256(target.read_bytes()).hexdigest(),
                         "expected": values, "rendering": "XFA-aware viewer required; static page appearance is not regenerated"})
    (OUTPUT / "synthetic-manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print("PASS: three interactive XFA datasets verified; source PDFs unchanged")

if __name__ == "__main__":
    create()

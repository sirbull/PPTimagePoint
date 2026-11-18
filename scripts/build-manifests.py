"""Lager manifest.xml (GitHub Pages) og manifest.local.xml (https://localhost:3000) fra malen."""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
tpl = (ROOT / "manifest.template.xml").read_text(encoding="utf-8")

TARGETS = {
    "manifest.xml": dict(ID="9bfd5010-5050-433e-9166-cedb07dc32f8", NAME="PPTimagePoint",
                         BASE="https://sirbull.github.io/PPTimagePoint", ORIGIN="https://sirbull.github.io"),
    "manifest.local.xml": dict(ID="9bfd5010-5050-433e-9166-cedb07dc32f9", NAME="PPTimagePoint (lokal)",
                               BASE="https://localhost:3000", ORIGIN="https://localhost:3000"),
}
for name, v in TARGETS.items():
    out = tpl
    for k, val in v.items():
        out = out.replace(f"__{k}__", val)
    (ROOT / name).write_text(out, encoding="utf-8")
    print("skrev", name)

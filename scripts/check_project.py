#!/usr/bin/env python3
"""Validate the BRD catalog and registered HTML paths."""
import json
from pathlib import Path
root = Path(__file__).resolve().parent.parent
catalog = json.loads((root / "web/catalog.json").read_text(encoding="utf-8"))
assert catalog.get("schemaVersion") == 1
assert isinstance(catalog.get("documents"), list)
seen = set()
for document in catalog["documents"]:
    identifier = document["id"]
    assert identifier not in seen, f"Duplicate BRD ID: {identifier}"
    seen.add(identifier)
    assert document.get("title") and document.get("version")
    assert document.get("illustrations")
    illustration_ids = set()
    for item in document["illustrations"]:
        assert item["id"] not in illustration_ids, f"Duplicate illustration ID: {item['id']}"
        illustration_ids.add(item["id"])
        path = item["path"]
        assert path.startswith("/illustrations/") and path.endswith(".html")
        assert ".." not in path and "//" not in path
        target = root / "web" / path.lstrip("/")
        assert target.is_file(), f"Missing HTML: {target}"
        print("OK:", target.relative_to(root))
print(f"Validated {len(seen)} BRD(s)")

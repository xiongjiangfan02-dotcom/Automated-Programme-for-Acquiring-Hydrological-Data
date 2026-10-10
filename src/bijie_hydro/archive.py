from __future__ import annotations

import json
import hashlib
from pathlib import Path
from typing import Any


class ArchiveStore:
    def __init__(self, root: Path, batch_id: str):
        self.batch_dir = root / batch_id[:8] / batch_id
        self.batch_dir.mkdir(parents=True, exist_ok=False)

    def save_bytes(self, name: str, value: bytes) -> Path:
        path = self.batch_dir / name
        path.write_bytes(value)
        return path

    def save_json(self, name: str, value: Any) -> Path:
        path = self.batch_dir / name
        path.write_text(json.dumps(value, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        return path

    def finalize_manifest(self) -> None:
        items = []
        for path in sorted(self.batch_dir.iterdir()):
            if path.is_file() and path.name != "manifest.json":
                body = path.read_bytes()
                items.append({"file": path.name, "bytes": len(body), "sha256": hashlib.sha256(body).hexdigest()})
        self.save_json("manifest.json", {"algorithm": "SHA-256", "files": items})


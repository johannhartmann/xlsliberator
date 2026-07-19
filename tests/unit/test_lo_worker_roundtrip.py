from __future__ import annotations

from pathlib import Path
from typing import Any

from xlsliberator import lo_worker


class _FakeUno:
    @staticmethod
    def systemPathToFileUrl(path: str) -> str:
        return f"file://{path}"


class _FakeDocument:
    def __init__(self) -> None:
        self.url = ""
        self.properties: tuple[Any, ...] = ()

    def storeToURL(self, url: str, properties: tuple[Any, ...]) -> None:
        self.url = url
        self.properties = properties
        Path(url.removeprefix("file://")).write_bytes(b"round-trip")


def test_store_ods_roundtrip_copy_exports_fresh_calc8_file(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    destination = tmp_path / "target.ods"
    destination.write_bytes(b"original")
    document = _FakeDocument()
    monkeypatch.setattr(lo_worker, "_property_value", lambda name, value: (name, value))

    roundtrip = lo_worker._store_ods_roundtrip_copy(
        {"uno": _FakeUno()},
        document,
        destination,
    )

    assert roundtrip == tmp_path / ".target.xlsliberator-roundtrip.ods"
    assert roundtrip.read_bytes() == b"round-trip"
    assert destination.read_bytes() == b"original"
    assert document.properties == (("FilterName", "calc8"), ("Overwrite", True))

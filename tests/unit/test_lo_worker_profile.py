from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree

from xlsliberator.lo_worker import _initialize_secure_office_profile


def test_secure_office_profile_disables_in_process_python(tmp_path: Path) -> None:
    registry_path = _initialize_secure_office_profile(tmp_path)

    assert registry_path == tmp_path / "user" / "registrymodifications.xcu"
    root = ElementTree.parse(registry_path).getroot()
    namespace = {"oor": "http://openoffice.org/2001/registry"}
    item = root.find(
        "item[@oor:path='/org.openoffice.Office.Common/Security/Scripting']",
        namespace,
    )
    assert item is not None
    setting = item.find("prop[@oor:name='DisablePythonRuntime']", namespace)
    assert setting is not None
    assert setting.findtext("value") == "true"

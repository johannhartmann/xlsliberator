"""Deterministic ODS control and event inventory."""

from __future__ import annotations

import itertools

# Used only for element types; all untrusted parsing uses defusedxml.
import xml.etree.ElementTree as ET  # nosec B405
import zipfile
from pathlib import Path
from typing import Any

from defusedxml.ElementTree import fromstring as safe_fromstring
from loguru import logger

from xlsliberator.validation_models import ControlIR, EventBindingIR, SourceRef, TargetRef

NS = {
    "form": "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "script": "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
}

# Form-namespace elements that are structural wrappers rather than controls.
_CONTROL_WRAPPER_LOCALS = {
    "form",
    "forms",
    "properties",
    "property",
    "event-listener",
    "events",
}


def extract_controls_from_ods(ods_path: Path) -> list[ControlIR]:
    """Extract form controls from ODS content.xml."""
    root = _read_content_xml(ods_path)
    if root is None:
        return []
    return _controls_from_root(root, ods_path)


def extract_event_bindings_from_ods(ods_path: Path) -> list[EventBindingIR]:
    """Extract event listener bindings from ODS content.xml."""
    root = _read_content_xml(ods_path)
    if root is None:
        return []
    return _bindings_from_root(root, ods_path)


def extract_controls_and_bindings_from_ods(
    ods_path: Path,
) -> tuple[list[ControlIR], list[EventBindingIR]]:
    """Parse content.xml once and return both controls and event bindings."""
    root = _read_content_xml(ods_path)
    if root is None:
        return [], []
    return _controls_from_root(root, ods_path), _bindings_from_root(root, ods_path)


def _controls_from_root(root: ET.Element, ods_path: Path) -> list[ControlIR]:
    controls: list[ControlIR] = []
    for element, sheet_name in _walk_with_sheet(root):
        namespace, local_name = _split_tag(element.tag)
        if namespace != NS["form"] or local_name in _CONTROL_WRAPPER_LOCALS:
            continue

        attrs = _clean_attrs(element.attrib)
        control_name = attrs.get("name") or attrs.get("id") or f"control-{len(controls) + 1}"
        control_id = attrs.get("id") or control_name
        controls.append(
            ControlIR(
                id=control_id,
                name=control_name,
                control_type=local_name,
                sheet=sheet_name,
                properties=attrs,
                linked_cell=attrs.get("linked-cell"),
                list_fill_range=attrs.get("list-source"),
                source_ref=SourceRef(
                    source_file=str(ods_path),
                    sheet=sheet_name,
                    artifact_type="control",
                    artifact_id=control_id,
                ),
                target_ref=TargetRef(
                    target_file=str(ods_path),
                    sheet=sheet_name,
                    artifact_type="control",
                    artifact_id=control_id,
                ),
            )
        )

    return controls


def _bindings_from_root(root: ET.Element, ods_path: Path) -> list[EventBindingIR]:
    bindings: list[EventBindingIR] = []
    counter = itertools.count(1)
    for control_element, element, sheet_name in _walk_event_listeners(root):
        attrs = _clean_attrs(element.attrib)
        control_attrs = _clean_attrs(control_element.attrib) if control_element is not None else {}
        control_id = control_attrs.get("id") or control_attrs.get("name")
        event_name = attrs.get("event-name") or attrs.get("listener-event") or "unknown"
        script_uri = attrs.get("href") or attrs.get("script") or ""
        binding_id = f"event-{next(counter)}"
        source_ref = SourceRef(
            source_file=str(ods_path),
            sheet=sheet_name,
            artifact_type="event_binding",
            artifact_id=binding_id,
        )
        bindings.append(
            EventBindingIR(
                id=binding_id,
                source_ref=source_ref,
                event_name=event_name,
                source_handler=script_uri,
                target_script_uri=script_uri if "language=Python" in script_uri else None,
                control_id=control_id,
                details={"attributes": attrs, "control_attributes": control_attrs},
            )
        )

    return bindings


def _read_content_xml(ods_path: Path) -> ET.Element | None:
    try:
        with zipfile.ZipFile(ods_path, "r") as archive:
            content = archive.read("content.xml")
        return safe_fromstring(content)
    except Exception as exc:
        logger.warning(f"Could not parse ODS controls from {ods_path}: {exc}")
        return None


def _walk_with_sheet(root: ET.Element) -> list[tuple[ET.Element, str | None]]:
    results: list[tuple[ET.Element, str | None]] = []

    def visit(element: ET.Element, sheet_name: str | None) -> None:
        namespace, local_name = _split_tag(element.tag)
        current_sheet = sheet_name
        if namespace == NS["table"] and local_name == "table":
            current_sheet = element.attrib.get(f"{{{NS['table']}}}name", sheet_name)
        results.append((element, current_sheet))
        for child in element:
            visit(child, current_sheet)

    visit(root, None)
    return results


def _walk_event_listeners(
    root: ET.Element,
) -> list[tuple[ET.Element | None, ET.Element, str | None]]:
    results: list[tuple[ET.Element | None, ET.Element, str | None]] = []

    def visit(element: ET.Element, sheet_name: str | None, control: ET.Element | None) -> None:
        namespace, local_name = _split_tag(element.tag)
        current_sheet = sheet_name
        if namespace == NS["table"] and local_name == "table":
            current_sheet = element.attrib.get(f"{{{NS['table']}}}name", sheet_name)

        # Track the nearest enclosing form control so listeners nested under an
        # <office:event-listeners> wrapper still resolve to their owning control,
        # not the wrapper element (which carries no id/name).
        current_control = control
        if namespace == NS["form"] and local_name not in _CONTROL_WRAPPER_LOCALS:
            current_control = element

        for child in element:
            child_namespace, child_local = _split_tag(child.tag)
            if child_namespace in {NS["form"], NS["script"]} and child_local == "event-listener":
                results.append((current_control, child, current_sheet))
            visit(child, current_sheet, current_control)

    visit(root, None, None)
    return results


def _split_tag(tag: str) -> tuple[str | None, str]:
    if tag.startswith("{"):
        namespace, local = tag[1:].split("}", 1)
        return namespace, local
    return None, tag


def _clean_attrs(attrs: dict[str, Any]) -> dict[str, str]:
    cleaned: dict[str, str] = {}
    for key, value in attrs.items():
        _namespace, local = _split_tag(key)
        cleaned[local] = str(value)
    return cleaned

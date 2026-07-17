"""Tests for fail-closed VBA translation and provenance."""

import json
from pathlib import Path

from xlsliberator.translation_service import (
    FakeTranslationProvider,
    TranslationProviderError,
    TranslationProviderUnavailable,
    TranslationService,
    TranslationStatus,
)
from xlsliberator.vba_execution import DifferentialProof
from xlsliberator.vba_parser import parse_vba_project

VBA = """Public Sub UpdateCell()
    Range("A1").Value = 1
End Sub
"""

VALID_PYTHON = """def UpdateCell(*args):
    return None

g_exportedScripts = (UpdateCell,)
"""


def _service(
    provider: FakeTranslationProvider,
    *,
    cache_path: Path | None = None,
    mapping_paths: tuple[Path, ...] = (),
    runtime: bool = True,
    evidence_dir: Path | None = None,
) -> TranslationService:
    return TranslationService(
        provider,
        cache_path=cache_path,
        mapping_paths=mapping_paths,
        target_runtime_image_id="sha256:runtime" if runtime else None,
        target_package_manifest=[{"name": "libobasis26.2-pyuno", "version": "26.2.4.2-2"}]
        if runtime
        else [],
        evidence_dir=evidence_dir,
    )


def _translate_proven(service: TranslationService, modules: dict[str, str]):  # type: ignore[no-untyped-def]
    normalized = {
        name: "\n".join(line.rstrip() for line in source.replace("\r\n", "\n").split("\n")).strip()
        for name, source in modules.items()
    }
    project = parse_vba_project("VBAProject", normalized)
    proofs = [
        DifferentialProof(
            procedure_id=procedure.procedure_id,
            source_trace_id=f"excel:{procedure.procedure_id}",
            target_trace_id=f"libreoffice:{procedure.procedure_id}",
            equivalent=True,
        )
        for module in project.modules
        for procedure in module.procedures
    ]
    return service.translate_project(modules, differential_proofs=proofs)


def test_provider_failure_is_failed_and_not_cached(tmp_path: Path) -> None:
    cache = tmp_path / "cache.json"
    provider = FakeTranslationProvider(error=TranslationProviderError("API failed"))

    result = _service(provider, cache_path=cache).translate_project({"Module1": VBA})

    assert result.status is TranslationStatus.FAILED
    assert not result.accepted
    assert not cache.exists()
    assert result.modules["Module1"].python_code is None


def test_unavailable_provider_is_not_translation_success(tmp_path: Path) -> None:
    provider = FakeTranslationProvider(error=TranslationProviderUnavailable("no credentials"))

    result = _service(provider, cache_path=tmp_path / "cache.json").translate_project(
        {"Module1": VBA}
    )

    assert result.status is TranslationStatus.UNAVAILABLE
    assert not result.accepted


def test_invalid_or_forbidden_output_is_rejected_and_not_cached(tmp_path: Path) -> None:
    cache = tmp_path / "cache.json"
    invalid = FakeTranslationProvider(default_response="def UpdateCell(:\n    pass")
    forbidden = FakeTranslationProvider(
        default_response=(
            "import subprocess\n\n"
            "def UpdateCell(*args):\n    return None\n\n"
            "g_exportedScripts = (UpdateCell,)\n"
        )
    )

    syntax_result = _service(invalid, cache_path=cache).translate_project({"Module1": VBA})
    import_result = _service(forbidden, cache_path=cache).translate_project({"Module1": VBA})

    assert syntax_result.status is TranslationStatus.REJECTED
    assert import_result.status is TranslationStatus.REJECTED
    assert import_result.modules["Module1"].validation.forbidden_imports == ["subprocess"]
    assert not cache.exists()


def test_missing_procedure_or_export_rejects_candidate() -> None:
    provider = FakeTranslationProvider(
        default_response="def Different(*args):\n    pass\n\ng_exportedScripts = ()\n"
    )

    result = _service(provider).translate_project({"Module1": VBA})

    validation = result.modules["Module1"].validation
    assert result.status is TranslationStatus.REJECTED
    assert validation.unresolved_source_procedures == ["UpdateCell"]
    assert validation.missing_exports == ["UpdateCell"]


def test_cache_key_invalidates_on_model_and_mapping_changes(tmp_path: Path) -> None:
    cache = tmp_path / "cache.json"
    mapping = tmp_path / "mapping.yaml"
    mapping.write_text("version: 1\n")
    first = FakeTranslationProvider(default_response=VALID_PYTHON, model_id="model-a")
    first_result = _translate_proven(
        _service(first, cache_path=cache, mapping_paths=(mapping,)), {"Module1": VBA}
    )
    assert first_result.accepted
    assert len(first.calls) == 1

    same = FakeTranslationProvider(default_response=VALID_PYTHON, model_id="model-a")
    cached = _translate_proven(
        _service(same, cache_path=cache, mapping_paths=(mapping,)), {"Module1": VBA}
    )
    assert cached.accepted
    assert same.calls == []
    assert cached.modules["Module1"].cache_hit

    changed_model = FakeTranslationProvider(default_response=VALID_PYTHON, model_id="model-b")
    assert _translate_proven(
        _service(changed_model, cache_path=cache, mapping_paths=(mapping,)),
        {"Module1": VBA},
    ).accepted
    assert len(changed_model.calls) == 1

    mapping.write_text("version: 2\n")
    changed_mapping = FakeTranslationProvider(default_response=VALID_PYTHON, model_id="model-a")
    assert _translate_proven(
        _service(changed_mapping, cache_path=cache, mapping_paths=(mapping,)),
        {"Module1": VBA},
    ).accepted
    assert len(changed_mapping.calls) == 1


def test_project_validation_rejects_unresolved_cross_module_reference() -> None:
    source_b = "Public Sub Existing()\nEnd Sub"
    provider = FakeTranslationProvider(
        responses={
            'module="ModuleA"': (
                "import ModuleB\n\n"
                "def UpdateCell(*args):\n    ModuleB.Missing()\n\n"
                "g_exportedScripts = (UpdateCell,)\n"
            ),
            'module="ModuleB"': (
                "def Existing(*args):\n    return None\n\ng_exportedScripts = (Existing,)\n"
            ),
        }
    )

    result = _service(provider).translate_project({"ModuleA": VBA, "ModuleB": source_b})

    assert result.status is TranslationStatus.REJECTED
    assert result.modules["ModuleA"].validation.unresolved_cross_module_references == [
        "ModuleB.Missing"
    ]


def test_runtime_unavailability_makes_valid_candidate_partial() -> None:
    provider = FakeTranslationProvider(default_response=VALID_PYTHON)

    result = _service(provider, runtime=False).translate_project({"Module1": VBA})

    assert result.status is TranslationStatus.PARTIAL
    assert result.runtime_status is TranslationStatus.UNAVAILABLE
    assert not result.accepted


def test_evidence_contains_prompt_hashes_validation_and_no_secret(tmp_path: Path) -> None:
    evidence = tmp_path / "evidence"
    provider = FakeTranslationProvider(default_response=VALID_PYTHON)

    result = _translate_proven(_service(provider, evidence_dir=evidence), {"Module1": VBA})

    assert result.accepted
    manifest = json.loads((evidence / "translation.json").read_text())
    module = manifest["modules"]["Module1"]
    assert "<SOURCE_DATA" in manifest["prompts"]["Module1"]
    assert len(module["response_hash"]) == 64
    assert module["validation"]["syntax_valid"] is True
    assert module["source_map"]["UpdateCell"].startswith("vba-node:")
    assert module["execution_strategies"]["UpdateCell"] == "translate_python"
    assert manifest["execution_plan"]["fully_executable"] is True
    assert manifest["secrets_stored"] is False

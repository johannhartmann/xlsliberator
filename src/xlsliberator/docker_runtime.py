"""Exclusive Docker boundary for LibreOffice and PyUNO execution."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import time
import uuid
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from xlsliberator.execution_sandbox import (
    ExecutionKind,
    SandboxJob,
    SandboxMount,
    SandboxPolicy,
    WorkspacePathPolicy,
    docker_sandbox_arguments,
)
from xlsliberator.workbook_security import validate_untrusted_workbook

LIBREOFFICE_VERSION = "26.2.4.2"
BASE_IMAGE_DIGEST = "sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df"
OFFICE_PROGRAM_PREFIX = "/opt/libreoffice26.2/program/"
DEFAULT_IMAGE = "xlsliberator-libreoffice:26.2.4.2"
STOCK_RUNTIME_VARIANT = "stock"
DEFAULT_TIMEOUT_SECONDS = 120


class DockerRuntimeUnavailable(RuntimeError):
    """Raised when the authoritative LibreOffice Docker runtime is unavailable."""


class DockerRuntimeTimeout(DockerRuntimeUnavailable):
    """Raised when a disposable runtime job exceeds its wall-time limit."""


class MalformedWorkerResponse(DockerRuntimeUnavailable):
    """Raised when the container response violates the JSON protocol."""


@dataclass(frozen=True)
class DockerRuntimeIdentity:
    """Immutable identity resolved before a runtime request."""

    image_reference: str
    image_id: str
    version: str
    runtime_variant: str = STOCK_RUNTIME_VARIANT
    architecture: str | None = None
    probe: dict[str, Any] = field(default_factory=dict)


class LibreOfficeDockerRuntime:
    """Run exactly one request in one disposable LibreOffice container."""

    def __init__(
        self,
        image: str | None = None,
        *,
        timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
        expected_variant: str | None = None,
        docker_executable: str = "docker",
        workspace_roots: list[Path] | tuple[Path, ...] | None = None,
        sandbox_policy: SandboxPolicy | None = None,
    ) -> None:
        self.image = image or os.environ.get("XLSLIBERATOR_LIBREOFFICE_IMAGE", DEFAULT_IMAGE)
        self.timeout_seconds = timeout_seconds
        self.expected_variant = expected_variant or os.environ.get(
            "XLSLIBERATOR_RUNTIME_VARIANT", STOCK_RUNTIME_VARIANT
        )
        self.docker_executable = docker_executable
        self.workspace_paths = WorkspacePathPolicy(workspace_roots)
        self.sandbox_policy = sandbox_policy or SandboxPolicy()

    def resolve_identity(self, *, probe: bool = True) -> DockerRuntimeIdentity:
        """Resolve the configured tag to an immutable local image ID."""
        self._require_docker()
        inspect = self._run_docker_cli(
            [
                self.docker_executable,
                "image",
                "inspect",
                self.image,
                "--format",
                "{{json .}}",
            ],
            timeout_seconds=15,
        )
        try:
            data = json.loads(inspect.stdout)
            image_id = str(data["Id"])
            architecture = str(data.get("Architecture") or "") or None
            labels = dict((data.get("Config") or {}).get("Labels") or {})
        except (json.JSONDecodeError, KeyError, TypeError) as exc:
            raise DockerRuntimeUnavailable("Docker returned malformed image identity data") from exc
        if not image_id.startswith("sha256:"):
            raise DockerRuntimeUnavailable("LibreOffice image did not resolve to an immutable ID")
        labelled_version = labels.get("org.xlsliberator.libreoffice.version")
        if labelled_version != LIBREOFFICE_VERSION:
            raise DockerRuntimeUnavailable(
                "LibreOffice image version mismatch: "
                f"expected {LIBREOFFICE_VERSION}, got {labelled_version or 'unlabelled'}"
            )
        labelled_variant = labels.get("org.xlsliberator.runtime.variant")
        if labelled_variant != self.expected_variant:
            raise DockerRuntimeUnavailable(
                "LibreOffice runtime variant mismatch: "
                f"expected {self.expected_variant}, got {labelled_variant or 'unlabelled'}"
            )
        probe_data: dict[str, Any] = {}
        if probe:
            response = self.request({"op": "runtime_probe"}, _identity=image_id)
            if not response.get("success"):
                raise DockerRuntimeUnavailable(
                    str((response.get("error") or {}).get("message") or "runtime probe failed")
                )
            probe_data = dict(response.get("data") or {})
            if probe_data.get("libreoffice_build") != LIBREOFFICE_VERSION:
                raise DockerRuntimeUnavailable("Runtime probe returned the wrong LibreOffice build")
            if not probe_data.get("uno_importable"):
                raise DockerRuntimeUnavailable("Runtime probe could not import matching PyUNO")
            self._validate_probe_provenance(
                probe_data, architecture, expected_variant=self.expected_variant
            )
        return DockerRuntimeIdentity(
            image_reference=self.image,
            image_id=image_id,
            version=LIBREOFFICE_VERSION,
            runtime_variant=self.expected_variant,
            architecture=architecture,
            probe=probe_data,
        )

    @staticmethod
    def _validate_probe_provenance(
        probe_data: dict[str, Any],
        image_architecture: str | None,
        *,
        expected_variant: str = STOCK_RUNTIME_VARIANT,
    ) -> None:
        """Reject a probe that cannot prove one coherent office/PyUNO runtime."""
        if probe_data.get("base_image_digest") != BASE_IMAGE_DIGEST:
            raise DockerRuntimeUnavailable("Runtime probe returned the wrong base-image digest")
        if probe_data.get("runtime_variant") != expected_variant:
            raise DockerRuntimeUnavailable("Runtime probe returned the wrong runtime variant")
        program_prefix = str(probe_data.get("office_program_prefix") or "")
        expected_prefix = (
            OFFICE_PROGRAM_PREFIX
            if expected_variant == STOCK_RUNTIME_VARIANT
            else "/opt/libreoffice/"
        ).rstrip("/")
        if not program_prefix.startswith(expected_prefix):
            raise DockerRuntimeUnavailable("Runtime probe returned the wrong office program prefix")
        for key in ("uno_module", "pyuno_native_module"):
            value = str(probe_data.get(key) or "")
            if not value.startswith(program_prefix):
                raise DockerRuntimeUnavailable(
                    f"Runtime probe returned mismatched PyUNO provenance for {key}"
                )
        python_executable = str(probe_data.get("python_executable") or "")
        if not python_executable.startswith(program_prefix) or Path(python_executable).name not in {
            "python",
            "python.bin",
        }:
            raise DockerRuntimeUnavailable(
                "Runtime worker is not using LibreOffice's bundled Python interpreter"
            )
        if expected_variant == STOCK_RUNTIME_VARIANT and not str(
            probe_data.get("python_version") or ""
        ).startswith("3.12."):
            raise DockerRuntimeUnavailable("Runtime worker returned the wrong Python version")
        for key in (
            "office_sha256",
            "uno_module_sha256",
            "pyuno_native_sha256",
            "worker_wrapper_sha256",
        ):
            value = str(probe_data.get(key) or "")
            if len(value) != 64:
                raise DockerRuntimeUnavailable(f"Runtime probe omitted a valid {key}")
        manifest = probe_data.get("installed_package_manifest")
        if not isinstance(manifest, list) or not manifest:
            raise DockerRuntimeUnavailable("Runtime probe omitted the installed-package manifest")
        if expected_variant == STOCK_RUNTIME_VARIANT:
            pyuno_packages = [
                package
                for package in manifest
                if isinstance(package, dict) and package.get("name") == "libobasis26.2-pyuno"
            ]
            if len(pyuno_packages) != 1 or not str(pyuno_packages[0].get("version", "")).startswith(
                f"{LIBREOFFICE_VERSION}-"
            ):
                raise DockerRuntimeUnavailable("Runtime PyUNO package does not match LibreOffice")
        else:
            source_commit = str(probe_data.get("source_commit") or "")
            patch_set_sha256 = str(probe_data.get("patch_set_sha256") or "")
            valid_patch_identity = (
                patch_set_sha256 == "none"
                if expected_variant == "stock-source"
                else len(patch_set_sha256) == 64
            )
            if len(source_commit) != 40 or not valid_patch_identity:
                raise DockerRuntimeUnavailable(
                    "Source runtime omitted source commit or patch-set identity"
                )
        probed_architecture = str(probe_data.get("architecture") or "")
        normalized = {"aarch64": "arm64", "x86_64": "amd64"}.get(
            probed_architecture, probed_architecture
        )
        if image_architecture and normalized != image_architecture:
            raise DockerRuntimeUnavailable("Runtime architecture does not match the image identity")

    def request(
        self,
        payload: dict[str, Any],
        *,
        _identity: str | None = None,
    ) -> dict[str, Any]:
        """Execute a worker request without importing or starting office on the host."""
        image_id = _identity or self.resolve_identity(probe=False).image_id
        configured_temp_root = os.environ.get("XLSLIBERATOR_RUNTIME_TEMP_ROOT")
        temp_root = Path(configured_temp_root).resolve() if configured_temp_root else None
        if temp_root is not None:
            temp_root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(
            prefix="xlsliberator-office-job-", dir=temp_root
        ) as tmpdir:
            sandbox_root = Path(tmpdir)
            input_dir = sandbox_root / "input"
            job_dir = sandbox_root / "job"
            input_dir.mkdir(mode=0o755)
            job_dir.mkdir(mode=0o777)
            job_dir.chmod(0o777)
            mapped_payload, outputs = self._stage_paths(payload, input_dir)
            job_id = uuid.uuid4().hex
            container_name = f"xlsliberator-lo-{job_id}"
            command = self._container_command(
                image_id, job_dir, container_name, input_dir=input_dir
            )
            started_at = datetime.now(UTC)
            started_clock = time.monotonic()
            result = self._run_docker_cli(
                command,
                input_text=json.dumps(mapped_payload),
                timeout_seconds=float(payload.get("timeout_seconds") or self.timeout_seconds),
                allow_failure=True,
            )
            ended_at = datetime.now(UTC)
            response = self._parse_response(result)
            if response.get("success"):
                for container_name, destination in outputs:
                    staged = job_dir / container_name
                    if not staged.is_file():
                        raise DockerRuntimeUnavailable(
                            f"Runtime reported success without output: {container_name}"
                        )
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(staged, destination)
            data = dict(response.get("data") or {})
            data.setdefault("container_image_id", image_id)
            data.setdefault("container_image_reference", self.image)
            data.setdefault("container_name", container_name)
            data.setdefault("job_id", job_id)
            data.setdefault("started_at", started_at.isoformat())
            data.setdefault("ended_at", ended_at.isoformat())
            data.setdefault("duration_seconds", time.monotonic() - started_clock)
            data.setdefault("container_exit_code", result.returncode)
            data.setdefault("container_stderr", result.stderr[-16000:])
            capabilities = _granted_capabilities(mapped_payload)
            sandbox_job = SandboxJob(
                job_id=job_id,
                kind=_execution_kind(mapped_payload),
                image_reference=self.image,
                image_digest=image_id,
                mounts=[
                    SandboxMount(
                        source=str(input_dir),
                        destination="/input",
                        mode="ro",
                        purpose="input",
                    ),
                    SandboxMount(
                        source=str(job_dir),
                        destination="/job",
                        mode="rw",
                        purpose="job",
                    ),
                ],
                granted_capabilities=capabilities,
                policy=self.sandbox_policy,
            )
            data.setdefault("resource_policy", self.sandbox_policy.evidence())
            data.setdefault("sandbox_job", sandbox_job.model_dump(mode="json"))
            data.setdefault("granted_capabilities", capabilities)
            response["data"] = data
            self._assert_tag_still_resolves(image_id)
            return response

    def convert(self, input_path: Path, output_path: Path) -> dict[str, Any]:
        """Convert through the Docker worker and copy the resulting ODS atomically."""
        response = self.request(
            {
                "op": "convert_document",
                "input_path": str(input_path),
                "output_path": str(output_path),
                "timeout_seconds": self.timeout_seconds,
            }
        )
        if not response.get("success"):
            error = response.get("error") or {}
            raise DockerRuntimeUnavailable(str(error.get("message") or "conversion failed"))
        return response

    def validate_document(self, ods_path: Path, *, image_id: str | None = None) -> dict[str, Any]:
        """Run the complete non-mutating target validation scenario."""
        return self.request(
            {
                "op": "validate_document",
                "ods_path": str(ods_path),
                "timeout_seconds": self.timeout_seconds,
            },
            _identity=image_id,
        )

    def parse_formula(
        self,
        ods_path: Path,
        formula: str,
        *,
        sheet_name: str,
        cell_address: str,
        image_id: str | None = None,
    ) -> dict[str, Any]:
        """Parse and round-trip one target formula in its target document context."""
        return self.request(
            {
                "op": "parse_formula",
                "ods_path": str(ods_path),
                "formula": formula,
                "sheet_name": sheet_name,
                "cell_address": cell_address,
                "timeout_seconds": self.timeout_seconds,
            },
            _identity=image_id,
        )

    def evaluate_formula(
        self,
        formula: str,
        *,
        image_id: str | None = None,
    ) -> dict[str, Any]:
        """Evaluate one minimized formula entirely in a disposable office container."""
        return self.request(
            {
                "op": "evaluate_formula",
                "formula": formula,
                "timeout_seconds": self.timeout_seconds,
            },
            _identity=image_id,
        )

    def _require_docker(self) -> None:
        if Path(self.docker_executable).name != "docker":
            raise DockerRuntimeUnavailable(
                "Only the Docker CLI may cross the application-container process boundary"
            )
        if shutil.which(self.docker_executable) is None:
            raise DockerRuntimeUnavailable(
                "Docker executable is unavailable; host fallback is disabled"
            )

    def _container_command(
        self,
        image_id: str,
        job_dir: Path,
        container_name: str,
        *,
        input_dir: Path | None = None,
    ) -> list[str]:
        readonly_input = input_dir or job_dir
        docker_input = self._docker_mount_source(readonly_input)
        docker_job = self._docker_mount_source(job_dir)
        return [
            self.docker_executable,
            "run",
            "--rm",
            "--interactive",
            "--name",
            container_name,
            *docker_sandbox_arguments(self.sandbox_policy),
            "--mount",
            f"type=bind,src={docker_input},dst=/input,readonly",
            "--mount",
            f"type=bind,src={docker_job},dst=/job",
            image_id,
            "worker",
        ]

    @staticmethod
    def _docker_mount_source(path: Path) -> Path:
        """Translate a nested-orchestrator path to its Docker-host bind path.

        A trusted web container can stage jobs on a host bind mount while the
        Docker daemon runs outside that container.  The daemon must receive the
        host-side path, not the web container's mount point.  In the CI
        orchestrator both paths are identical, so no translation is configured.
        """
        host_root_raw = os.environ.get("XLSLIBERATOR_DOCKER_HOST_RUNTIME_TEMP_ROOT")
        if not host_root_raw:
            return path
        local_root_raw = os.environ.get("XLSLIBERATOR_RUNTIME_TEMP_ROOT")
        if not local_root_raw:
            raise DockerRuntimeUnavailable(
                "Docker host runtime root requires XLSLIBERATOR_RUNTIME_TEMP_ROOT"
            )
        local_root = Path(local_root_raw).resolve(strict=True)
        resolved = path.resolve(strict=True)
        if not resolved.is_relative_to(local_root):
            raise DockerRuntimeUnavailable(
                "Runtime job path escaped the configured local runtime root"
            )
        host_root = Path(host_root_raw)
        if not host_root.is_absolute() or ".." in host_root.parts:
            raise DockerRuntimeUnavailable(
                "Docker host runtime root must be absolute and normalized"
            )
        return host_root / resolved.relative_to(local_root)

    def _assert_tag_still_resolves(self, expected_image_id: str) -> None:
        """Block evidence if the configured mutable reference moved during a job."""
        result = self._run_docker_cli(
            [self.docker_executable, "image", "inspect", self.image, "--format", "{{.Id}}"],
            timeout_seconds=15,
        )
        if result.stdout.strip() != expected_image_id:
            raise DockerRuntimeUnavailable(
                "LibreOffice image tag drifted during the runtime job; evidence is invalid"
            )

    def _stage_paths(
        self, payload: dict[str, Any], input_dir: Path
    ) -> tuple[dict[str, Any], list[tuple[str, Path]]]:
        mapped = dict(payload)
        outputs: list[tuple[str, Path]] = []
        for key in ("input_path", "ods_path"):
            raw = payload.get(key)
            if not raw:
                continue
            source = self.workspace_paths.input_file(str(raw))
            validate_untrusted_workbook(source)
            staged_name = f"input-{key}{source.suffix}"
            shutil.copy2(source, input_dir / staged_name)
            (input_dir / staged_name).chmod(0o444)
            mapped[key] = f"/input/{staged_name}"
        for key, prefix, default_suffix in (
            ("output_path", "output", ".ods"),
            ("attachment_output_path", "attachment-output", ".bin"),
        ):
            raw_output = payload.get(key)
            if not raw_output:
                continue
            destination = self.workspace_paths.output_file(str(raw_output))
            staged_name = f"{prefix}{destination.suffix or default_suffix}"
            mapped[key] = f"/job/{staged_name}"
            outputs.append((staged_name, destination))
        mapped.pop("office_executable", None)
        return mapped, outputs

    def _run_docker_cli(
        self,
        command: list[str],
        *,
        input_text: str | None = None,
        timeout_seconds: float,
        allow_failure: bool = False,
    ) -> subprocess.CompletedProcess[str]:
        self._require_docker()
        if not command or command[0] != self.docker_executable:
            raise DockerRuntimeUnavailable(
                "LibreOffice runtime may start only the configured Docker CLI"
            )
        try:
            result = subprocess.run(
                command,
                input=input_text,
                capture_output=True,
                text=True,
                timeout=timeout_seconds,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            self._cleanup_timed_out_container(command)
            raise DockerRuntimeTimeout(
                f"Docker runtime exceeded the {timeout_seconds:g}s wall-time limit"
            ) from exc
        except (OSError, subprocess.SubprocessError) as exc:
            raise DockerRuntimeUnavailable(f"Docker runtime invocation failed: {exc}") from exc
        if result.returncode != 0 and not allow_failure:
            raise MalformedWorkerResponse(
                f"Docker runtime exited {result.returncode}: "
                f"{result.stderr.strip() or result.stdout.strip()}"
            )
        return result

    def _cleanup_timed_out_container(self, command: list[str]) -> None:
        """Remove the named container so its complete process tree is killed."""

        if "run" not in command or "--name" not in command:
            return
        try:
            name = command[command.index("--name") + 1]
            subprocess.run(
                [self.docker_executable, "rm", "--force", name],
                capture_output=True,
                text=True,
                timeout=15,
                check=False,
                env={"PATH": os.environ.get("PATH", "")},
            )
        except (OSError, ValueError, subprocess.SubprocessError):
            return

    @staticmethod
    def _parse_response(result: subprocess.CompletedProcess[str]) -> dict[str, Any]:
        try:
            response = json.loads(result.stdout.strip())
        except json.JSONDecodeError as exc:
            failure = (
                f"exit {result.returncode}: {result.stderr.strip()}" if result.returncode else ""
            )
            raise MalformedWorkerResponse(
                f"Docker worker returned malformed JSON {failure}: {result.stdout[:200]}"
            ) from exc
        if not isinstance(response, dict):
            raise MalformedWorkerResponse("Docker worker response is not an object")
        return response


def _granted_capabilities(payload: dict[str, Any]) -> list[str]:
    environment = payload.get("environment")
    if not isinstance(environment, dict):
        return []
    granted = environment.get("granted_capabilities")
    values = {str(item) for item in granted} if isinstance(granted, list) else set()
    typed = environment.get("typed_capabilities")
    if isinstance(typed, list):
        values.update(
            str(item["capability"])
            for item in typed
            if isinstance(item, dict) and item.get("granted") and item.get("capability")
        )
    return sorted(values)


def _execution_kind(payload: dict[str, Any]) -> ExecutionKind:
    environment = payload.get("environment")
    capabilities = _granted_capabilities(payload)
    if "macro_execution" in capabilities:
        return ExecutionKind.MACRO
    if isinstance(environment, dict) and "gui_interaction" in capabilities:
        return ExecutionKind.GUI
    return ExecutionKind.LIBREOFFICE_TARGET

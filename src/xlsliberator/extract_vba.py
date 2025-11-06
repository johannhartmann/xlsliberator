"""VBA code extraction and dependency analysis (Phase F7)."""

import re
from collections import defaultdict
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from loguru import logger

try:
    import oletools.olevba as olevba
except ImportError:
    olevba = None


class VBAExtractionError(Exception):
    """Raised when VBA extraction fails."""


class VBAModuleType(Enum):
    """VBA module types."""

    STANDARD = "Standard"  # Standard code module
    CLASS = "Class"  # Class module
    FORM = "Form"  # UserForm
    DOCUMENT = "Document"  # ThisWorkbook, Sheet modules
    UNKNOWN = "Unknown"


@dataclass
class VBAModuleIR:
    """Intermediate representation of a VBA module."""

    name: str
    module_type: VBAModuleType
    source_code: str
    procedures: list[str] = field(default_factory=list)
    dependencies: set[str] = field(default_factory=set)
    api_calls: dict[str, int] = field(default_factory=dict)


@dataclass
class VBADependencyGraph:
    """VBA module dependency graph."""

    modules: dict[str, VBAModuleIR]
    edges: dict[str, set[str]]  # module_name -> set of dependencies
    api_usage: dict[str, int]  # API call -> count across all modules


def extract_vba_modules(file_path: str | Path) -> list[VBAModuleIR]:
    """Extract VBA modules from Excel file.

    Args:
        file_path: Path to Excel file (.xlsm, .xlsb, .xls)

    Returns:
        List of VBAModuleIR objects

    Raises:
        VBAExtractionError: If extraction fails

    Note:
        Phase F7 implementation - uses oletools to extract VBA code.
        Statically analyzes code without execution.
    """
    if olevba is None:
        raise VBAExtractionError("oletools not available. Install with: pip install oletools")

    file_path = Path(file_path)

    if not file_path.exists():
        raise VBAExtractionError(f"File not found: {file_path}")

    logger.info(f"Extracting VBA modules from {file_path}")

    try:
        vba_parser = olevba.VBA_Parser(str(file_path))

        if not vba_parser.detect_vba_macros():
            logger.info("No VBA macros found in file")
            return []

        modules: list[VBAModuleIR] = []

        # Extract each module
        for _filename, _stream_path, vba_filename, vba_code in vba_parser.extract_macros():
            # Determine module type from stream path or filename
            module_type = _detect_module_type(vba_filename, vba_code)

            # Extract procedures
            procedures = _extract_procedures(vba_code)

            # Extract dependencies (module references)
            dependencies = _extract_dependencies(vba_code)

            # Extract API calls
            api_calls = _extract_api_calls(vba_code)

            module_ir = VBAModuleIR(
                name=vba_filename,
                module_type=module_type,
                source_code=vba_code,
                procedures=procedures,
                dependencies=dependencies,
                api_calls=api_calls,
            )

            modules.append(module_ir)
            logger.debug(
                f"Extracted module '{vba_filename}': "
                f"{len(procedures)} procedures, {len(dependencies)} dependencies, "
                f"{sum(api_calls.values())} API calls"
            )

        logger.success(f"Extracted {len(modules)} VBA modules")
        return modules

    except Exception as e:
        raise VBAExtractionError(f"Failed to extract VBA: {e}") from e


def _detect_module_type(module_name: str, source_code: str) -> VBAModuleType:
    """Detect VBA module type from name and content.

    Args:
        module_name: Module filename
        source_code: VBA source code

    Returns:
        VBAModuleType enum value
    """
    # Check for explicit type declarations
    if "Attribute VB_Name" in source_code:
        if "Attribute VB_PredeclaredId = True" in source_code:
            return VBAModuleType.STANDARD
        if "Attribute VB_Exposed = True" in source_code:
            return VBAModuleType.CLASS

    # Check module name patterns
    module_lower = module_name.lower()

    if module_lower.startswith("userform") or module_lower.startswith("frm"):
        return VBAModuleType.FORM

    if module_lower in ("thisworkbook", "thisworksheet") or module_lower.startswith("sheet"):
        return VBAModuleType.DOCUMENT

    if module_lower.startswith("class") or module_lower.endswith("cls"):
        return VBAModuleType.CLASS

    # Check content for form indicators
    if "Begin {" in source_code or "UserForm" in source_code:
        return VBAModuleType.FORM

    # Default to standard module
    return VBAModuleType.STANDARD


def _extract_procedures(source_code: str) -> list[str]:
    """Extract procedure names from VBA source code.

    Args:
        source_code: VBA source code

    Returns:
        List of procedure names (Sub, Function, Property)
    """
    procedures = []

    # Regex patterns for procedure declarations
    patterns = [
        r"(?:Public|Private|Friend)?\s+(?:Static\s+)?Sub\s+(\w+)\s*\(",
        r"(?:Public|Private|Friend)?\s+(?:Static\s+)?Function\s+(\w+)\s*\(",
        r"(?:Public|Private|Friend)?\s+Property\s+(?:Get|Let|Set)\s+(\w+)\s*\(",
    ]

    for pattern in patterns:
        matches = re.finditer(pattern, source_code, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            proc_name = match.group(1)
            if proc_name not in procedures:
                procedures.append(proc_name)

    return procedures


def _extract_dependencies(source_code: str) -> set[str]:
    """Extract module dependencies from VBA source code.

    Args:
        source_code: VBA source code

    Returns:
        Set of referenced module names
    """
    dependencies = set()

    # Look for module-level references (calls to other modules)
    # This is a simplified approach - full parser would be more accurate

    # Find procedure calls that might be to other modules
    # Pattern: ModuleName.ProcedureName
    module_calls = re.findall(r"\b([A-Z]\w+)\.(\w+)", source_code)

    for module_name, _ in module_calls:
        # Filter out common VBA keywords and objects
        if module_name not in (
            "Application",
            "WorksheetFunction",
            "ActiveSheet",
            "ActiveWorkbook",
            "ThisWorkbook",
            "Range",
            "Cells",
            "Worksheets",
            "Debug",
            "VBA",
        ):
            dependencies.add(module_name)

    return dependencies


def _extract_api_calls(source_code: str) -> dict[str, int]:
    """Extract API calls and count occurrences.

    Args:
        source_code: VBA source code

    Returns:
        Dict mapping API call to count

    Note:
        Tracks key Excel/VBA APIs:
        - Range/Cells/Worksheets
        - Application.*
        - WorksheetFunction.*
        - UserForm
        - DoEvents
    """
    api_calls: dict[str, int] = defaultdict(int)

    # Define API patterns to track
    api_patterns = {
        "Range": r"\bRange\s*\(",
        "Cells": r"\bCells\s*\(",
        "Worksheets": r"\bWorksheets\s*\(",
        "Workbooks": r"\bWorkbooks\s*\(",
        "ActiveSheet": r"\bActiveSheet\b",
        "ActiveWorkbook": r"\bActiveWorkbook\b",
        "ThisWorkbook": r"\bThisWorkbook\b",
        "Application": r"\bApplication\.",
        "WorksheetFunction": r"\bWorksheetFunction\.",
        "UserForm": r"\bUserForm\b",
        "DoEvents": r"\bDoEvents\b",
        "MsgBox": r"\bMsgBox\s*\(",
        "InputBox": r"\bInputBox\s*\(",
        "CreateObject": r"\bCreateObject\s*\(",
        "GetObject": r"\bGetObject\s*\(",
    }

    # Count occurrences of each API
    for api_name, pattern in api_patterns.items():
        matches = re.findall(pattern, source_code, re.IGNORECASE)
        if matches:
            api_calls[api_name] = len(matches)

    return dict(api_calls)


def build_vba_dependency_graph(modules: list[VBAModuleIR]) -> VBADependencyGraph:
    """Build dependency graph from VBA modules.

    Args:
        modules: List of VBA modules

    Returns:
        VBADependencyGraph with modules, edges, and API usage

    Note:
        Phase F7 implementation - creates static dependency graph.
        Aggregates API usage across all modules.
    """
    # Create module lookup
    module_dict = {mod.name: mod for mod in modules}

    # Build edges (module dependencies)
    edges: dict[str, set[str]] = {}
    for module in modules:
        edges[module.name] = module.dependencies.copy()

    # Aggregate API usage
    api_usage: dict[str, int] = defaultdict(int)
    for module in modules:
        for api_call, count in module.api_calls.items():
            api_usage[api_call] += count

    graph = VBADependencyGraph(
        modules=module_dict,
        edges=edges,
        api_usage=dict(api_usage),
    )

    logger.info(
        f"Built dependency graph: {len(modules)} modules, "
        f"{sum(len(deps) for deps in edges.values())} edges, "
        f"{len(api_usage)} unique APIs"
    )

    return graph


def get_top_api_calls(graph: VBADependencyGraph, top_n: int = 10) -> list[tuple[str, int]]:
    """Get top N most-used API calls.

    Args:
        graph: VBA dependency graph
        top_n: Number of top APIs to return

    Returns:
        List of (api_name, count) tuples, sorted by count descending
    """
    sorted_apis = sorted(graph.api_usage.items(), key=lambda x: x[1], reverse=True)
    return sorted_apis[:top_n]


def detect_cycles(graph: VBADependencyGraph) -> list[list[str]]:
    """Detect circular dependencies in VBA module graph.

    Args:
        graph: VBA dependency graph

    Returns:
        List of cycles, each cycle is a list of module names
    """
    cycles = []

    def dfs(node: str, visited: set[str], path: list[str]) -> None:
        if node in path:
            # Found a cycle
            cycle_start = path.index(node)
            cycle = path[cycle_start:] + [node]
            cycles.append(cycle)
            return

        if node in visited:
            return

        visited.add(node)
        path.append(node)

        for dep in graph.edges.get(node, set()):
            if dep in graph.modules:  # Only follow edges to known modules
                dfs(dep, visited, path)

        path.pop()

    visited: set[str] = set()
    for module_name in graph.modules:
        if module_name not in visited:
            dfs(module_name, visited, [])

    return cycles

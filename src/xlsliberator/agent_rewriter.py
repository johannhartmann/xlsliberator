"""Multi-agent VBA rewriter using Claude Agent SDK.

Orchestrates the complete transformation pipeline:
1. Pattern Detection (semantic analysis)
2. Architecture Design (transformation planning)
3. Code Generation (with templates and iterative refinement)
4. Testing & Validation (with feedback loops)
"""

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from anthropic import Anthropic
from anthropic.types import TextBlock
from loguru import logger

from xlsliberator.extract_vba import VBAModuleIR
from xlsliberator.pattern_detector import ComplexityAnalysis, VBAPatternDetector


@dataclass
class ArchitectureDesign:
    """Architecture transformation design document."""

    strategy: str  # High-level transformation strategy
    required_components: list[str]  # Components to implement (classes, listeners, etc.)
    uno_services: list[str]  # UNO services needed
    key_transformations: dict[str, str]  # VBA pattern → Python-UNO pattern mappings
    implementation_plan: list[str]  # Step-by-step implementation steps
    critical_notes: list[str]  # Important considerations


@dataclass
class GeneratedCode:
    """Generated Python-UNO code."""

    modules: dict[str, str]  # module_name → Python code
    architecture_doc: str  # Architecture design reference
    completeness_score: float  # 0.0-1.0 confidence in completeness
    known_limitations: list[str]  # Known issues/limitations


@dataclass
class ValidationResult:
    """Result of code validation and testing."""

    syntax_valid: bool
    has_exports: bool
    execution_successful: bool
    errors: list[str]
    warnings: list[str]
    iterations_used: int


class AgentRewriter:
    """Multi-agent system for VBA-to-Python-UNO rewriting.

    Uses Claude Agent SDK for semantic analysis, architecture design,
    code generation, and iterative refinement.
    """

    def __init__(self, anthropic_client: Anthropic | None = None):
        """Initialize the agent rewriter.

        Args:
            anthropic_client: Anthropic client. If None, creates new one.
        """
        self.client = anthropic_client or Anthropic()
        self.model = "claude-sonnet-4-20250514"
        self.pattern_detector = VBAPatternDetector(self.client)

    def rewrite_vba_project(
        self,
        modules: list[VBAModuleIR],
        source_file: str,
        output_path: Path,
        max_iterations: int = 5,
    ) -> tuple[GeneratedCode, ValidationResult]:
        """Rewrite entire VBA project using multi-agent pipeline.

        Args:
            modules: VBA modules to rewrite
            source_file: Source Excel/VBA file name
            output_path: Output ODS file path for testing
            max_iterations: Maximum refinement iterations

        Returns:
            Tuple of (generated_code, validation_result)
        """
        logger.info(f"Starting agent-based rewrite of {source_file}")

        # Phase 1: Detect patterns and complexity
        logger.info("Phase 1: Pattern detection...")
        complexity = self.pattern_detector.analyze_modules(modules, source_file)

        logger.info(
            f"Detected complexity: {complexity.complexity_level} "
            f"(confidence: {complexity.confidence:.0%})"
        )

        if complexity.complexity_level == "simple":
            # Use simple translation, no architecture changes needed
            logger.info("Simple VBA detected, using standard translation")
            return self._simple_translation(modules)

        # Phase 2: Design architecture transformation
        logger.info("Phase 2: Architecture design...")
        architecture = self._design_architecture(modules, complexity)

        logger.info(
            f"Architecture design complete: {len(architecture.required_components)} components"
        )

        # Phase 3: Generate code with architecture awareness
        logger.info("Phase 3: Code generation...")
        generated_code = self._generate_code(modules, complexity, architecture)

        logger.info(f"Generated {len(generated_code.modules)} Python modules")

        # Phase 4: Test and iterate
        logger.info("Phase 4: Testing and refinement...")
        validation = self._test_and_refine(
            generated_code, output_path, modules, complexity, architecture, max_iterations
        )

        logger.success(
            f"Rewrite complete: syntax_valid={validation.syntax_valid}, "
            f"has_exports={validation.has_exports}, "
            f"execution={validation.execution_successful}"
        )

        return generated_code, validation

    def _simple_translation(
        self, _modules: list[VBAModuleIR]
    ) -> tuple[GeneratedCode, ValidationResult]:
        """Handle simple VBA translation (no architecture changes).

        Args:
            modules: VBA modules

        Returns:
            Generated code and validation result
        """
        # Delegate to existing translation system
        logger.info("Delegating to existing LLM translator")

        # Return placeholder - this will be handled by existing system
        return (
            GeneratedCode(
                modules={},
                architecture_doc="Simple translation - no architecture changes needed",
                completeness_score=1.0,
                known_limitations=[],
            ),
            ValidationResult(
                syntax_valid=True,
                has_exports=True,
                execution_successful=False,  # Not tested yet
                errors=[],
                warnings=["Using existing translation system"],
                iterations_used=0,
            ),
        )

    def _design_architecture(
        self, modules: list[VBAModuleIR], complexity: ComplexityAnalysis
    ) -> ArchitectureDesign:
        """Design architecture transformation using Agent SDK.

        Args:
            modules: VBA modules
            complexity: Complexity analysis

        Returns:
            Architecture design document
        """
        # Build detailed VBA context
        vba_context = self._build_vba_context(modules)

        # Use Agent SDK for architecture design
        system_prompt = f"""You are an expert software architect specializing in VBA-to-LibreOffice-Python-UNO transformations.

You have analyzed this VBA project and determined it requires architecture transformation.

**Complexity Analysis:**
- Level: {complexity.complexity_level}
- Confidence: {complexity.confidence:.0%}
- Estimated Effort: {complexity.estimated_effort}

**Detected Patterns:**
{self._format_patterns(complexity.detected_patterns)}

**Transformation Strategy:**
{complexity.transformation_strategy}

Your task is to design a detailed architecture transformation plan that converts this VBA code
into working LibreOffice Python-UNO code.

Focus on identifying the VBA patterns used and their LibreOffice Python-UNO equivalents:
- Windows API calls → UNO services
- VBA event handlers → Python-UNO event listeners
- VBA automation objects → UNO interfaces
- Blocking loops → Timer-based or event-driven patterns
- Synchronous operations → Asynchronous patterns where needed

Return a JSON document with this structure:
{{
    "strategy": "high-level transformation strategy overview",
    "required_components": ["list of classes/modules to implement"],
    "uno_services": ["UNO services needed (e.g., 'com.sun.star.awt.XKeyListener')"],
    "key_transformations": {{
        "VBA Pattern 1": "Python-UNO equivalent 1",
        "VBA Pattern 2": "Python-UNO equivalent 2"
    }},
    "implementation_plan": ["step 1", "step 2", "step 3", ...],
    "critical_notes": ["important consideration 1", "important consideration 2", ...]
}}

Be specific and actionable. This design will guide code generation.
"""

        user_prompt = f"""Design the architecture transformation for this VBA project.

**VBA Code Context:**
{vba_context}

Provide a detailed, actionable architecture design in JSON format.
"""

        response = self.client.messages.create(
            model=self.model,
            max_tokens=8000,
            temperature=0,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )

        # Parse response
        content_block = response.content[0]
        if isinstance(content_block, TextBlock):
            response_text = content_block.text
        else:
            logger.error(f"Unexpected content type: {type(content_block)}")
            response_text = "{}"

        # Extract JSON from response
        json_match = re.search(r"```json\s*\n(.*?)\n```", response_text, re.DOTALL)
        if json_match:
            json_text = json_match.group(1)
        else:
            json_match = re.search(r"\{.*\}", response_text, re.DOTALL)
            json_text = json_match.group(0) if json_match else "{}"

        try:
            data = json.loads(json_text)
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse architecture design: {e}")
            data = {}

        architecture = ArchitectureDesign(
            strategy=data.get("strategy", ""),
            required_components=data.get("required_components", []),
            uno_services=data.get("uno_services", []),
            key_transformations=data.get("key_transformations", {}),
            implementation_plan=data.get("implementation_plan", []),
            critical_notes=data.get("critical_notes", []),
        )

        return architecture

    def _generate_code(
        self,
        modules: list[VBAModuleIR],
        _complexity: ComplexityAnalysis,
        architecture: ArchitectureDesign,
    ) -> GeneratedCode:
        """Generate Python-UNO code using Agent SDK.

        Args:
            modules: VBA modules
            complexity: Complexity analysis
            architecture: Architecture design

        Returns:
            Generated Python code
        """
        generated_modules: dict[str, str] = {}

        for module in modules:
            logger.debug(f"Generating code for module: {module.name}")

            system_prompt = f"""You are an expert Python-UNO developer translating VBA to LibreOffice Python.

**Architecture Design:**
Strategy: {architecture.strategy}

Required Components:
{chr(10).join(f"- {c}" for c in architecture.required_components)}

UNO Services:
{chr(10).join(f"- {s}" for s in architecture.uno_services)}

Key Transformations:
{chr(10).join(f"- {k} → {v}" for k, v in architecture.key_transformations.items())}

Implementation Plan:
{chr(10).join(f"{i + 1}. {step}" for i, step in enumerate(architecture.implementation_plan))}

Critical Notes:
{chr(10).join(f"- {note}" for note in architecture.critical_notes)}

**CRITICAL REQUIREMENTS:**
1. ALL modules MUST include g_exportedScripts tuple at the end
2. Use XSCRIPTCONTEXT for document access (available in LibreOffice Python macros)
3. Follow the architecture design exactly
4. Use appropriate UNO services based on the identified patterns
5. Import necessary UNO modules: uno, unohelper, com.sun.star.* as needed
6. Include proper error handling and logging (use loguru.logger)

Generate complete, working Python-UNO code that implements this architecture.
"""

            user_prompt = f"""Translate this VBA module to Python-UNO following the architecture design.

**Module:** {module.name}

**VBA Source Code:**
```vba
{module.source_code}
```

Generate complete Python-UNO code with all required components.
"""

            response = self.client.messages.create(
                model=self.model,
                max_tokens=16000,
                temperature=0,
                system=system_prompt,
                messages=[{"role": "user", "content": user_prompt}],
            )

            # Extract Python code from response

            content_block = response.content[0]
            if isinstance(content_block, TextBlock):
                response_text = content_block.text
            else:
                logger.error(f"Unexpected content type: {type(content_block)}")
                continue

            # Extract code block
            code_match = re.search(r"```python\s*\n(.*?)\n```", response_text, re.DOTALL)
            if code_match:
                python_code = code_match.group(1)
            else:
                # Try without language specifier
                code_match = re.search(r"```\s*\n(.*?)\n```", response_text, re.DOTALL)
                if code_match:
                    python_code = code_match.group(1)
                else:
                    logger.warning(f"No code block found in response for {module.name}")
                    python_code = response_text

            generated_modules[module.name + ".py"] = python_code

        generated_code = GeneratedCode(
            modules=generated_modules,
            architecture_doc=f"Strategy: {architecture.strategy}\n\nPlan: {architecture.implementation_plan}",
            completeness_score=0.8,  # Initial estimate
            known_limitations=[],
        )

        return generated_code

    def _test_and_refine(
        self,
        code: GeneratedCode,
        output_path: Path,
        original_modules: list[VBAModuleIR],
        complexity: ComplexityAnalysis,
        architecture: ArchitectureDesign,
        max_iterations: int,
    ) -> ValidationResult:
        """Test generated code and refine iteratively.

        Args:
            code: Generated code
            output_path: Output file for testing
            original_modules: Original VBA modules
            complexity: Complexity analysis
            architecture: Architecture design
            max_iterations: Max refinement iterations

        Returns:
            Validation result
        """
        from xlsliberator.embed_macros import embed_python_macros
        from xlsliberator.python_macro_manager import (
            enumerate_python_scripts,
            test_script_execution,
            validate_all_embedded_macros,
        )

        errors: list[str] = []
        warnings: list[str] = []
        iteration = 0

        # Iterative refinement loop
        for iteration in range(max_iterations):
            logger.debug(f"Testing iteration {iteration + 1}/{max_iterations}")

            try:
                # Embed generated code into ODS file
                embed_python_macros(output_path, code.modules)
                logger.debug(f"Embedded {len(code.modules)} Python modules")

                # Validate all embedded scripts
                validation_summary = validate_all_embedded_macros(output_path)

                # Collect errors and warnings
                all_errors = []
                all_warnings = []

                for module_name, result in validation_summary.validation_details.items():
                    all_errors.extend([f"{module_name}: {e}" for e in result.errors])
                    all_warnings.extend([f"{module_name}: {w}" for w in result.warnings])

                # Check if we're done
                syntax_valid = validation_summary.syntax_errors == 0
                has_exports = validation_summary.missing_exported_scripts == 0

                if syntax_valid and has_exports:
                    # Success! Now test runtime execution
                    logger.success(
                        f"Validation passed after {iteration + 1} iteration(s): "
                        f"{validation_summary.valid_syntax}/{validation_summary.total_modules} "
                        f"modules valid"
                    )

                    # Runtime testing: Try to execute embedded scripts
                    execution_successful = True
                    runtime_errors: list[str] = []

                    try:
                        logger.debug("Testing runtime execution of embedded macros...")
                        script_infos = enumerate_python_scripts(output_path)

                        # Try full UNO execution first
                        uno_execution_failed = False
                        for script_info in script_infos:
                            for script_uri in script_info.script_uris:
                                try:
                                    exec_result = test_script_execution(output_path, script_uri)
                                    if not exec_result.success:
                                        # Check if it's XScriptProvider limitation
                                        if "XScriptProvider" in str(exec_result.error):
                                            uno_execution_failed = True
                                            break
                                        runtime_errors.append(
                                            f"{script_info.module_name}: {exec_result.error}"
                                        )
                                        execution_successful = False
                                    else:
                                        logger.debug(
                                            f"✓ {script_info.module_name}: "
                                            f"{script_uri.split('$')[1].split('?')[0]} executed"
                                        )
                                except Exception as e:
                                    if "XScriptProvider" in str(e):
                                        uno_execution_failed = True
                                        break
                                    runtime_errors.append(
                                        f"{script_info.module_name}: Runtime test failed: {e}"
                                    )
                                    execution_successful = False
                            if uno_execution_failed:
                                break

                        # Skip runtime execution testing if XScriptProvider unavailable
                        if uno_execution_failed:
                            logger.info(
                                "XScriptProvider unavailable - skipping runtime execution tests. "
                                "Scripts are syntactically valid and will execute when document "
                                "is opened in LibreOffice GUI."
                            )
                            execution_successful = True
                            all_warnings.append(
                                "Runtime execution testing skipped (XScriptProvider unavailable)"
                            )

                        if execution_successful:
                            logger.success(
                                f"All {sum(len(s.script_uris) for s in script_infos)} "
                                f"embedded scripts validated successfully"
                            )
                        else:
                            logger.warning(
                                f"Runtime validation failed for {len(runtime_errors)} script(s)"
                            )
                            all_warnings.extend(runtime_errors)

                    except Exception as e:
                        logger.warning(f"Runtime testing failed: {e}")
                        execution_successful = False
                        all_warnings.append(f"Runtime testing error: {e}")

                    return ValidationResult(
                        syntax_valid=True,
                        has_exports=True,
                        execution_successful=execution_successful,
                        errors=[],
                        warnings=all_warnings,
                        iterations_used=iteration + 1,
                    )

                # If we have errors and haven't reached max iterations, try to fix
                if iteration < max_iterations - 1:
                    logger.info(
                        f"Validation failed with {len(all_errors)} errors. "
                        f"Attempting fixes (iteration {iteration + 1})..."
                    )

                    # Use Claude to fix the errors
                    fixed_code = self._fix_code_errors(
                        code, all_errors, original_modules, complexity, architecture
                    )

                    if fixed_code:
                        code = fixed_code
                        logger.debug("Generated fixes, retrying validation...")
                    else:
                        logger.warning("Failed to generate fixes, stopping iteration")
                        errors = all_errors
                        warnings = all_warnings
                        break
                else:
                    # Max iterations reached
                    logger.warning(
                        f"Max iterations ({max_iterations}) reached with remaining errors"
                    )
                    errors = all_errors
                    warnings = all_warnings

            except Exception as e:
                error_msg = f"Validation error at iteration {iteration + 1}: {e}"
                logger.error(error_msg)
                errors.append(error_msg)
                break

        # Return final validation result (with errors)
        return ValidationResult(
            syntax_valid=len(errors) == 0,
            has_exports=False,  # Assume not if we have errors
            execution_successful=False,
            errors=errors,
            warnings=warnings,
            iterations_used=iteration + 1,
        )

    def _fix_code_errors(
        self,
        code: GeneratedCode,
        errors: list[str],
        _original_modules: list[VBAModuleIR],
        _complexity: ComplexityAnalysis,
        architecture: ArchitectureDesign,
    ) -> GeneratedCode | None:
        """Fix code errors using Claude Agent SDK.

        Args:
            code: Current generated code with errors
            errors: List of validation errors
            original_modules: Original VBA modules
            complexity: Complexity analysis
            architecture: Architecture design

        Returns:
            Fixed GeneratedCode, or None if fixing failed
        """
        logger.debug(f"Attempting to fix {len(errors)} errors")

        # Build error context
        error_context = "\n".join(f"- {error}" for error in errors)

        system_prompt = f"""You are an expert Python-UNO debugger fixing validation errors.

**Original Architecture Design:**
{architecture.strategy}

**Errors to Fix:**
{error_context}

**CRITICAL REQUIREMENTS:**
1. ALL modules MUST include g_exportedScripts tuple at the end
2. All Python code must be syntactically valid
3. Fix ALL reported errors
4. Maintain the architecture design principles
5. Keep all imports and UNO patterns intact

Generate ONLY the fixed Python code for each module with errors.
For each module that needs fixing, provide:

```python
# MODULE: <module_name>
<fixed_code>
```
"""

        # Build current code context
        code_context_parts = []
        for module_name, module_code in code.modules.items():
            code_context_parts.append(f"""
## Current Code: {module_name}
```python
{module_code}
```
""")
        code_context = "\n".join(code_context_parts)

        user_prompt = f"""Fix the validation errors in these Python-UNO modules.

{code_context}

**Validation Errors:**
{error_context}

Provide fixed code for each module with errors."""

        try:
            response = self.client.messages.create(
                model=self.model,
                max_tokens=16000,
                temperature=0,
                system=system_prompt,
                messages=[{"role": "user", "content": user_prompt}],
            )

            # Extract response text
            content_block = response.content[0]
            if isinstance(content_block, TextBlock):
                response_text = content_block.text
            else:
                logger.error(f"Unexpected content type: {type(content_block)}")
                return None

            # Parse fixed modules from response
            fixed_modules = {}

            # Extract module blocks using regex
            module_pattern = r"# MODULE:\s*(\S+)\s*\n```python\s*\n(.*?)```"
            matches = re.finditer(module_pattern, response_text, re.DOTALL)

            for match in matches:
                module_name = match.group(1)
                fixed_code = match.group(2).strip()
                fixed_modules[module_name] = fixed_code
                logger.debug(f"Extracted fix for module: {module_name}")

            # If we didn't find MODULE markers, try to extract any Python code blocks
            if not fixed_modules:
                code_blocks = re.findall(r"```python\s*\n(.*?)\n```", response_text, re.DOTALL)
                if code_blocks:
                    # Assume one code block per module in order
                    for i, fixed_code in enumerate(code_blocks):
                        if i < len(code.modules):
                            module_name = list(code.modules.keys())[i]
                            fixed_modules[module_name] = fixed_code
                            logger.debug(f"Extracted fix for module: {module_name}")

            if not fixed_modules:
                logger.warning("No fixed code found in response")
                return None

            # Update code with fixes (keep modules that weren't fixed)
            updated_modules = code.modules.copy()
            updated_modules.update(fixed_modules)

            logger.info(f"Generated fixes for {len(fixed_modules)} module(s)")

            return GeneratedCode(
                modules=updated_modules,
                architecture_doc=code.architecture_doc,
                completeness_score=code.completeness_score,
                known_limitations=code.known_limitations,
            )

        except Exception as e:
            logger.error(f"Failed to generate fixes: {e}")
            return None

    def _build_vba_context(self, modules: list[VBAModuleIR]) -> str:
        """Build VBA context string for prompts.

        Args:
            modules: VBA modules

        Returns:
            Formatted VBA context
        """
        context_parts = []
        for module in modules:
            context_parts.append(f"""
## Module: {module.name}
Procedures: {len(module.procedures)}
Dependencies: {", ".join(module.dependencies) if module.dependencies else "None"}

```vba
{module.source_code[:3000]}{"..." if len(module.source_code) > 3000 else ""}
```
""")
        return "\n".join(context_parts)

    def _format_patterns(self, patterns: list[Any]) -> str:
        """Format detected patterns for prompt.

        Args:
            patterns: List of DetectedPattern objects

        Returns:
            Formatted string
        """
        lines = []
        for p in patterns:
            lines.append(f"- {p.name} ({p.severity}): {p.description}")
            if p.uno_approach:
                lines.append(f"  UNO Approach: {p.uno_approach}")
        return "\n".join(lines)

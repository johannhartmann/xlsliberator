"""Agent-based VBA pattern detection and complexity analysis.

Uses Claude Agent SDK to semantically analyze VBA code and determine
the appropriate transformation strategy.
"""

import json
import re
from dataclasses import dataclass

from anthropic import Anthropic
from anthropic.types import TextBlock
from loguru import logger

from xlsliberator.extract_vba import VBAModuleIR


@dataclass
class DetectedPattern:
    """A pattern detected in VBA code."""

    name: str
    description: str
    severity: str  # "simple" | "moderate" | "complex" | "blocking"
    vba_examples: list[str]
    uno_approach: str
    requires_architecture_change: bool


@dataclass
class ComplexityAnalysis:
    """Result of VBA complexity analysis."""

    complexity_level: str  # "simple" | "game" | "advanced" | "untranslatable"
    confidence: float  # 0.0-1.0
    detected_patterns: list[DetectedPattern]
    reasoning: str
    transformation_strategy: str
    estimated_effort: str  # "low" | "medium" | "high" | "manual"
    blockers: list[str]  # Issues that prevent automatic translation
    recommendations: list[str]


class VBAPatternDetector:
    """Agent-based VBA pattern detector using Claude for semantic analysis."""

    def __init__(self, anthropic_client: Anthropic | None = None):
        """Initialize the pattern detector.

        Args:
            anthropic_client: Anthropic client instance. If None, creates new one.
        """
        self.client = anthropic_client or Anthropic()
        self.model = "claude-sonnet-4-5"

    def analyze_modules(self, modules: list[VBAModuleIR], source_file: str) -> ComplexityAnalysis:
        """Analyze VBA modules and determine complexity.

        Args:
            modules: List of VBA modules to analyze
            source_file: Source Excel/VBA file name for context

        Returns:
            ComplexityAnalysis with detected patterns and strategy
        """
        logger.info(f"Starting agent-based complexity analysis of {len(modules)} modules")

        # Build context from VBA modules
        analysis_context = self._build_analysis_context(modules, source_file)

        # Use Claude Agent SDK to analyze
        response = self.client.messages.create(
            model=self.model,
            max_tokens=8000,
            temperature=0,
            system=self._get_system_prompt(),
            messages=[
                {
                    "role": "user",
                    "content": self._build_analysis_prompt(analysis_context),
                }
            ],
        )

        # Parse agent response into structured analysis
        # Get text from first content block
        content_block = response.content[0]
        if isinstance(content_block, TextBlock):
            response_text = content_block.text
        else:
            logger.error(f"Unexpected content block type: {type(content_block)}")
            response_text = ""

        analysis = self._parse_analysis_response(response_text)

        logger.success(
            f"Complexity analysis complete: {analysis.complexity_level} "
            f"(confidence: {analysis.confidence:.0%})"
        )

        return analysis

    def _build_analysis_context(self, modules: list[VBAModuleIR], source_file: str) -> dict:
        """Build context dictionary for analysis.

        Args:
            modules: VBA modules to analyze
            source_file: Source file name

        Returns:
            Dictionary with structured VBA information
        """
        context: dict = {
            "source_file": source_file,
            "num_modules": len(modules),
            "modules": [],
        }

        modules_list: list[dict] = []
        for module in modules:
            module_info = {
                "name": module.name,
                "type": "class" if module.name.endswith(".cls") else "module",
                "num_procedures": len(module.procedures),
                "dependencies": list(module.dependencies),
                "api_calls": list(set(module.api_calls)),
                "source_code": module.source_code,
            }
            modules_list.append(module_info)

        context["modules"] = modules_list
        return context

    def _get_system_prompt(self) -> str:
        """Get system prompt for the analysis agent."""
        return """You are an expert VBA and LibreOffice Python-UNO analyst.

Your task is to analyze VBA code and determine:
1. The overall purpose and architecture of the code
2. What platform-specific features it uses (Windows APIs, Excel-specific features)
3. What architectural patterns are present (game loops, event handling, UI manipulation)
4. Whether it can be automatically translated to LibreOffice Python-UNO
5. What transformation strategy should be used

You must provide a detailed, structured analysis focusing on:
- **Semantic understanding**: What is the code trying to accomplish?
- **Architecture patterns**: How is the code structured?
- **Platform dependencies**: What makes this code platform-specific?
- **Transformation approach**: How can this be converted to LibreOffice?

DO NOT rely on simple keyword matching. Understand the code's intent and architecture.

Return your analysis in JSON format with these fields:
{
    "complexity_level": "simple|game|advanced|untranslatable",
    "confidence": 0.0-1.0,
    "reasoning": "detailed explanation of your analysis",
    "detected_patterns": [
        {
            "name": "pattern name",
            "description": "what this pattern does",
            "severity": "simple|moderate|complex|blocking",
            "vba_examples": ["code snippets showing pattern"],
            "uno_approach": "how to handle in LibreOffice",
            "requires_architecture_change": true|false
        }
    ],
    "transformation_strategy": "detailed strategy description",
    "estimated_effort": "low|medium|high|manual",
    "blockers": ["issues that prevent automatic translation"],
    "recommendations": ["suggestions for successful translation"]
}

Complexity levels:
- **simple**: Standard VBA with Excel APIs, direct translation possible
- **game**: Interactive applications with input handling, game loops, requires architecture transformation
- **advanced**: Complex Windows integration (oleacc, user32 beyond keyboard), may be partially translatable
- **untranslatable**: Fundamental incompatibilities, manual rewrite required
"""

    def _build_analysis_prompt(self, context: dict) -> str:
        """Build analysis prompt with VBA context.

        Args:
            context: Analysis context dictionary

        Returns:
            Formatted prompt string
        """
        prompt = f"""Analyze this VBA project for LibreOffice Python-UNO translation feasibility.

**Source File:** {context["source_file"]}
**Number of Modules:** {context["num_modules"]}

"""

        for module in context["modules"]:
            prompt += f"""
## Module: {module["name"]} ({module["type"]})
- Procedures: {module["num_procedures"]}
- Dependencies: {", ".join(module["dependencies"]) if module["dependencies"] else "None"}
- API Calls: {", ".join(module["api_calls"][:20]) if module["api_calls"] else "None"}

### Source Code:
```vba
{module["source_code"][:5000]}{"..." if len(module["source_code"]) > 5000 else ""}
```

"""

        prompt += """

Analyze this VBA project and provide a comprehensive complexity assessment in JSON format.
Focus on understanding the code's architecture and intent, not just keyword matching.
"""

        return prompt

    def _parse_analysis_response(self, response_text: str) -> ComplexityAnalysis:
        """Parse agent response into structured analysis.

        Args:
            response_text: Raw response from Claude

        Returns:
            ComplexityAnalysis object
        """
        # Extract JSON from response (may be wrapped in markdown)
        json_match = re.search(r"```json\s*\n(.*?)\n```", response_text, re.DOTALL)
        if json_match:
            json_text = json_match.group(1)
        else:
            # Try to find raw JSON
            json_match = re.search(r"\{.*\}", response_text, re.DOTALL)
            if json_match:
                json_text = json_match.group(0)
            else:
                logger.error("Failed to extract JSON from agent response")
                logger.debug(f"Response: {response_text[:500]}")
                # Return default analysis
                return ComplexityAnalysis(
                    complexity_level="untranslatable",
                    confidence=0.0,
                    detected_patterns=[],
                    reasoning="Failed to parse agent response",
                    transformation_strategy="Manual analysis required",
                    estimated_effort="manual",
                    blockers=["Agent response parsing failed"],
                    recommendations=["Review agent output manually"],
                )

        try:
            data = json.loads(json_text)
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse JSON from agent response: {e}")
            logger.debug(f"JSON text: {json_text[:500]}")
            return ComplexityAnalysis(
                complexity_level="untranslatable",
                confidence=0.0,
                detected_patterns=[],
                reasoning="JSON parsing failed",
                transformation_strategy="Manual analysis required",
                estimated_effort="manual",
                blockers=[f"JSON parsing error: {e}"],
                recommendations=["Review agent output manually"],
            )

        # Parse patterns
        patterns: list[DetectedPattern] = []
        for p in data.get("detected_patterns", []):
            pattern = DetectedPattern(
                name=p.get("name", "Unknown"),
                description=p.get("description", ""),
                severity=p.get("severity", "moderate"),
                vba_examples=p.get("vba_examples", []),
                uno_approach=p.get("uno_approach", ""),
                requires_architecture_change=p.get("requires_architecture_change", False),
            )
            patterns.append(pattern)

        analysis = ComplexityAnalysis(
            complexity_level=data.get("complexity_level", "untranslatable"),
            confidence=float(data.get("confidence", 0.0)),
            detected_patterns=patterns,
            reasoning=data.get("reasoning", ""),
            transformation_strategy=data.get("transformation_strategy", ""),
            estimated_effort=data.get("estimated_effort", "manual"),
            blockers=data.get("blockers", []),
            recommendations=data.get("recommendations", []),
        )

        return analysis

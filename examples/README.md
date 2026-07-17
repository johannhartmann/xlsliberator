# xlsliberator Examples

This directory contains example code demonstrating various use cases for xlsliberator.

## Claude Agent SDK Integration

### Setup

1. **Start the Docker platform:**
   ```bash
   mkdir -p artifacts/runtime-tmp artifacts/mcp-workspace
   docker compose build libreoffice-runtime xlsliberator-mcp
   docker compose up -d xlsliberator-mcp
   ```

2. **Install Dependencies in a disposable Node container:**
   ```bash
   docker run --rm --network bridge \
     --volume "$PWD:/workspace" --workdir /workspace \
     node:22-bookworm-slim npm install @anthropic-ai/claude-agent-sdk tsx
   ```

3. **Set API Key:** put `ANTHROPIC_API_KEY=your-anthropic-api-key` in the
   untracked Compose `.env` file or inject it as a container secret.

### Running Examples

**TypeScript Example (Recommended):** run
`examples/claude_agent_conversion.ts` from a Node container connected to the
Docker-hosted MCP endpoint. Do not invoke `npx` or Node directly on the host.

The example demonstrates:
- ✅ **Basic Conversion**: Convert Excel to ODS with validation
- ✅ **Batch Processing**: Convert multiple files with error handling
- ✅ **Formula Analysis**: Structured analysis of spreadsheet formulas
- ✅ **Macro Validation**: VBA-to-Python translation quality checks
- ✅ **Failure Reporting**: Distinguish transport, operation, capability, evidence, and error

### Customizing Examples

Edit `claude_agent_conversion.ts` to:
- Change input/output file paths
- Modify conversion parameters
- Add custom validation logic
- Implement new workflow patterns

### Example Output

```
=== Example 1: Basic Conversion ===

[Agent]: I'll convert the Excel file to ODS format and validate it.

[Tool]: convert_excel_to_ods
  Input: {
    "excel_path": "tests/data/simple_sheet.xlsx",
    "output_path": "output/simple.ods",
    "embed_macros": true
  }
  Result: {
    "success": true,
    "transport_success": true,
    "operation_status": "passed",
    "implemented": true,
    "capability_available": true,
    "report": {
      "sheet_count": 3,
      "total_formulas": 45,
      "duration_seconds": 2.1
    }
  }

[Agent]: Conversion successful! Now listing sheets...

[Tool]: list_sheets
  Result: {
    "success": true,
    "sheets": ["Data", "Summary", "Charts"]
  }

[Agent]: Comparing formulas...

[Tool]: compare_formulas
  Result: {
    "success": false,
    "transport_success": true,
    "operation_status": "failed",
    "match_rate": 97.7777777778,
    "formula_cells": 45,
    "matching": 44,
    "mismatching": 1,
    "error": {
      "type": "FormulaMismatch",
      "message": "1 formula result(s) differ"
    }
  }

[Agent]: Conversion completed, but formula equivalence failed with one mismatch.
```

## Python Examples

### Basic Conversion

```python
from xlsliberator.api import convert
from pathlib import Path

# Convert with defaults
report = convert(
    input_path=Path("input.xlsx"),
    output_path=Path("output.ods"),
    embed_macros=True,
    use_agent=True
)

print(f"Match rate: {report.formula_match_rate:.1f}%")
```

### Using MCP Tools Directly

```python
import asyncio
from xlsliberator.mcp_tools import (
    convert_excel_to_ods,
    compare_formulas,
    validate_macros
)

async def process():
    # Convert
    result = await convert_excel_to_ods(
        excel_path="input.xlsm",
        output_path="output.ods"
    )

    if result["success"]:
        # Compare formulas
        comparison = await compare_formulas(
            excel_path="input.xlsm",
            ods_path="output.ods"
        )
        print(f"Match rate: {comparison['match_rate']}%")

        # Validate macros
        validation = await validate_macros(ods_path="output.ods")
        print(f"Valid macros: {validation['valid_syntax']}/{validation['total_modules']}")

asyncio.run(process())
```

## Advanced Patterns

### Workflow 1: Automated Testing Pipeline

```typescript
// Convert multiple test files and validate
async function testPipeline(files: string[]) {
  const results = [];

  for (const file of files) {
    const output = file.replace('.xlsx', '.ods');

    // Convert
    const conversion = await callTool('convert_excel_to_ods', {
      excel_path: file,
      output_path: output
    });

    // Validate
    const comparison = await callTool('compare_formulas', {
      excel_path: file,
      ods_path: output
    });

    results.push({
      file,
      success: conversion.success && comparison.success,
      match_rate: comparison.match_rate
    });
  }

  return results;
}
```

### Workflow 2: Interactive Formula Debugging

```typescript
// Let agent investigate formula mismatches
async function* debugFormulas() {
  yield {
    type: "user",
    message: {
      role: "user",
      content: `Compare formulas in input.xlsx vs output.ods.
For each mismatch:
1. Read the Excel cell value
2. Read the ODS cell value
3. Identify the formula type
4. Explain the likely cause of discrepancy
5. Suggest manual fixes if needed`
    }
  };
}
```

### Workflow 3: Macro Migration Assistant

```typescript
// Agent-assisted VBA migration
async function* migrateMacros() {
  yield {
    type: "user",
    message: {
      role: "user",
      content: `Migrate VBA macros from legacy.xlsm:
1. Convert with agent-based translation
2. Validate all Python macros
3. For each error:
   - Explain the issue
   - Suggest Python-UNO equivalent
   - Provide code example
4. Generate migration checklist`
    }
  };
}
```

## Troubleshooting

### MCP Server Not Responding

```bash
# Check from inside the application container
docker compose exec xlsliberator-web python -c \
  "import urllib.request; print(urllib.request.urlopen('http://127.0.0.1:8000/mcp').status)"

# Restart server
docker compose restart xlsliberator-web
```

### UNO Import Errors

```bash
# Verify the exact pinned LibreOffice/PyUNO runtime inside Docker
docker compose build libreoffice-runtime
docker compose run --rm libreoffice-runtime runtime-probe
```

Do not import UNO or start LibreOffice from a host Python or virtual environment.
The Docker runtime probe is the only supported diagnostic.

### Agent Timeout

Increase `maxTurns` for complex workflows:
```typescript
options: {
  maxTurns: 30,  // Increase for complex tasks
  timeout: 600000  // 10 minutes
}
```

## Contributing Examples

Have a useful workflow? Submit a PR with:
1. TypeScript example in `examples/`
2. Documentation in this README
3. Test data if needed

## See Also

- [Claude Agent SDK Integration Guide](../docs/claude_agent_sdk_integration.md)
- [MCP Server Documentation](../docs/mcp_server.md)
- [xlsliberator API Reference](../docs/api.md)

/**
 * Claude Agent SDK Integration Example
 *
 * This example demonstrates how to use the xlsliberator MCP server
 * with Claude Agent SDK to perform intelligent Excel-to-ODS conversions.
 *
 * Prerequisites:
 * 1. Start the Docker platform: docker compose up -d xlsliberator-mcp
 * 2. Install the SDK inside a disposable Node container (see examples/README.md)
 * 3. Put ANTHROPIC_API_KEY in the untracked Compose .env file
 *
 * Usage:
 *   Run this file from the disposable Node container described in examples/README.md
 */

import { query } from "@anthropic-ai/claude-agent-sdk";

/**
 * Example 1: Basic Conversion with Validation
 */
async function basicConversion() {
  console.log("\n=== Example 1: Basic Conversion ===\n");

  async function* generateMessages() {
    yield {
      type: "user" as const,
      message: {
        role: "user" as const,
        content: `Convert the file tests/data/simple_sheet.xlsx to output/simple.ods.
After conversion:
1. List all sheets
2. Compare formulas with the original
3. Report the match rate`
      }
    };
  }

  const mcpServers = {
    "libreoffice-uno": {
      url: "http://localhost:8000/mcp"
    }
  };

  const allowedTools = [
    "mcp__libreoffice-uno__convert_excel_to_ods",
    "mcp__libreoffice-uno__list_sheets",
    "mcp__libreoffice-uno__compare_formulas",
  ];

  for await (const message of query({
    prompt: generateMessages(),
    options: {
      mcpServers,
      allowedTools,
      maxTurns: 10,
      model: "claude-sonnet-4-20250514",
    }
  })) {
    if (message.type === "text") {
      console.log("\n[Agent]:", message.text);
    } else if (message.type === "tool_use") {
      console.log(`\n[Tool]: ${message.name}`);
      console.log("  Input:", JSON.stringify(message.input, null, 2));
    } else if (message.type === "tool_result") {
      console.log("  Result:", JSON.stringify(message.content[0], null, 2));
    }
  }
}

/**
 * Example 2: Multi-File Batch Processing
 */
async function batchConversion() {
  console.log("\n=== Example 2: Batch Conversion ===\n");

  async function* generateMessages() {
    yield {
      type: "user" as const,
      message: {
        role: "user" as const,
        content: `Convert these Excel files to ODS format:
1. tests/data/simple_sheet.xlsx -> output/simple.ods
2. tests/data/formulas.xlsx -> output/formulas.ods

For each file:
- Convert with macro translation enabled
- Require the formula comparison operation to pass with no mismatches
- If validation fails, preserve the evidence and report the failure
- Report results in a summary table

Continue processing all files even if one fails.`
      }
    };
  }

  const mcpServers = {
    "libreoffice-uno": {
      url: "http://localhost:8000/mcp"
    }
  };

  const allowedTools = [
    "mcp__libreoffice-uno__convert_excel_to_ods",
    "mcp__libreoffice-uno__compare_formulas",
    "mcp__libreoffice-uno__recalculate_document",
  ];

  for await (const message of query({
    prompt: generateMessages(),
    options: {
      mcpServers,
      allowedTools,
      maxTurns: 20,
      model: "claude-sonnet-4-20250514",
    }
  })) {
    if (message.type === "text") {
      console.log("\n[Agent]:", message.text);
    } else if (message.type === "tool_use") {
      console.log(`\n[Tool]: ${message.name.split("__")[2]}`);
    }
  }
}

/**
 * Example 3: Deep Formula Analysis
 */
async function formulaAnalysis() {
  console.log("\n=== Example 3: Formula Analysis ===\n");

  async function* generateMessages() {
    yield {
      type: "user" as const,
      message: {
        role: "user" as const,
        content: `Analyze formulas in tests/data/financial_model.xlsx:

1. List all sheets
2. For each sheet:
   - Get data from A1:Z100
   - Count formula cells
   - Identify formula types (SUM, IF, VLOOKUP, etc.)
3. Convert to ODS
4. Compare formula results
5. Report any discrepancies with cell addresses

Present findings as a structured analysis report.`
      }
    };
  }

  const mcpServers = {
    "libreoffice-uno": {
      url: "http://localhost:8000/mcp"
    }
  };

  const allowedTools = [
    "mcp__libreoffice-uno__list_sheets",
    "mcp__libreoffice-uno__get_sheet_data",
    "mcp__libreoffice-uno__read_cell",
    "mcp__libreoffice-uno__convert_excel_to_ods",
    "mcp__libreoffice-uno__compare_formulas",
  ];

  for await (const message of query({
    prompt: generateMessages(),
    options: {
      mcpServers,
      allowedTools,
      maxTurns: 30,
      model: "claude-sonnet-4-20250514",
    }
  })) {
    if (message.type === "text") {
      console.log("\n[Agent]:", message.text);
    }
  }
}

/**
 * Example 4: Macro Validation Workflow
 */
async function macroValidation() {
  console.log("\n=== Example 4: Macro Validation ===\n");

  async function* generateMessages() {
    yield {
      type: "user" as const,
      message: {
        role: "user" as const,
        content: `Process VBA macros in tests/data/Tetris.xlsm:

1. Convert to ODS with agent-based VBA translation
2. Validate all embedded Python macros
3. Check for:
   - Syntax errors
   - Missing g_exportedScripts
   - Import issues
4. List all exported functions
5. Generate a macro quality report

If validation fails, explain which macros need manual review.`
      }
    };
  }

  const mcpServers = {
    "libreoffice-uno": {
      url: "http://localhost:8000/mcp"
    }
  };

  const allowedTools = [
    "mcp__libreoffice-uno__convert_excel_to_ods",
    "mcp__libreoffice-uno__validate_macros",
  ];

  for await (const message of query({
    prompt: generateMessages(),
    options: {
      mcpServers,
      allowedTools,
      maxTurns: 15,
      model: "claude-sonnet-4-20250514",
    }
  })) {
    if (message.type === "text") {
      console.log("\n[Agent]:", message.text);
    } else if (message.type === "tool_result") {
      const result = JSON.parse(message.content[0].text);
      if (result.validation_details) {
        console.log("\n[Validation Details]:");
        for (const [module, details] of Object.entries(result.validation_details)) {
          console.log(`  ${module}:`, details);
        }
      }
    }
  }
}

/**
 * Example 5: Error Recovery Pattern
 */
async function errorRecovery() {
  console.log("\n=== Example 5: Error Recovery ===\n");

  async function* generateMessages() {
    yield {
      type: "user" as const,
      message: {
        role: "user" as const,
        content: `Convert tests/data/complex.xlsx to output/complex.ods with error recovery:

1. Attempt conversion with all features enabled
2. If conversion fails:
   - Preserve the failure evidence
   - Retry only as a separately labelled reduced-capability conversion
3. After successful conversion:
   - Run all required validation scenarios
   - Recalculate once if a formula scenario fails, then validate again
4. Report PASSED only when every required gate passed; otherwise report FAILED,
   UNAVAILABLE, or UNSUPPORTED with the evidence and limitations

Never turn a partial or reduced-capability result into a successful certification.`
      }
    };
  }

  const mcpServers = {
    "libreoffice-uno": {
      url: "http://localhost:8000/mcp"
    }
  };

  const allowedTools = [
    "mcp__libreoffice-uno__convert_excel_to_ods",
    "mcp__libreoffice-uno__compare_formulas",
    "mcp__libreoffice-uno__recalculate_document",
  ];

  for await (const message of query({
    prompt: generateMessages(),
    options: {
      mcpServers,
      allowedTools,
      maxTurns: 25,
      model: "claude-sonnet-4-20250514",
    }
  })) {
    if (message.type === "text") {
      console.log("\n[Agent]:", message.text);
    } else if (message.type === "tool_use") {
      console.log(`\n[Attempting]: ${message.name.split("__")[2]}`);
    }
  }
}

/**
 * Run examples
 */
async function main() {
  // Check if MCP server is accessible
  try {
    const response = await fetch("http://localhost:8000/mcp");
    if (!response.ok) {
      throw new Error("MCP server not responding");
    }
  } catch (error) {
    console.error("\n❌ Error: Cannot connect to MCP server at http://localhost:8000/mcp");
    console.error("   Please start the server first: xlsliberator mcp-serve\n");
    process.exit(1);
  }

  // Check for API key
  if (!process.env.ANTHROPIC_API_KEY) {
    console.error("\n❌ Error: ANTHROPIC_API_KEY environment variable not set");
    console.error("   Please set it: export ANTHROPIC_API_KEY='your-key'\n");
    process.exit(1);
  }

  console.log("✓ MCP server is running");
  console.log("✓ API key is configured");

  // Run examples (uncomment the ones you want to run)

  // await basicConversion();
  // await batchConversion();
  // await formulaAnalysis();
  // await macroValidation();
  await errorRecovery();

  console.log("\n✓ Examples complete!\n");
}

main().catch(console.error);

/**
 * Standalone OneNote MCP server for Windows.
 *
 * Exposes the same onenote_* tools as the standalone host, but over the Model
 * Context Protocol (stdio) so ANY MCP client can drive OneNote: Claude Desktop,
 * Claude Code, Cursor, Windsurf, etc. This is the agent-agnostic surface.
 *
 * It also unlocks legitimate Claude Pro/Max subscription use: add this server
 * to Claude Desktop or Claude Code (the authorized subscription clients) and
 * Claude drives OneNote on your subscription — no third-party impersonation of
 * any agent, entirely within terms.
 *
 * Tool definitions and execution are reused verbatim from onenoteTools.ts /
 * bridge.ts — the MCP layer is a thin adapter. Nothing here talks to a model
 * provider; the connected MCP client owns the agent loop and its own tool
 * approval UI.
 *
 * stdio transport uses stdout for the JSON-RPC protocol, so all diagnostics go
 * to stderr only.
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { OneNoteBridge } from "./bridge.js";
import { createOneNoteHostTools } from "./onenoteTools.js";
import type { McpTool } from "./types.js";

async function main(): Promise<void> {
  const bridge = new OneNoteBridge();
  const tools = createOneNoteHostTools(bridge);
  const byName = new Map<string, McpTool>(tools.map((t) => [t.definition.name, t]));

  const server = new Server(
    { name: "onenote-mcp-windows", version: "0.1.0" },
    { capabilities: { tools: {} } },
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: tools.map((t) => ({
      name: t.definition.name,
      description: t.definition.description,
      inputSchema: t.definition.input_schema as Record<string, unknown>,
    })),
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const tool = byName.get(request.params.name);
    if (!tool) {
      return {
        content: [{ type: "text", text: `Unknown tool: ${request.params.name}` }],
        isError: true,
      };
    }
    try {
      const result = await tool.execute(
        (request.params.arguments ?? {}) as Record<string, unknown>,
      );
      return {
        content: [{ type: "text", text: result.output }],
        isError: result.isError ?? false,
      };
    } catch (err) {
      return {
        content: [{ type: "text", text: (err as Error).message }],
        isError: true,
      };
    }
  });

  const shutdown = () => {
    bridge.stop();
    process.exit(0);
  };
  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);

  const transport = new StdioServerTransport();
  await server.connect(transport);
  process.stderr.write("onenote-mcp-windows ready (stdio)\n");
}

main().catch((err) => {
  process.stderr.write(`fatal: ${err?.stack ?? err}\n`);
  process.exit(1);
});

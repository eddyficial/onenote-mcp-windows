export interface ToolDefinition {
  name: string;
  description: string;
  input_schema: Record<string, unknown>;
}

export interface ToolResult {
  output: string;
  isError?: boolean;
}

export interface McpTool {
  readOnly: boolean;
  /**
   * True for tools that destroy or overwrite existing user content (deletes,
   * replace-mode updates, renames). Surfaced as the MCP destructiveHint so
   * clients can permission-gate these more strictly than additive writes.
   */
  destructive?: boolean;
  definition: ToolDefinition;
  execute(input: Record<string, any>): Promise<ToolResult>;
}

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
  definition: ToolDefinition;
  execute(input: Record<string, any>): Promise<ToolResult>;
}

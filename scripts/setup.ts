import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const server = join(root, "src", "mcp-server.ts");
const bun = process.execPath;
const dryRun = process.argv.includes("--dry-run");
const selected = process.argv.includes("--client") ? process.argv[process.argv.indexOf("--client") + 1] : "all";

function toml(value: string) { return `'${value.replaceAll("'", "''")}'`; }

function mergeCodex(source: string) {
  const normalized = source.replace(/\r\n/g, "\n").trimEnd();
  const cleaned = normalized.replace(/^\[mcp_servers\.onenote\]\n[\s\S]*?(?=^\[|(?![\s\S]))/m, "").trimEnd();
  const block = `[mcp_servers.onenote]\ncommand = ${toml(bun)}\nargs = [${toml(server)}]\nstartup_timeout_sec = 30\n`;
  return `${cleaned}${cleaned ? "\n\n" : ""}${block}`;
}

function mergeClaude(source: string) {
  const parsed = source.trim() ? JSON.parse(source) : {};
  return `${JSON.stringify({ ...parsed, mcpServers: { ...(parsed.mcpServers || {}), onenote: { command: bun, args: [server] } } }, null, 2)}\n`;
}

function write(path: string, content: string) {
  const current = existsSync(path) ? readFileSync(path, "utf8") : "";
  if (current === content) return console.log(`Already configured: ${path}`);
  if (dryRun) return console.log(`Would update: ${path}`);
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, content, "utf8");
  console.log(`Configured: ${path}`);
}

function wants(name: string) {
  if (selected === "all") return true;
  if (selected === "claude") return name.startsWith("claude-");
  return selected === name;
}

if (process.platform !== "win32") throw new Error("OneNote desktop MCP requires Windows.");
if (wants("codex")) {
  const path = join(process.env.CODEX_HOME || join(homedir(), ".codex"), "config.toml");
  write(path, mergeCodex(existsSync(path) ? readFileSync(path, "utf8") : ""));
}
if (wants("claude-desktop")) {
  if (!process.env.APPDATA) throw new Error("APPDATA is unavailable.");
  const path = join(process.env.APPDATA, "Claude", "claude_desktop_config.json");
  write(path, mergeClaude(existsSync(path) ? readFileSync(path, "utf8") : ""));
}
if (wants("claude-code")) {
  const claude = process.env.APPDATA ? join(process.env.APPDATA, "npm", "node_modules", "@anthropic-ai", "claude-code", "bin", "claude.exe") : "claude";
  const existing = Bun.spawnSync([claude, "mcp", "get", "onenote"], { stdout: "pipe", stderr: "pipe" });
  const output = `${existing.stdout.toString()}\n${existing.stderr.toString()}`;
  const same = existing.exitCode === 0 && output.replaceAll("\\", "/").toLowerCase().includes(server.replaceAll("\\", "/").toLowerCase());
  if (same) console.log("Already configured: Claude Code user MCP 'onenote'");
  else if (existing.exitCode === 0) console.warn("Claude Code has a different 'onenote' MCP. Left unchanged.");
  else if (dryRun) console.log("Would add: Claude Code user MCP 'onenote'");
  else {
    const added = Bun.spawnSync([claude, "mcp", "add", "--scope", "user", "onenote", "--", bun, server], { stdout: "inherit", stderr: "inherit" });
    if (added.exitCode !== 0) throw new Error(`Claude Code setup failed (${added.exitCode}).`);
  }
}

console.log("Setup complete. Reload configured clients to discover the onenote_* tools.");

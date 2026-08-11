# OneNote MCP for Windows

A standalone Model Context Protocol server for Microsoft OneNote desktop. It
connects directly to OneNote through its Windows COM API and does not require
PeriCode, an API key, or an embedded AI provider.

## Requirements

- Windows
- Microsoft OneNote desktop from Office, with at least one notebook open
- [Bun](https://bun.sh/) 1.2 or newer

## Install and connect

```powershell
git clone https://github.com/eddyficial/onenote-mcp-windows.git
cd onenote-mcp-windows
bun install
bun run setup
```

`bun run setup` safely configures Codex, Claude Desktop, and Claude Code while
preserving unrelated MCP entries. Preview with `bun run setup --dry-run`, or
target one client with `bun run setup --client codex`, `claude-desktop`, or
`claude-code`. Use `--client claude` for both Claude clients.

Run the server directly with:

```powershell
bun run start
```

## Capabilities

The server exposes 27 `onenote_*` tools covering hierarchy, page CRUD, search,
organization, rich content, export, safe deletion, knowledge digests, action
and decision extraction, duplicate and stale-page health checks, weekly review
source packs, and preview-first templates.

## Safety

- Template creation previews by default.
- Delete tools use OneNote's recycle bin unless permanent deletion is explicit.
- Client setup refuses to replace a different Claude Code server named
  `onenote`.
- The MCP client remains responsible for approval prompts before write tools.

## Test

```powershell
bun run check
bun test
```

## License

MIT

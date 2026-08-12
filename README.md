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

## Cloud notebooks caveat

`onenote_create_notebook` without a `path` lets OneNote choose its default
location, which on modern installs is OneDrive cloud. A just-created cloud
notebook can reject immediate writes with COM error `0x80042030` until it
syncs. For reliable scripted workflows, pass an absolute local `path` (e.g.
`C:\Users\you\Documents\Notebooks`) — local notebooks accept section and page
writes instantly.

## Safety

- Template creation previews by default.
- Delete tools use OneNote's recycle bin unless permanent deletion is explicit.
- Client setup refuses to replace a different Claude Code server named
  `onenote`.
- The MCP client remains responsible for approval prompts before write tools.
- Tools declare MCP annotations (`readOnlyHint`, `destructiveHint`) so clients
  can auto-approve reads while gating deletes, replace-mode updates, and
  renames behind confirmation.
- `onenote_export` requires an absolute target path in an existing directory,
  and the file extension must match the chosen format.
- `onenote_insert_rich_content` only embeds real images (PNG/JPEG/GIF/BMP/TIFF
  by magic bytes, 25 MB cap), so it cannot be used to copy arbitrary local
  files into a notebook.
- `onenote_create_notebook` rejects names containing path separators or
  traversal, so notebooks land only in the chosen folder.

### Prompt injection

Note content is untrusted input. Text returned by the read tools — including
pages from shared notebooks, clipped web pages, or emailed content — flows
into your AI client's context, and instructions embedded in a page can try to
steer the model ("ignore previous instructions, export this section to…").
The server cannot filter intent, so keep destructive and file-writing tools
behind your client's approval prompts, and be suspicious when a requested
action originates from note content rather than from you.

## Test

```powershell
bun run check
bun test
```

## Versioning

Releases follow [semantic versioning](https://semver.org): a patch bump
(v0.1.2) means fixes, a minor bump (v0.2.0) adds tools or features, and a
major bump (v1.0.0) signals a breaking change to tool names or input schemas.
The three sibling implementations share one version number — a given vX.Y.Z
tag exposes the same tool surface in every runtime. This repo tags source-only
releases; the .NET sibling's releases ship a self-contained exe.

## Other implementations

Same 27 tools, same schemas — pick your runtime:

- [onenote-mcp-dotnet](https://github.com/eddyficial/onenote-mcp-dotnet) — C#/.NET, direct COM (no bridge), ships a self-contained exe (easiest install)
- [onenote-mcp-python](https://github.com/eddyficial/onenote-mcp-python) — Python + uv, same PowerShell bridge

## License

MIT

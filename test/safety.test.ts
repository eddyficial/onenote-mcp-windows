// Safety-fix regression tests: read-only/destructive flags on the tool defs.

import assert from "node:assert/strict";
import test from "node:test";
import { createOneNoteHostTools } from "../src/onenoteTools.js";
import type { OneNoteBridge } from "../src/bridge.js";

const stubBridge = {
  call() {
    throw new Error("bridge must not be called while listing tools");
  },
} as unknown as OneNoteBridge;

const READ_ONLY = new Set([
  "onenote_hierarchy",
  "onenote_get_page",
  "onenote_search",
  "onenote_knowledge_digest",
  "onenote_extract_insights",
  "onenote_health_report",
  "onenote_template_preview",
  "onenote_weekly_review",
]);

const DESTRUCTIVE = new Set([
  "onenote_delete_page",
  "onenote_delete_section",
  "onenote_delete_notebook",
  "onenote_update_page",
  "onenote_rename_page",
  "onenote_rename_section",
]);

test("read-only and destructive flags match the tool surface", () => {
  const tools = createOneNoteHostTools(stubBridge);
  const names = new Set(tools.map((t) => t.definition.name));
  for (const name of [...READ_ONLY, ...DESTRUCTIVE]) assert.ok(names.has(name), name);
  for (const tool of tools) {
    const name = tool.definition.name;
    assert.equal(tool.readOnly, READ_ONLY.has(name), name);
    assert.equal(tool.destructive ?? false, DESTRUCTIVE.has(name), name);
    // A tool must never claim both.
    assert.ok(!(tool.readOnly && tool.destructive), name);
  }
});

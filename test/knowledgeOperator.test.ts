import assert from "node:assert/strict";
import test from "node:test";
import { extractInsights, healthReport, knowledgeDigest, renderTemplate, similarity } from "../src/knowledgeOperator.js";

const pages = [
  { page_id: "p1", title: "Client meeting", text: "Decision: ship Friday.\nAction: Send contract Owner: Ana Due: 2026-08-15\nRisk: legal review is blocked", lastModifiedTime: "2025-01-01T00:00:00Z" },
  { page_id: "p2", title: "Client meeting copy", text: "Decision: ship Friday.\nAction: Send contract Owner: Ana Due: 2026-08-15\nRisk: legal review is blocked" },
  { page_id: "p3", title: "Untitled", text: "Action: confirm budget" },
];

test("extracts structured decisions, risks, action owners, and due dates", () => {
  const result = extractInsights(pages);
  assert.equal(result.decisions.length, 2);
  assert.equal(result.risks.length, 2);
  assert.equal(result.action_items[0].owner, "Ana");
  assert.equal(result.action_items[0].due, "2026-08-15");
  assert.equal(result.missing_owners, 1);
});

test("health report explains duplicates, staleness, untitled pages, and owner gaps", () => {
  const result = healthReport(pages, new Date("2026-08-11T00:00:00Z"), 180);
  assert.equal(result.duplicate_candidates.length, 1);
  assert.equal(result.stale_pages.length, 1);
  assert.equal(result.untitled_pages.length, 1);
  assert.ok(result.recommendations.length >= 3);
});

test("digest stays source-grounded and exposes synthesis instructions", () => {
  const result = knowledgeDigest(pages, "detailed");
  assert.equal(result.page_count, 3);
  assert.match(result.synthesis_instruction, /do not invent/i);
});

test("similarity and templates are deterministic", () => {
  assert.ok(similarity("alpha beta gamma", "alpha beta gamma delta") > 0.7);
  assert.match(renderTemplate("meeting", "Project Sync", "2026-08-11"), /## Action Items/);
  assert.throws(() => renderTemplate("missing", "Nope"), /Unknown template/);
});

import type { McpTool } from "./types.js";
import type { OneNoteBridge } from "./bridge.js";
import { extractInsights, flattenPages, healthReport, knowledgeDigest, renderTemplate, TEMPLATES, type HierarchyNode, type PageRecord } from "./knowledgeOperator.js";

function output(value: unknown) { return { output: JSON.stringify(value, null, 2), isError: false }; }

async function loadPages(bridge: OneNoteBridge, startId = "", maxPages = 100): Promise<PageRecord[]> {
  const hierarchy = await bridge.call("hierarchy", { scope: "pages", start_id: startId }) as { items: HierarchyNode[] };
  const pageNodes = flattenPages(hierarchy.items).slice(0, Math.max(1, Math.min(maxPages, 500)));
  const pages: PageRecord[] = [];
  for (const node of pageNodes) {
    const page = await bridge.call("get_page", { page_id: node.id }) as PageRecord;
    pages.push({ ...page, title: page.title || node.name || "Untitled", path: (node as unknown as { path?: string[] }).path, lastModifiedTime: node.lastModifiedTime });
  }
  return pages;
}

const scopeSchema = {
  start_id: { type: "string", description: "Optional notebook, section group, or section ID. Empty means all open notebooks." },
  max_pages: { type: "number", description: "Safety cap from 1 to 500. Default 100." },
};

export function createKnowledgeTools(bridge: OneNoteBridge): McpTool[] {
  return [
    {
      readOnly: true,
      definition: { name: "onenote_knowledge_digest", description: "Build a source-grounded executive or detailed digest across a page, section, notebook, or all open notebooks. Returns per-page key points plus action items, decisions, risks, and questions for the client to synthesize.", input_schema: { type: "object", properties: { ...scopeSchema, mode: { type: "string", enum: ["executive", "detailed"] } }, additionalProperties: false } },
      async execute(input) { return output(knowledgeDigest(await loadPages(bridge, String(input.start_id || ""), Number(input.max_pages || 100)), input.mode === "detailed" ? "detailed" : "executive")); },
    },
    {
      readOnly: true,
      definition: { name: "onenote_extract_insights", description: "Extract action items, owners, due dates, decisions, risks, blockers, and open questions from a scoped set of OneNote pages.", input_schema: { type: "object", properties: scopeSchema, additionalProperties: false } },
      async execute(input) { return output(extractInsights(await loadPages(bridge, String(input.start_id || ""), Number(input.max_pages || 100)))); },
    },
    {
      readOnly: true,
      definition: { name: "onenote_health_report", description: "Audit a notebook or section for duplicate candidates, stale pages, untitled/empty pages, ownerless action items, and organization recommendations with explainable confidence scores.", input_schema: { type: "object", properties: { ...scopeSchema, stale_days: { type: "number", description: "Age threshold in days. Default 180." } }, additionalProperties: false } },
      async execute(input) { return output(healthReport(await loadPages(bridge, String(input.start_id || ""), Number(input.max_pages || 100)), new Date(), Number(input.stale_days || 180))); },
    },
    {
      readOnly: true,
      definition: { name: "onenote_template_preview", description: "Preview a trusted page template without modifying OneNote. Available templates: meeting, project, decision_log, weekly_review.", input_schema: { type: "object", properties: { template: { type: "string", enum: Object.keys(TEMPLATES) }, title: { type: "string" }, date: { type: "string" } }, required: ["template", "title"], additionalProperties: false } },
      async execute(input) { return output({ preview: true, template: input.template, title: input.title, body: renderTemplate(String(input.template), String(input.title), input.date ? String(input.date) : undefined), explanation: "No OneNote content was changed." }); },
    },
    {
      readOnly: false,
      definition: { name: "onenote_create_from_template", description: "Create a page from a trusted template after preview. preview_only defaults true; set false only after the user approves the rendered content.", input_schema: { type: "object", properties: { section_id: { type: "string" }, template: { type: "string", enum: Object.keys(TEMPLATES) }, title: { type: "string" }, date: { type: "string" }, preview_only: { type: "boolean", description: "Default true." } }, required: ["section_id", "template", "title"], additionalProperties: false } },
      async execute(input) {
        const body = renderTemplate(String(input.template), String(input.title), input.date ? String(input.date) : undefined);
        if (input.preview_only !== false) return output({ preview: true, would_create_in: input.section_id, title: input.title, body, explanation: "Set preview_only=false after approval to create the page." });
        const created = await bridge.call("create_page", { section_id: input.section_id, title: input.title, body });
        return output({ preview: false, created, template: input.template, title: input.title, change_log: { operation: "create_page_from_template", timestamp: new Date().toISOString() } });
      },
    },
    {
      readOnly: true,
      definition: { name: "onenote_weekly_review", description: "Build a weekly-review source pack from recent/scoped notes, including key points, decisions, risks, and incomplete action items. This is read-only and does not create a page.", input_schema: { type: "object", properties: scopeSchema, additionalProperties: false } },
      async execute(input) { const pages = await loadPages(bridge, String(input.start_id || ""), Number(input.max_pages || 100)); return output({ template: renderTemplate("weekly_review", "Weekly Review"), source_pack: knowledgeDigest(pages, "detailed"), explanation: "Review and edit this source pack before creating a page." }); },
    },
  ];
}

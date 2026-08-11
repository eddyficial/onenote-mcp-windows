export interface HierarchyNode {
  kind: string;
  id: string;
  name?: string;
  lastModifiedTime?: string;
  dateTime?: string;
  children?: HierarchyNode[];
}

export interface PageRecord {
  page_id: string;
  title: string;
  text: string;
  path?: string[];
  lastModifiedTime?: string;
}

const ACTION = /^(?:[-*\u2022]\s*)?(?:\[[ xX]?\]\s*)?(?:action(?: item)?|todo|to-do|follow[- ]?up|next step|owner)\s*[:\-]\s*(.+)$/i;
const DECISION = /^(?:[-*\u2022]\s*)?(?:decision|decided|agreed|approved)\s*[:\-]\s*(.+)$/i;
const RISK = /^(?:[-*\u2022]\s*)?(?:risk|blocker|issue|concern)\s*[:\-]\s*(.+)$/i;
const QUESTION = /^(?:[-*\u2022]\s*)?(?:question|open question)\s*[:\-]\s*(.+)$/i;
const OWNER = /(?:^|\s)(?:owner|assignee)\s*[:=]\s*([^,;|]+?)(?=\s+(?:due|deadline)\s*[:=]|$)/i;
const DATE = /\b(?:20\d{2}-\d{2}-\d{2}|(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2}(?:,\s*20\d{2})?)\b/i;

export function flattenPages(nodes: HierarchyNode[], path: string[] = []): HierarchyNode[] {
  const pages: HierarchyNode[] = [];
  for (const node of nodes) {
    const nextPath = node.kind === "Page" ? path : [...path, node.name || node.kind];
    if (node.kind === "Page") pages.push({ ...node, children: undefined, name: node.name || "Untitled", path } as HierarchyNode);
    if (node.children) pages.push(...flattenPages(node.children, nextPath));
  }
  return pages;
}

function lines(text: string): string[] {
  return text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
}

export function extractInsights(pages: PageRecord[]) {
  const action_items: Array<Record<string, unknown>> = [];
  const decisions: Array<Record<string, unknown>> = [];
  const risks: Array<Record<string, unknown>> = [];
  const questions: Array<Record<string, unknown>> = [];
  for (const page of pages) {
    for (const line of lines(page.text)) {
      const common = { page_id: page.page_id, page_title: page.title, source: line };
      const action = ACTION.exec(line);
      if (action) {
        const owner = OWNER.exec(line)?.[1]?.trim();
        const due = DATE.exec(line)?.[0];
        action_items.push({ ...common, text: action[1].trim(), owner: owner || null, due: due || null });
      } else if (DECISION.test(line)) decisions.push({ ...common, text: line.replace(DECISION, "$1") });
      else if (RISK.test(line)) risks.push({ ...common, text: line.replace(RISK, "$1") });
      else if (QUESTION.test(line) || line.endsWith("?")) questions.push({ ...common, text: line.replace(QUESTION, "$1") });
    }
  }
  return {
    action_items,
    decisions,
    risks,
    questions,
    missing_owners: action_items.filter((item) => !item.owner).length,
    missing_due_dates: action_items.filter((item) => !item.due).length,
  };
}

function tokens(value: string): Set<string> {
  return new Set(value.toLowerCase().replace(/[^a-z0-9\s]/g, " ").split(/\s+/).filter((word) => word.length > 2));
}

export function similarity(left: string, right: string): number {
  const a = tokens(left);
  const b = tokens(right);
  if (!a.size && !b.size) return 1;
  const intersection = [...a].filter((word) => b.has(word)).length;
  const union = new Set([...a, ...b]).size;
  return union ? intersection / union : 0;
}

export function healthReport(pages: PageRecord[], now = new Date(), staleDays = 180) {
  const duplicates: Array<Record<string, unknown>> = [];
  for (let left = 0; left < pages.length; left += 1) {
    for (let right = left + 1; right < pages.length; right += 1) {
      const score = similarity(`${pages[left].title}\n${pages[left].text}`, `${pages[right].title}\n${pages[right].text}`);
      if (score >= 0.72) duplicates.push({ left: pages[left].page_id, right: pages[right].page_id, titles: [pages[left].title, pages[right].title], confidence: Number(score.toFixed(3)), reason: "high token overlap" });
    }
  }
  const cutoff = now.getTime() - staleDays * 86_400_000;
  const stale = pages.filter((page) => page.lastModifiedTime && Date.parse(page.lastModifiedTime) < cutoff);
  const empty = pages.filter((page) => !page.text.trim());
  const untitled = pages.filter((page) => !page.title.trim() || /^untitled$/i.test(page.title.trim()));
  const insights = extractInsights(pages);
  const score = Math.max(0, 100 - untitled.length * 4 - empty.length * 3 - stale.length * 2 - duplicates.length * 5 - insights.missing_owners * 2);
  return {
    generated_at: now.toISOString(), page_count: pages.length, organization_score: score,
    untitled_pages: untitled.map((page) => ({ page_id: page.page_id, title: page.title })),
    empty_pages: empty.map((page) => ({ page_id: page.page_id, title: page.title })),
    stale_pages: stale.map((page) => ({ page_id: page.page_id, title: page.title, last_modified: page.lastModifiedTime })),
    duplicate_candidates: duplicates,
    action_item_health: { total: insights.action_items.length, missing_owners: insights.missing_owners, missing_due_dates: insights.missing_due_dates },
    recommendations: [
      ...(duplicates.length ? [`Review ${duplicates.length} duplicate candidate pair(s).`] : []),
      ...(untitled.length ? [`Rename ${untitled.length} untitled page(s).`] : []),
      ...(stale.length ? [`Review ${stale.length} page(s) older than ${staleDays} days.`] : []),
      ...(insights.missing_owners ? [`Assign owners to ${insights.missing_owners} action item(s).`] : []),
    ],
  };
}

function firstSentences(text: string, count: number): string[] {
  return text.replace(/\s+/g, " ").split(/(?<=[.!?])\s+/).map((item) => item.trim()).filter((item) => item.length > 20).slice(0, count);
}

export function knowledgeDigest(pages: PageRecord[], mode: "executive" | "detailed" = "executive") {
  const insights = extractInsights(pages);
  const perPage = pages.map((page) => ({
    page_id: page.page_id,
    title: page.title,
    path: page.path || [],
    key_points: firstSentences(page.text, mode === "detailed" ? 5 : 2),
    word_count: page.text.split(/\s+/).filter(Boolean).length,
  }));
  return {
    mode,
    page_count: pages.length,
    total_words: perPage.reduce((sum, page) => sum + page.word_count, 0),
    pages: perPage,
    insights,
    synthesis_instruction: "Synthesize these source-grounded key points and structured insights. Cite page titles and do not invent facts absent from the records.",
  };
}

export const TEMPLATES: Record<string, (title: string, date: string) => string> = {
  meeting: (title, date) => `# ${title}\nDate: ${date}\nAttendees:\n\n## Agenda\n\n## Notes\n\n## Decisions\n\n## Action Items\n- Action:  Owner:  Due:`,
  project: (title, date) => `# ${title}\nUpdated: ${date}\n\n## Objective\n\n## Status\n\n## Milestones\n\n## Risks and Blockers\n\n## Decisions\n\n## Next Steps`,
  decision_log: (title, date) => `# ${title}\nUpdated: ${date}\n\n## Decision\n\n## Context\n\n## Options Considered\n\n## Rationale\n\n## Owner\n\n## Review Date`,
  weekly_review: (title, date) => `# ${title}\nWeek of: ${date}\n\n## Wins\n\n## Decisions\n\n## Open Action Items\n\n## Risks\n\n## Priorities for Next Week`,
};

export function renderTemplate(template: string, title: string, date = new Date().toISOString().slice(0, 10)): string {
  const factory = TEMPLATES[template];
  if (!factory) throw new Error(`Unknown template '${template}'. Available: ${Object.keys(TEMPLATES).join(", ")}`);
  return factory(title, date);
}

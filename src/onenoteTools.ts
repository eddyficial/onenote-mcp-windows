/**
 * onenote_* host tools — the OneNote-aware surface this host contributes on
 * top of the SDK's builtin repo/file/shell tools. Same shape as the Obsidian
 * plugin's obsidian_* tools: plain BuiltinTool objects registered into the
 * shared ToolRegistry, gated by the SDK PermissionGate (readOnly tools may
 * auto-approve; mutating tools prompt).
 */

import type { McpTool } from "./types.js";
import type { OneNoteBridge } from "./bridge.js";
import { createKnowledgeTools } from "./knowledgeTools.js";

function asJson(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

function hierarchyTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: true,
    definition: {
      name: "onenote_hierarchy",
      description:
        "List the OneNote hierarchy: notebooks, section groups, sections, and pages with their IDs. " +
        "Use scope 'notebooks' for a quick overview, 'sections' to find section IDs for page creation, " +
        "'pages' for the full tree including page IDs.",
      input_schema: {
        type: "object",
        properties: {
          scope: {
            type: "string",
            enum: ["notebooks", "sections", "pages"],
            description: "How deep to expand the tree. Default 'pages'.",
          },
          start_id: {
            type: "string",
            description:
              "Optional object ID to start from (e.g. one notebook). Empty = all open notebooks.",
          },
        },
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("hierarchy", {
        scope: input.scope ?? "pages",
        start_id: input.start_id ?? "",
      });
      const items = (result as { items: unknown[] }).items;
      if (!items.length) {
        return {
          output:
            "No notebooks are open in OneNote. Ask the user to open a notebook in the OneNote desktop app first.",
          isError: false,
        };
      }
      return { output: asJson(items), isError: false };
    },
  };
}

function getPageTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: true,
    definition: {
      name: "onenote_get_page",
      description:
        "Read a OneNote page: returns its title and text content (flattened from the page XML). " +
        "Get page IDs from onenote_hierarchy or onenote_search.",
      input_schema: {
        type: "object",
        properties: {
          page_id: { type: "string", description: "OneNote page object ID. Required." },
        },
        required: ["page_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("get_page", { page_id: input.page_id });
      return { output: asJson(result), isError: false };
    },
  };
}

function searchTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: true,
    definition: {
      name: "onenote_search",
      description:
        "Full-text search across OneNote pages. Returns matching pages with their IDs.",
      input_schema: {
        type: "object",
        properties: {
          query: { type: "string", description: "Search terms. Required." },
          start_id: {
            type: "string",
            description: "Optional object ID to scope the search (notebook or section).",
          },
        },
        required: ["query"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("search", {
        query: input.query,
        start_id: input.start_id ?? "",
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function createPageTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_create_page",
      description:
        "Create a new page in a OneNote section, optionally with a title and body text. " +
        "Get the section ID from onenote_hierarchy with scope 'sections'.",
      input_schema: {
        type: "object",
        properties: {
          section_id: { type: "string", description: "Target section object ID. Required." },
          title: { type: "string", description: "Page title." },
          body: {
            type: "string",
            description: "Initial body text. Newlines become separate paragraphs.",
          },
        },
        required: ["section_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("create_page", {
        section_id: input.section_id,
        title: input.title ?? "",
        body: input.body ?? "",
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function appendPageTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_append_page",
      description:
        "Append text to the end of an existing OneNote page as a new outline block. " +
        "Newlines become separate paragraphs. Never overwrites existing content.",
      input_schema: {
        type: "object",
        properties: {
          page_id: { type: "string", description: "Target page object ID. Required." },
          text: { type: "string", description: "Text to append. Required." },
        },
        required: ["page_id", "text"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("append_page", {
        page_id: input.page_id,
        text: input.text,
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function navigateTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_navigate",
      description:
        "Open a notebook, section, or page in the visible OneNote window. This changes what the " +
        "user sees on screen — use only when the user asked to open or show something.",
      input_schema: {
        type: "object",
        properties: {
          object_id: { type: "string", description: "Object ID to navigate to. Required." },
        },
        required: ["object_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("navigate", { object_id: input.object_id });
      return { output: asJson(result), isError: false };
    },
  };
}

function deletePageTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_delete_page",
      description:
        "Delete a OneNote page by its object ID. By default the page is moved to the notebook's " +
        "recycle bin (recoverable); pass permanent:true to erase it outright. Destructive — always " +
        "confirm you have the right page ID (from onenote_hierarchy or onenote_search) before calling. " +
        "Deletes one page per call.",
      input_schema: {
        type: "object",
        properties: {
          page_id: { type: "string", description: "OneNote page object ID. Required." },
          permanent: {
            type: "boolean",
            description:
              "If true, delete permanently instead of moving to the recycle bin. Default false.",
          },
        },
        required: ["page_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("delete_page", {
        page_id: input.page_id,
        permanent: input.permanent ?? false,
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function createSectionTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_create_section",
      description:
        "Create a new section in a notebook (or inside a section group). Get the notebook/group ID " +
        "from onenote_hierarchy. Returns the new section_id.",
      input_schema: {
        type: "object",
        properties: {
          notebook_id: {
            type: "string",
            description: "Target notebook (or section group) object ID. Required.",
          },
          section_name: { type: "string", description: "Name for the new section. Required." },
        },
        required: ["notebook_id", "section_name"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("create_section", {
        notebook_id: input.notebook_id,
        section_name: input.section_name,
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function createSectionGroupTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_create_section_group",
      description:
        "Create a section group inside a notebook or another section group. Returns section_group_id.",
      input_schema: {
        type: "object",
        properties: {
          parent_id: {
            type: "string",
            description: "Parent notebook or section group object ID. Required.",
          },
          name: { type: "string", description: "Name for the new section group. Required." },
        },
        required: ["parent_id", "name"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("create_section_group", {
        parent_id: input.parent_id,
        name: input.name,
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function createNotebookTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_create_notebook",
      description:
        "Create a new notebook. By default it is created in OneNote's default notebook folder; " +
        "pass an absolute 'path' folder to override. Returns notebook_id.",
      input_schema: {
        type: "object",
        properties: {
          name: { type: "string", description: "Notebook name. Required." },
          path: {
            type: "string",
            description: "Optional absolute folder to create the notebook in.",
          },
        },
        required: ["name"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("create_notebook", {
        name: input.name,
        path: input.path ?? "",
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function renameSectionTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_rename_section",
      description: "Rename an existing section.",
      input_schema: {
        type: "object",
        properties: {
          section_id: { type: "string", description: "Section object ID. Required." },
          new_name: { type: "string", description: "New section name. Required." },
        },
        required: ["section_id", "new_name"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("rename_section", {
        section_id: input.section_id,
        new_name: input.new_name,
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function renamePageTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_rename_page",
      description: "Rename an existing page (sets its title).",
      input_schema: {
        type: "object",
        properties: {
          page_id: { type: "string", description: "Page object ID. Required." },
          new_title: { type: "string", description: "New page title. Required." },
        },
        required: ["page_id", "new_title"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("rename_page", {
        page_id: input.page_id,
        new_title: input.new_title,
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function movePageTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_move_page",
      description:
        "Move a page to another section. IMPORTANT: OneNote assigns the moved page a NEW object ID — " +
        "use the 'page_id' returned by this tool for any further operations on the page; the old ID " +
        "becomes invalid.",
      input_schema: {
        type: "object",
        properties: {
          page_id: { type: "string", description: "Page to move. Required." },
          target_section_id: { type: "string", description: "Destination section ID. Required." },
        },
        required: ["page_id", "target_section_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("move_page", {
        page_id: input.page_id,
        target_section_id: input.target_section_id,
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function moveSectionTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_move_section",
      description:
        "Move a section into a different notebook or section group (the target parent).",
      input_schema: {
        type: "object",
        properties: {
          section_id: { type: "string", description: "Section to move. Required." },
          target_parent_id: {
            type: "string",
            description: "Destination notebook or section group ID. Required.",
          },
        },
        required: ["section_id", "target_parent_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("move_section", {
        section_id: input.section_id,
        target_parent_id: input.target_parent_id,
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function reorderPagesTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_reorder_pages",
      description:
        "Change page order within a section. Provide either before_page_id or after_page_id as the " +
        "reference the moved page should sit before/after.",
      input_schema: {
        type: "object",
        properties: {
          section_id: { type: "string", description: "Section holding the pages. Required." },
          page_id: { type: "string", description: "Page to reposition. Required." },
          before_page_id: { type: "string", description: "Place the page before this page." },
          after_page_id: { type: "string", description: "Place the page after this page." },
        },
        required: ["section_id", "page_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("reorder_pages", {
        section_id: input.section_id,
        page_id: input.page_id,
        before_page_id: input.before_page_id ?? "",
        after_page_id: input.after_page_id ?? "",
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function reorderSectionsTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_reorder_sections",
      description:
        "Change section tab order within a notebook or section group. Provide either " +
        "before_section_id or after_section_id as the reference.",
      input_schema: {
        type: "object",
        properties: {
          parent_id: {
            type: "string",
            description: "Notebook or section group holding the sections. Required.",
          },
          section_id: { type: "string", description: "Section to reposition. Required." },
          before_section_id: { type: "string", description: "Place before this section." },
          after_section_id: { type: "string", description: "Place after this section." },
        },
        required: ["parent_id", "section_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("reorder_sections", {
        parent_id: input.parent_id,
        section_id: input.section_id,
        before_section_id: input.before_section_id ?? "",
        after_section_id: input.after_section_id ?? "",
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function updatePageTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_update_page",
      description:
        "Update a page's body content. mode 'replace' (default) clears the existing body outlines and " +
        "writes the new content; mode 'append' adds to the end. The page title is preserved. Newlines " +
        "become separate paragraphs.",
      input_schema: {
        type: "object",
        properties: {
          page_id: { type: "string", description: "Page object ID. Required." },
          content: { type: "string", description: "New body text. Required." },
          mode: {
            type: "string",
            enum: ["replace", "append"],
            description: "replace (default) or append.",
          },
        },
        required: ["page_id", "content"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("update_page", {
        page_id: input.page_id,
        content: input.content,
        mode: input.mode ?? "replace",
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function insertRichContentTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_insert_rich_content",
      description:
        "Append rich content to a page: a well-formed XHTML fragment and/or an image from a local " +
        "file path. Supported blocks: h1-h6 (rendered as sized bold text), p, ul/ol with nested " +
        "lists, table/tr/th/td, pre (Consolas lines), blockquote, div (recursed). Inline: b/strong, " +
        "i/em, u, code, a href, span with style. Multiple sibling root elements are fine. The " +
        "fragment must parse as XML — self-close void tags and match every open tag. Provide html " +
        "and/or image_path.",
      input_schema: {
        type: "object",
        properties: {
          page_id: { type: "string", description: "Page object ID. Required." },
          html: {
            type: "string",
            description:
              "Well-formed XHTML fragment. Block elements become OneNote paragraphs/lists/tables; " +
              "headings render as sized bold text.",
          },
          image_path: { type: "string", description: "Absolute path to an image file to embed." },
        },
        required: ["page_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("insert_rich_content", {
        page_id: input.page_id,
        html: input.html ?? "",
        image_path: input.image_path ?? "",
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function exportTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_export",
      description:
        "Export a page or section to a file. format: pdf (default), html, docx, mhtml, xps, or onenote. " +
        "target_path is the absolute output file path.",
      input_schema: {
        type: "object",
        properties: {
          object_id: { type: "string", description: "Page or section object ID. Required." },
          target_path: { type: "string", description: "Absolute output file path. Required." },
          format: {
            type: "string",
            enum: ["pdf", "html", "docx", "mhtml", "xps", "onenote"],
            description: "Export format. Default pdf.",
          },
        },
        required: ["object_id", "target_path"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("export", {
        object_id: input.object_id,
        target_path: input.target_path,
        format: input.format ?? "pdf",
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function deleteSectionTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_delete_section",
      description:
        "Delete a section. Moves it to the notebook's recycle bin by default (recoverable); pass " +
        "permanent:true to erase. Destructive — confirm the section ID first. Deletes the whole " +
        "section and all its pages.",
      input_schema: {
        type: "object",
        properties: {
          section_id: { type: "string", description: "Section object ID. Required." },
          permanent: { type: "boolean", description: "Erase permanently instead of recycle bin. Default false." },
        },
        required: ["section_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("delete_section", {
        section_id: input.section_id,
        permanent: input.permanent ?? false,
      });
      return { output: asJson(result), isError: false };
    },
  };
}

function deleteNotebookTool(bridge: OneNoteBridge): McpTool {
  return {
    readOnly: false,
    definition: {
      name: "onenote_delete_notebook",
      description:
        "Delete/close a whole notebook. Highly destructive — confirm the notebook ID explicitly with " +
        "the user first. Recoverable via recycle bin by default; permanent:true erases.",
      input_schema: {
        type: "object",
        properties: {
          notebook_id: { type: "string", description: "Notebook object ID. Required." },
          permanent: { type: "boolean", description: "Erase permanently. Default false." },
        },
        required: ["notebook_id"],
        additionalProperties: false,
      },
    },
    async execute(input) {
      const result = await bridge.call("delete_notebook", {
        notebook_id: input.notebook_id,
        permanent: input.permanent ?? false,
      });
      return { output: asJson(result), isError: false };
    },
  };
}

export function createOneNoteHostTools(bridge: OneNoteBridge): McpTool[] {
  return [
    // Read
    hierarchyTool(bridge),
    getPageTool(bridge),
    searchTool(bridge),
    // Knowledge operator
    ...createKnowledgeTools(bridge),
    // Page content
    createPageTool(bridge),
    appendPageTool(bridge),
    updatePageTool(bridge),
    insertRichContentTool(bridge),
    renamePageTool(bridge),
    movePageTool(bridge),
    reorderPagesTool(bridge),
    // Hierarchy
    createSectionTool(bridge),
    createSectionGroupTool(bridge),
    createNotebookTool(bridge),
    renameSectionTool(bridge),
    moveSectionTool(bridge),
    reorderSectionsTool(bridge),
    // Navigation / export
    navigateTool(bridge),
    exportTool(bridge),
    // Delete
    deletePageTool(bridge),
    deleteSectionTool(bridge),
    deleteNotebookTool(bridge),
  ];
}

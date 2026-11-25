const express = require("express");
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");
const cors = require("cors");
const { execFile, execSync } = require("child_process");
const selfsigned = require("selfsigned");

const LAST_SESSION_PATH = path.join(__dirname, ".last-session.json");

function saveLastSession(history) {
  try {
    fs.writeFileSync(LAST_SESSION_PATH, JSON.stringify((history || []).slice(-16), null, 2), "utf8");
  } catch (e) {}
}

function loadLastSession() {
  try {
    if (fs.existsSync(LAST_SESSION_PATH)) {
      const data = JSON.parse(fs.readFileSync(LAST_SESSION_PATH, "utf8"));
      return Array.isArray(data) ? data : [];
    }
  } catch (e) {}
  return [];
}

const app = express();
app.use(cors());
app.use(express.json({ limit: "20mb" }));
app.use(express.urlencoded({ extended: true, limit: "20mb" }));

// Anti-caching middleware for all webview and static assets
app.use((req, res, next) => {
  res.set("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
  res.set("Pragma", "no-cache");
  res.set("Expires", "0");
  res.set("Surrogate-Control", "no-store");
  next();
});

// Serve static taskpane assets with zero caching
app.use(
  express.static(path.join(__dirname, "public"), {
    maxAge: 0,
    etag: false,
    lastModified: false
  })
);

app.get("/", (req, res) => {
  res.redirect("/taskpane.html");
});

function findAgyBinary() {
  const home = os.homedir();
  const binaryName = process.platform === "win32" ? "agy.cmd" : "agy";
  const candidates = [
    path.join(home, ".local", "bin", binaryName),
    path.join(home, "AppData", "Local", "Programs", "agy", binaryName),
    "/opt/homebrew/bin/agy",
    "/usr/local/bin/agy",
    "/usr/bin/agy"
  ];
  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      try {
        fs.accessSync(candidate, fs.constants.X_OK);
        return candidate;
      } catch (e) {
        // Check remaining candidates
      }
    }
  }
  return "agy";
}

const AGY_PATH = findAgyBinary();

const ENRICHED_PATH = [
  path.join(process.env.HOME || "", ".local", "bin"),
  "/opt/homebrew/bin",
  "/usr/local/bin",
  "/usr/bin",
  "/bin",
  process.env.PATH || ""
].filter(Boolean).join(":");

const DEFAULT_AGY_MODEL = "gemini-3.8-flash-high";
const SUPPORTED_HOSTS = ["Word", "Excel", "PowerPoint"];

const HOST_OP_TYPES = {
  Word: [
    "replace_full_document",
    "replace_section",
    "standardize_document",
    "replace_selection",
    "search_and_replace",
    "insert_at_location",
    "format_text",
    "add_comment",
    "delete_text",
    "set_header",
    "set_footer",
    "insert_page_break",
    "update_chart",
    "format_table",
    "execute_office_js"
  ],
  Excel: [
    "set_range_values",
    "set_formula",
    "format_range",
    "format_table",
    "insert_chart",
    "create_sheet",
    "clear_range",
    "execute_office_js"
  ],
  PowerPoint: [
    "generate_slides",
    "insert_slide",
    "apply_theme",
    "format_slide_table",
    "replace_slide_selection"
  ]
};

const SUPPORTED_OP_TYPES = Array.from(
  new Set([...HOST_OP_TYPES.Word, ...HOST_OP_TYPES.Excel, ...HOST_OP_TYPES.PowerPoint])
);

const BUILTIN_SKILLS_DIR = path.join(__dirname, "skills");
const USER_HOME = process.env.HOME || process.env.USERPROFILE || "";
const USER_DATA_DIR = USER_HOME ? path.join(USER_HOME, ".local", "share", "office-ai") : path.join(__dirname, ".local-data");
const CUSTOM_SKILLS_DIR = path.join(USER_DATA_DIR, "custom_skills");
const SKILLS_DIR = BUILTIN_SKILLS_DIR; // legacy alias

function getSkillDirectories() {
  const dirs = [BUILTIN_SKILLS_DIR];
  if (CUSTOM_SKILLS_DIR && !dirs.includes(CUSTOM_SKILLS_DIR)) {
    dirs.push(CUSTOM_SKILLS_DIR);
  }
  return dirs;
}

function parseListValue(rawVal) {
  return String(rawVal || "")
    .replace(/^\[|\]$/g, "")
    .split(",")
    .map((s) => s.trim().replace(/^["']|["']$/g, ""))
    .filter(Boolean);
}

function parseSkillFile(filePath) {
  const raw = fs.readFileSync(filePath, "utf8");
  const id = path.basename(filePath, ".md");
  const fmMatch = raw.match(/^---\s*\n([\s\S]*?)\n---\s*\n?([\s\S]*)$/);
  const meta = {
    id,
    name: id,
    description: "",
    hosts: [...SUPPORTED_HOSTS],
    alwaysActive: false,
    keywords: [],
    learned: id.startsWith("learned-")
  };
  let body = raw.trim();

  if (fmMatch) {
    body = fmMatch[2].trim();
    for (const line of fmMatch[1].split("\n")) {
      const colonIdx = line.indexOf(":");
      if (colonIdx === -1) continue;
      const key = line.slice(0, colonIdx).trim();
      const val = line.slice(colonIdx + 1).trim();
      if (key === "name") meta.name = val.replace(/^["']|["']$/g, "");
      else if (key === "description") meta.description = val.replace(/^["']|["']$/g, "");
      else if (key === "hosts") meta.hosts = parseListValue(val);
      else if (key === "alwaysActive") meta.alwaysActive = val.toLowerCase() === "true";
      else if (key === "keywords") meta.keywords = parseListValue(val).map((k) => k.toLowerCase());
    }
  }

  return { ...meta, body };
}

function listAllSkills() {
  const seen = new Set();
  const all = [];
  for (const dir of getSkillDirectories()) {
    if (!fs.existsSync(dir)) continue;
    try {
      const files = fs
        .readdirSync(dir)
        .filter((f) => f.endsWith(".md"))
        .sort();
      for (const file of files) {
        const id = path.basename(file, ".md");
        if (seen.has(id)) continue;
        seen.add(id);
        all.push(parseSkillFile(path.join(dir, file)));
      }
    } catch (e) {}
  }
  return all;
}

function loadSkillsForRequest({ host = "Word", prompt = "", selectedText = "", fullDocumentText = "" }) {
  const activeHost = normalizeHost(host);
  const searchCorpus = `${prompt} ${selectedText} ${String(fullDocumentText || "").slice(0, 1500)}`.toLowerCase();

  return listAllSkills().filter((skill) => {
    const hostMatch =
      !skill.hosts ||
      skill.hosts.length === 0 ||
      skill.hosts.some((h) => h.toLowerCase() === activeHost.toLowerCase());
    if (!hostMatch) return false;
    if (skill.alwaysActive || skill.learned) return true;
    return skill.keywords.some((kw) => searchCorpus.includes(kw));
  });
}

function writeSkillToDirectories(filename, markdownContent) {
  // Always write learned or user custom skills ONLY into custom_skills directories (never into built-in git-tracked skills/)
  const targetDirs = [CUSTOM_SKILLS_DIR];
  const sourceMarker = path.join(__dirname, ".source-dir");
  if (fs.existsSync(sourceMarker)) {
    try {
      const sourceRoot = fs.readFileSync(sourceMarker, "utf8").trim();
      if (sourceRoot) targetDirs.push(path.join(sourceRoot, "custom_skills"));
    } catch (e) {}
  }
  for (const dir of targetDirs) {
    try {
      fs.mkdirSync(dir, { recursive: true });
      fs.writeFileSync(path.join(dir, filename), markdownContent, "utf8");
    } catch (err) {
      console.warn(`[Skills] Could not write to ${dir}: ${err.message}`);
    }
  }
}

function deleteSkillFromDirectories(skillId) {
  const safeId = String(skillId || "").replace(/[^a-z0-9_-]/gi, "");
  if (!safeId || safeId !== skillId) return false;
  let removed = false;
  for (const dir of getSkillDirectories()) {
    const filePath = path.join(dir, `${safeId}.md`);
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
        removed = true;
      }
    } catch (err) {
      console.warn(`[Skills] Could not delete ${filePath}: ${err.message}`);
    }
  }
  return removed;
}

function normalizeHost(host) {
  const value = String(host || "Word").toLowerCase();
  if (value.includes("excel")) return "Excel";
  if (value.includes("powerpoint") || value.includes("ppt")) return "PowerPoint";
  return "Word";
}

app.get("/api/status", (req, res) => {
  let agyAvailable = false;
  try {
    agyAvailable = path.isAbsolute(AGY_PATH) ? fs.existsSync(AGY_PATH) : true;
  } catch (e) {
    agyAvailable = false;
  }

  res.json({
    ok: true,
    name: "0x15k Office AI",
    author: "0x15k - YL",
    version: "4.5.0",
    agyPath: AGY_PATH,
    agyAvailable,
    defaultAgyModel: DEFAULT_AGY_MODEL,
    supportedHosts: SUPPORTED_HOSTS,
    hostOperationTypes: HOST_OP_TYPES,
    supportedProviders: ["agy", "gemini", "claude", "openai"],
    skillsCount: listAllSkills().length,
    capabilities: {
      agenticOperations: true,
      urlContextFetching: true,
      skillsEngine: true,
      skillfy: true,
      cli: true,
      supportedOperationTypes: SUPPORTED_OP_TYPES,
      redlineDiff: true,
      ooxmlUndo: true
    },
    timestamp: new Date().toISOString()
  });
});

app.get("/api/skills", (req, res) => {
  const skills = listAllSkills().map(({ id, name, description, hosts, alwaysActive, keywords, learned }) => ({
    id,
    name,
    description,
    hosts,
    alwaysActive,
    keywords,
    learned
  }));
  res.json({ ok: true, skills });
});

app.delete("/api/skills/:id", (req, res) => {
  const { id } = req.params;
  if (!id || !id.startsWith("learned-")) {
    return res.status(400).json({ error: "Only learned-* skills can be deleted." });
  }
  const removed = deleteSkillFromDirectories(id);
  if (!removed) {
    return res.status(404).json({ error: `Skill not found: ${id}` });
  }
  res.json({ ok: true, deleted: id });
});

app.post("/api/session/clear", (req, res) => {
  try {
    if (fs.existsSync(LAST_SESSION_PATH)) {
      fs.unlinkSync(LAST_SESSION_PATH);
    }
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

function getHostOperationsSchema(host) {
  if (host === "Excel") {
    return (
      "Valid Excel operation types:\n" +
      '- format_table: {"type":"format_table","address":"A1:E10","headerBgColor":"#156082","headerTextColor":"#FFFFFF","headerBold":true,"autoFit":true,"description":"..."} -> Directly and safely styles an Excel table with corporate header shading (#156082), white bold text, and column auto-fit.\n' +
      '- insert_chart: {"type":"insert_chart","dataRange":"A1:B10","chartType":"ColumnClustered"|"BarClustered"|"Pie"|"Line","title":"Sales Analysis","description":"..."} -> Creates a native Excel chart with corporate colors.\n' +
      '- set_range_values: {"type":"set_range_values","startCell":"A1","values":[["Header 1","Header 2"],["Row 1",100]],"createTable":true,"description":"..."}\n' +
      '- set_formula: {"type":"set_formula","cell":"C2","formula":"=SUM(A2:B2)","description":"..."}\n' +
      '- format_range: {"type":"format_range","address":"A1:D1","style":{"bold":true,"fillColor":"#156082","fontColor":"#FFFFFF","numberFormat":"$#,##0.00","autoFit":true},"description":"..."}\n' +
      '- create_sheet: {"type":"create_sheet","name":"Summary","values":[["Metric","Value"]],"description":"..."}\n' +
      '- clear_range: {"type":"clear_range","address":"A1:Z50","description":"..."}\n' +
      '- execute_office_js: {"type":"execute_office_js","description":"...","code":"context.workbook...; await context.sync();"}'
    );
  }

  if (host === "PowerPoint") {
    return (
      "Valid PowerPoint operation types:\n" +
      '- apply_theme: {"type":"apply_theme","primaryColor":"#156082","fontName":"Aptos","description":"..."} -> Standardizes entire deck typography to Aptos and applies corporate #156082 styling to titles and key accents.\n' +
      '- format_slide_table: {"type":"format_slide_table","headerBgColor":"#156082","headerTextColor":"#FFFFFF","description":"..."} -> Formats slide tables with corporate #156082 header shading and white bold text.\n' +
      '- insert_slide: {"type":"insert_slide","title":"Executive Overview","bullets":["Key takeaway 1","Key takeaway 2"],"description":"..."}\n' +
      '- generate_slides: {"type":"generate_slides","slides":[{"title":"Slide Title","bullets":["Point 1","Point 2"]}],"description":"..."}\n' +
      '- replace_slide_selection: {"type":"replace_slide_selection","content":"Updated slide text","description":"..."}'
    );
  }

  return (
    "Valid Word operation types:\n" +
    '- replace_full_document: {"type":"replace_full_document","fontName":"Times New Roman","fontSize":10.5,"lineSpacing":1.15,"content":"<full markdown>","description":"..."} (fontName/fontSize/lineSpacing optional; always place replace_full_document BEFORE standardize_document)\n' +
    '- replace_section: {"type":"replace_section","heading":"<exact section heading>","content":"<markdown section including heading>","description":"..."} -> Replaces an entire section from its heading up to the next heading.\n' +
    '- standardize_document: {"type":"standardize_document","fontName":"Calibri","fontSize":10.5,"lineSpacing":1.15,"description":"..."}\n' +
    '- replace_selection: {"type":"replace_selection","content":"<markdown>","description":"..."}\n' +
    '- search_and_replace: {"type":"search_and_replace","find":"<exact text>","replace":"<new text>","matchCase":false,"description":"..."}\n' +
    '- insert_at_location: {"type":"insert_at_location","location":"start"|"end"|"after_target"|"before_target","target":"<anchor>","content":"<markdown>","description":"..."}\n' +
    '- format_text: {"type":"format_text","target":"<exact text or empty>","style":{"bold":true,"italic":false,"underline":false,"color":"#111827","highlight":"#FEF08A","fontSize":11,"fontName":"Calibri","alignment":"Left"|"Centered"|"Right"|"Justified","heading":1},"description":"..."}\n' +
    '- add_comment: {"type":"add_comment","target":"<anchor or empty>","comment":"<text>","description":"..."}\n' +
    '- delete_text: {"type":"delete_text","target":"<exact text>","description":"..."}\n' +
    '- set_header / set_footer: {"type":"set_header"|"set_footer","location":"replace"|"end","content":"<markdown>","description":"..."}\n' +
    '- insert_page_break: {"type":"insert_page_break","location":"end"|"after_target"|"selection","target":"<optional>","description":"..."}\n' +
    '- update_chart: {"type":"update_chart","target":"<heading, section anchor, or empty if chart is selected>","chartType":"pie"|"doughnut"|"bar"|"column"|"line","title":"<optional title>","showDataLabels":false,"fontName":"Aptos","fontSize":9,"categories":["Cat 1","Cat 2",...],"series":[{"name":"Series Name","values":[10,20,...]}],"colors":{"CategoryOrSeries":"#HEX"},"description":"..."} -> Updates or creates a DrawingML OOXML chart in Word with native Chart Design & Format ribbon support. Universal and domain-agnostic: accepts arbitrary categories, numeric series values, custom colors, data label toggle, and font styling.\n' +
    '- format_table: {"type":"format_table","target":"all"|"first"|1|2,"headerBgColor":"#156082","headerTextColor":"#FFFFFF","headerBold":true,"enableBorders":true,"fontName":"Aptos","fontSize":10,"description":"..."} -> Directly and safely styles native Word tables: sets header row background shading (e.g. institutional #156082), white bold text, typography, and clean borders without replacing text or duplicating sections.\n' +
    '- execute_office_js: {"type":"execute_office_js","description":"...","code":"context.document...; await context.sync();"} (Code Mode: direct transactional Word.js execution)'
  );
}

/**
 * Builds a concise, host-aware, language-agnostic system prompt with dynamically loaded skills.
 */
function buildSystemPrompt(customSystemPrompt, agentMode = true, host = "Word", activeSkills = []) {
  const activeHost = normalizeHost(host);
  const basePrompt =
    customSystemPrompt && customSystemPrompt.trim()
      ? customSystemPrompt.trim()
      : `You are a fast, precise, autonomous AI agent running directly inside Microsoft ${activeHost}.\n` +
        "You are backed by the 0x15k local daemon with the Autonomous Skills Engine, local filesystem skills storage (skills/), and the full Office.js / OOXML execution bridge.\n" +
        "LANGUAGE RULE: Never hardcode a single language. Preserve the document's existing language when editing (unless asked to translate), and write your brief chat reply in the user's prompt language.\n" +
        `NEVER tell the user to perform manual ${activeHost} UI steps. Never claim you lack access to tools or skills. Apply edits directly via JSON operations.\n` +
        "MULTI-POINT CONSISTENCY: When the user requests a status change (e.g. 'findings mitigated', 'actualiza estado', 'cambió el resultado'), you MUST update all referenced sections atomically: Summary Tables, Detail Records, Charts, and Conclusions. Never leave a chart or conclusion contradicting the data.";

  const skillsBlock =
    Array.isArray(activeSkills) && activeSkills.length > 0
      ? "\n\n[ACTIVE_SKILLS]\n" +
        activeSkills.map((s) => `## Skill: ${s.name}\n${s.body}`).join("\n\n") +
        "\n[/ACTIVE_SKILLS]"
      : "";

  const wordRules =
    activeHost === "Word"
      ? "\nWORD FORMATTING, STYLE INHERITANCE & HYPERLINK RULES:\n" +
        "- STYLE INHERITANCE & PERSISTENCE: Always inspect and match [DOCUMENT_STYLE_STATE]. If [DOCUMENT_STYLE_STATE] specifies a predominant font (e.g. Aptos, Calibri), size, or multiple line spacing (e.g. 1.15x), all newly generated text, replacements, and standardize_document operations MUST adhere to these exact parameters. Never reset line spacing or font family to arbitrary defaults on subsequent turns.\n" +
        "- HYPERLINK PRESERVATION: Any links in [FULL_CONTEXT] or [DOCUMENT_HYPERLINKS] formatted as `[Display Text](url)` MUST be preserved when rewriting or updating sections. When generating links, always output them as `[Display Text](url)` so they render as native clickable hyperlinks in Word. Never drop the target URL.\n" +
        "- HEADER ALIGNMENT: Candidate names and contact detail lines centered at the top of resumes must remain centered (<div align=\"center\"> or centered paragraphs). Never force centered resume headers to the left.\n" +
        "- RESUME DUAL-ALIGNMENT (LOCATIONS & DATES): For entries having titles/companies on the left and dates/locations on the right, ALWAYS use dual-split formatting (`Entity | Location` or markdown tables `| Entity | Location |`). This guarantees locations and dates anchor flush to the right margin across all sections. NEVER use typewriter tab stops (`\\t`).\n" +
        "- SECTION HEADINGS: Format all major section headings consistently as Heading 2 uppercase (e.g. `## EDUCATION`, `## EXPERIENCE`, `## TECHNICAL SKILLS`, `## HONORS & AWARDS`, `## CERTIFICATIONS`).\n" +
        "- LIST SECTIONS INTEGRITY (HONORS & AWARDS, CERTIFICATIONS): In list sections like HONORS & AWARDS, AWARDS, or CERTIFICATIONS, EVERY single item MUST be formatted as an active bullet item (`- Item`). NEVER leave any award, competition ranking, or certification sitting as unbulleted plain text. If an item is unbulleted, format it as a bullet!\n" +
        "- BULLET ITEMS: In bulleted lists, never prefix list items with redundant bullet markers, dashes, or arrows (write `- Conducted penetration tests...`, NOT `- -> Conducted...` and NOT `- • Conducted...`). Never emit empty bullet items.\n" +
        "- EXHAUSTIVE DOCUMENT REVIEW: When the user asks to review or clean the whole document ('revisa todo', 'revisa formato y estructura'), you MUST audit EVERY section in [FULL_CONTEXT] (Header, Summary, Education, Experience, Certifications, Technical Skills, Honors & Awards). If any section has defects (missing bullets, unformatted awards, double arrows, residual tabs), emit operations for each defective section. NEVER claim in chat that you assigned bullets to certifications or awards unless your operations block actually includes the update for that section!\n" +
        "- LINE SPACING: Standard multiple line spacing is specified as a multiplier (e.g. 1.15 for 1.15x multiple, 1.5 for 1.5x multiple) in standardize_document and replace_full_document.\n" +
        "- CHARTS & DRAWINGML MUTATION VIA OOXML: Charts in Word are DrawingML objects stored in OOXML (word/charts/chart1.xml). You can transform or create any chart via update_chart with arbitrary categories, series, and types (bar, column, pie, doughnut, line). Word renders them with full Chart Design and Format ribbon tabs. NEVER claim Word cannot edit charts, and NEVER claim you need Python or external scripts.\n" +
        "- MANDATORY CHART ACTION PROTOCOL: If the user asks to change, format, re-shape, or transform a chart (e.g. 'cambia el gráfico', 'hazlo circular', 'pie chart', 'cambia la forma', 'actualiza datos', 'ponlo en columnas', 'ponlo de barras', 'sales chart', etc.):\n" +
        "  1. You MUST ALWAYS append the ```json:operations``` block containing an `update_chart` operation.\n" +
        "  2. Set `chartType: \"pie\"` (or \"doughnut\", \"column\", \"bar\", \"line\" as requested).\n" +
        "  3. Set `categories` and `series` with the exact categories and quantities from the document context or user request (e.g. categories: [\"Q1\", \"Q2\", ...], series: [{ name: \"Revenue\", values: [100, 200, ...] }]). For single-series pie charts, pass the slice labels in categories and slice counts in series[0].values.\n" +
        "  4. CONTEXT-DRIVEN COLORS & ACTIVE SKILLS: Inspect [ACTIVE_SKILLS] and document context for corporate color schemes (e.g. status colors like MITIGADO #92D050, VULNERABLE #FF0000; or severity colors like CRÍTICO #501549, ALTO #FF0000, MEDIO #FFC000, BAJO #4EA72E). If colors are defined in an active skill or requested by the user, pass them in `colors: { \"CategoryOrSeriesName\": \"#HEX\" }` within `update_chart`. If no color scheme is provided by any active skill or context, ask the user for their preferred brand colors or use standard professional harmonic tones.\n" +
        "  5. VISUAL EXHIBIT SEMANTICS & METRIC DIMENSION: In executive summaries or status reports, always align the chart metric with the document's primary business question:\n" +
        "     - In Recertification / Audit reports, the executive metric is remediation status (e.g. 100% Mitigado vs 0% Vulnerable). Slicing a pie chart by severity when all findings are mitigated misleads stakeholders into thinking vulnerabilities remain active! If showing severities in a recertification, explicitly reflect mitigation (e.g. 'Media (Mitigada)' #FFC000, 'Baja (Mitigada)' #4EA72E) or chart status breakdown ('MITIGADO' 100% #92D050, 'VULNERABLE' 0% #FF0000).\n" +
        "     - Never invent hardcoded domain logic: derive categories, series, and values strictly from the governing summary table and active skills.\n" +
        "  6. CHART LABELS, TYPOGRAPHY & TITLE HYGIENE:\n" +
        "     - When the user asks to remove or clean labels ('quita labels', 'sin etiquetas', 'menos texto', 'limpio', 'quita tantos labels'), set `showDataLabels: false`. In bar and column charts, ALWAYS default `showDataLabels: false` so bars remain crisp without overlapping numbers (the axis scale already shows quantities).\n" +
        "     - When the user requests a specific font or size (e.g. 'sea aptos de número 9', 'letra aptos 9', 'fuente 9'), set `fontName: \"Aptos\"` and `fontSize: 9`.\n" +
        "     - When the user asks for no title ('que no lleve título', 'sin encabezado', 'sin título'), set `title: \"\"`.\n" +
        "  7. If [SELECTED_CONTEXT] indicates a chart is selected ([SELECTED_CHART: ...]), set `\"target\": \"\"` to update the selected chart directly.\n" +
        "  8. NEVER say in your text reply that you changed or updated the chart without actually appending the update_chart operation in the exact same turn!\n"
      : "";

  const excelRules =
    activeHost === "Excel"
      ? "\nEXCEL FORMULAS, DATA GRID & STYLING RULES:\n" +
        "- INSTITUTIONAL DESIGN TOKENS (#156082): Corporate table headers require background fill `#156082` with white bold text (`#FFFFFF`) and `Aptos` font. Use `format_table` with `headerBgColor: \"#156082\"` and `autoFit: true`.\n" +
        "- FORMULAS SYNTAX: All formulas MUST begin with an equals sign `=` (e.g. `=SUM(B2:B10)`, `=AVERAGE(C2:C12)`, `=(C2-B2)/B2`, `=IF(D2>0, \"Profit\", \"Loss\")`). Never write formulas as raw text without `=`. Always use standard uppercase A1 notation.\n" +
        "- DATA INSERTION: Use `set_range_values` with 2D arrays (`values: [[\"Header 1\", \"Header 2\"], [\"Value 1\", 100]]`). For placeholder cells calculated by `set_formula`, provide an empty string `\"\"` in `set_range_values`.\n" +
        "- STRUCTURED TABLES: When building tables, models, or data breakdowns, set `createTable: true` in `set_range_values` to create an official Excel structured table with filter buttons and banded rows.\n" +
        "- FINANCIAL & METRIC FORMATTING: Pair numeric insertions with `format_range` to apply standard number formats (`numberFormat: \"$#,##0.00\"` for currency, `\"0.0%\"` for percentages, `\"#,##0\"` for integers, `\"YYYY-MM-DD\"` for dates) and highlight headers (`bold: true`, `fillColor: \"#156082\"`, `fontColor: \"#FFFFFF\"`, `autoFit: true`).\n" +
        "- NATIVE CHARTS: Use `insert_chart` with `dataRange` and `chartType` (\"ColumnClustered\", \"BarClustered\", \"Pie\", \"Line\") to plot data cleanly without external images.\n" +
        "- EXISTING DATA RESPECT: When updating or calculating in existing sheets, inspect [FULL_CONTEXT] for cell coordinates and existing values. Only modify or target the relevant ranges.\n"
      : "";

  const powerPointRules =
    activeHost === "PowerPoint"
      ? "\nPOWERPOINT SLIDE DESIGN & EXECUTIVE STORYTELLING RULES:\n" +
        "- INSTITUTIONAL BRANDING (#156082): Use corporate primary color `#156082` for slide titles and accent cards. Apply Aptos font across all text frames via `apply_theme`.\n" +
        "- SLIDE TABLES: Style all slide comparison tables with `format_slide_table` using `#156082` header shading and white bold text.\n" +
        "- ACTION-ORIENTED TITLES: Every slide title in `generate_slides` or `insert_slide` must be a clear, informative action headline conveying the core takeaway (3 to 8 words), not a generic single word.\n" +
        "- CONCISE BULLET POINTS: Each slide in `generate_slides` MUST contain 3 to 5 concise bullet points. Keep each bullet under 15 words. Never emit dense paragraphs.\n" +
        "- BULLETS ARRAY FORMAT: In `generate_slides`, `bullets` MUST be an array of plain strings (`[\"Point 1\", \"Point 2\"]`). NEVER prefix strings in the array with Markdown hashes, hyphens, or asterisks (PowerPoint adds native bullet formatting automatically).\n" +
        "- SLIDE SELECTION REPLACEMENT: Use `replace_slide_selection` with `content` when modifying or refining currently selected slide text or shape content.\n"
      : "";

  const hostSpecificRules = wordRules + excelRules + powerPointRules;

  if (!agentMode) {
    return basePrompt + hostSpecificRules + skillsBlock;
  }

  return (
    basePrompt +
    hostSpecificRules +
    skillsBlock +
    `\n\n${activeHost.toUpperCase()} MANDATORY ACTION PROTOCOL:\n` +
    "1. DIRECT ACTION BIAS: You are an active document editor directly executing inside Microsoft " +
    activeHost +
    ".\n" +
    "2. If the user mentions ANY defect, visual flaw, asymmetry, or asks you to review/check an issue (e.g. 'revisa esto', 'esto está mal', 'sigue a la izquierda', 'por qué no está alineado', 'corrige esto'), DO NOT JUST EXPLAIN IT IN WORDS.\n" +
    "3. Keep your explanation very concise (1-2 sentences), and ALWAYS append the ```json:operations``` block to actively apply the fix in " +
    activeHost +
    " in that exact same turn.\n" +
    "4. NEVER say 'A continuación se aplica...' or 'I will now apply...' without actually appending the ```json:operations``` block!\n" +
    "5. Format for operations block:\n" +
    "```json:operations\n" +
    `[{"type":"${HOST_OP_TYPES[activeHost][0]}","description":"..."}]\n` +
    "```\n" +
    getHostOperationsSchema(activeHost) +
    "\nOutput strictly valid JSON inside the operations block."
  );
}

function stripHtmlToPlainText(html) {
  return String(html || "")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Extracts HTTP/HTTPS URLs from the prompt and fetches readable page text with a 4s timeout.
 */
async function extractUrlContextFromPrompt(prompt) {
  const matches = String(prompt || "").match(/https?:\/\/[^\s<>"')]+/gi);
  if (!matches || matches.length === 0) return [];

  const uniqueUrls = Array.from(new Set(matches.map((u) => u.replace(/[.,;:!?]+$/, ""))))
    .filter((u) => {
      try {
        const parsed = new URL(u);
        const host = parsed.hostname.toLowerCase();
        // Prevent SSRF / internal host probes
        if (
          host === "localhost" ||
          host === "127.0.0.1" ||
          host === "0.0.0.0" ||
          host.startsWith("169.254.") ||
          host.startsWith("10.") ||
          host.startsWith("192.168.") ||
          host.endsWith(".local") ||
          host.endsWith(".internal")
        ) {
          return false;
        }
        return parsed.protocol === "http:" || parsed.protocol === "https:";
      } catch {
        return false;
      }
    })
    .slice(0, 2);
  const results = [];

  for (const url of uniqueUrls) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 4000);
    try {
      const response = await fetch(url, {
        signal: controller.signal,
        headers: {
          "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X) OfficeAgent/4.0",
          Accept: "text/html,application/xhtml+xml,text/plain,application/json"
        }
      });
      if (!response.ok) continue;
      const contentType = (response.headers.get("content-type") || "").toLowerCase();
      const rawBody = await response.text();
      const cleanText = contentType.includes("html")
        ? stripHtmlToPlainText(rawBody)
        : rawBody.replace(/\s+/g, " ").trim();
      if (cleanText) {
        results.push({ url, text: cleanText.slice(0, 6000) });
      }
    } catch (err) {
      console.warn(`[URL Fetch] Skipped ${url}: ${err.message}`);
    } finally {
      clearTimeout(timer);
    }
  }

  return results;
}

function buildContextPayload({
  prompt,
  selectedText,
  fullDocumentText,
  documentScope,
  documentOutline,
  webContexts = [],
  host = "Word",
  hyperlinks = {},
  documentStyle = {}
}) {
  const sections = [`[HOST: ${normalizeHost(host)}]`];

  if (documentStyle && typeof documentStyle === "object" && Object.keys(documentStyle).length > 0) {
    const styleLines = [
      `- Predominant Font: ${documentStyle.predominantFont || "Calibri"} (${documentStyle.predominantFontSize || 10.5}pt)`,
      `- Line Spacing: ${documentStyle.predominantLineSpacingMultiplier || 1.15}x Multiple (${documentStyle.predominantLineSpacingPt || 13.8}pt)`,
      `- Paragraph Space After: ${documentStyle.spaceAfter ?? 2}pt`,
      `- Top Header Alignment: ${documentStyle.topHeaderAlignment || "Centered"}`,
      `- Active Hyperlinks: ${documentStyle.hyperlinksCount || Object.keys(hyperlinks || {}).length}`
    ];
    sections.push(`[DOCUMENT_STYLE_STATE]\n${styleLines.join("\n")}\n[/DOCUMENT_STYLE_STATE]`);
  }

  if (hyperlinks && typeof hyperlinks === "object" && Object.keys(hyperlinks).length > 0) {
    const linkLines = Object.entries(hyperlinks)
      .map(([text, url]) => `- "${text}" -> ${url}`)
      .join("\n");
    sections.push(`[DOCUMENT_HYPERLINKS]\n${linkLines}\n[/DOCUMENT_HYPERLINKS]`);
  }

  if (Array.isArray(webContexts) && webContexts.length > 0) {
    for (const web of webContexts) {
      sections.push(`[WEB_PAGE_CONTEXT url="${web.url}"]\n${web.text}\n[/WEB_PAGE_CONTEXT]`);
    }
  }

  if (Array.isArray(documentOutline) && documentOutline.length > 0) {
    const outlineLines = documentOutline
      .slice(0, 500)
      .map((item) => `[#${item.index}${item.style ? "|" + item.style : ""}] ${item.text}`)
      .join("\n");
    sections.push(`[DOCUMENT_OUTLINE]\n${outlineLines}\n[/DOCUMENT_OUTLINE]`);
  }

  if (selectedText) {
    sections.push(`[SELECTED_CONTEXT]\n${selectedText}\n[/SELECTED_CONTEXT]`);
  }

  if (documentScope === "document" && fullDocumentText) {
    sections.push(`[FULL_CONTEXT]\n${fullDocumentText.slice(0, 250000)}\n[/FULL_CONTEXT]`);
  } else if (fullDocumentText && fullDocumentText !== selectedText) {
    sections.push(`[ADDITIONAL_CONTEXT]\n${fullDocumentText.slice(0, 150000)}\n[/ADDITIONAL_CONTEXT]`);
  }

  sections.push(`[USER_REQUEST]\n${prompt}\n[/USER_REQUEST]`);
  return sections.join("\n\n");
}

function buildCompactHistory(conversationHistory = []) {
  if (!Array.isArray(conversationHistory) || conversationHistory.length === 0) {
    return [];
  }
  return conversationHistory.slice(-4).map((msg) => ({
    role: msg.role === "assistant" ? "assistant" : "user",
    content: String(msg.content || "").slice(0, 400)
  }));
}

function buildFullPrompt({
  prompt,
  selectedText,
  fullDocumentText,
  documentScope,
  documentOutline,
  webContexts,
  conversationHistory = [],
  systemPrompt,
  agentMode = true,
  host = "Word",
  activeSkills = [],
  hyperlinks = {},
  documentStyle = {}
}) {
  const parts = [];
  parts.push(`[SYSTEM]\n${buildSystemPrompt(systemPrompt, agentMode, host, activeSkills)}\n[/SYSTEM]`);

  const recentHistory = buildCompactHistory(conversationHistory);
  if (recentHistory.length > 0) {
    const historyLines = recentHistory
      .map((msg) => `${msg.role === "user" ? "User" : "Assistant"}: ${msg.content}`)
      .join("\n");
    parts.push(`[RECENT_HISTORY]\n${historyLines}\n[/RECENT_HISTORY]`);
  }

  parts.push(
    buildContextPayload({
      prompt,
      selectedText,
      fullDocumentText,
      documentScope,
      documentOutline,
      webContexts,
      host,
      hyperlinks,
      documentStyle
    })
  );

  return parts.join("\n\n");
}

function parseAgentResponse(rawReply) {
  const raw = (rawReply || "").trim();
  if (!raw) {
    return { reply: "", operations: [], rawReply: "" };
  }

  const allOpTypes = SUPPORTED_OP_TYPES.join("|");
  const blockPatterns = [
    /```(?:json:operations|operations)\s*([\s\S]*?)```/i,
    /```json\s*(\[\s*\{[\s\S]*?"type"\s*:[\s\S]*?\}\s*\])\s*```/i,
    /```\s*(\[\s*\{[\s\S]*?"type"\s*:[\s\S]*?\}\s*\])\s*```/i,
    new RegExp(`(\\[\\s*\\{\\s*"type"\\s*:\\s*"(?:${allOpTypes})"[\\s\\S]*?\\}\\s*\\])`, "i")
  ];

  let operations = [];
  let cleanReply = raw;

  for (const regex of blockPatterns) {
    const match = cleanReply.match(regex);
    if (match && match[1]) {
      try {
        const parsed = JSON.parse(match[1].trim());
        const list = Array.isArray(parsed)
          ? parsed
          : Array.isArray(parsed?.operations)
            ? parsed.operations
            : [];
        const validOps = list.filter(
          (op) => op && typeof op === "object" && typeof op.type === "string" && SUPPORTED_OP_TYPES.includes(op.type)
        );
        if (validOps.length > 0) {
          operations = validOps;
          cleanReply = cleanReply.replace(match[0], "").trim();
          break;
        }
      } catch (err) {
        console.warn("[parseAgentResponse] Failed to parse operations JSON:", err.message);
      }
    }
  }

  if (!cleanReply && operations.length > 0) {
    cleanReply = operations
      .map((op, idx) => `${idx + 1}. **${op.type}**: ${op.description || "Proposed change"}`)
      .join("\n");
  }

  return {
    reply: cleanReply,
    operations,
    rawReply: raw
  };
}

function execAgyOnce(fullPrompt, agyModel, allowEffort = true) {
  return new Promise((resolve, reject) => {
    const agyArgs = ["--model", agyModel];
    const hasEffortSuffix = /-(low|medium|high|xhigh|max)$/i.test(agyModel);
    if (allowEffort && !hasEffortSuffix) {
      agyArgs.push("--effort", "low");
    }
    agyArgs.push("--disable-slash-commands", "-p", fullPrompt);

    execFile(
      AGY_PATH,
      agyArgs,
      {
        cwd: "/tmp",
        maxBuffer: 1024 * 1024 * 10,
        timeout: 90000,
        env: {
          ...process.env,
          PATH: ENRICHED_PATH
        }
      },
      (error, stdout, stderr) => {
        if (error) {
          const detail = (stderr || error.message || "").trim();
          if (allowEffort && /conflicts with --effort|effort.*not supported|invalid.*effort/i.test(detail)) {
            console.warn(`[execAgyOnce] Effort flag conflict detected for model ${agyModel}. Retrying without --effort...`);
            return execAgyOnce(fullPrompt, agyModel, false).then(resolve).catch(reject);
          }
          return reject(new Error(`Antigravity CLI error (${AGY_PATH}): ${detail}`));
        }
        resolve(stdout.trim());
      }
    );
  });
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function callAgyCli(fullPrompt, model) {
  const agyModel = model && model.trim() ? model.trim() : DEFAULT_AGY_MODEL;
  const maxAttempts = 2;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await execAgyOnce(fullPrompt, agyModel);
    } catch (err) {
      const msg = err.message || "";
      const isTransient = /empty conversation id|failed to start conversation|ECONNRESET|ETIMEDOUT/i.test(msg);
      if (attempt < maxAttempts && isTransient) {
        console.warn(`[callAgyCli] Transient error on attempt ${attempt}, retrying in 600ms...`);
        await sleep(600);
        continue;
      }
      throw err;
    }
  }
}

async function callGeminiApi({
  contextText,
  apiKey,
  model = "gemini-2.0-flash",
  systemPrompt,
  conversationHistory = []
}) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  const contents = buildCompactHistory(conversationHistory).map((m) => ({
    role: m.role === "assistant" ? "model" : "user",
    parts: [{ text: m.content }]
  }));
  contents.push({ role: "user", parts: [{ text: contextText }] });

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      contents,
      systemInstruction: { parts: [{ text: systemPrompt }] },
      generationConfig: { temperature: 0.4 }
    })
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Gemini API HTTP ${response.status}: ${errText}`);
  }

  const data = await response.json();
  const reply = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!reply) throw new Error("Empty response from Gemini API.");
  return reply.trim();
}

async function callClaudeApi({
  contextText,
  apiKey,
  model = "claude-3-5-sonnet-latest",
  systemPrompt,
  conversationHistory = []
}) {
  const url = "https://api.anthropic.com/v1/messages";
  const messages = [...buildCompactHistory(conversationHistory), { role: "user", content: contextText }];

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    body: JSON.stringify({
      model,
      system: systemPrompt,
      messages,
      max_tokens: 4096,
      temperature: 0.4
    })
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Claude API HTTP ${response.status}: ${errText}`);
  }

  const data = await response.json();
  const reply = data?.content?.[0]?.text;
  if (!reply) throw new Error("Empty response from Claude API.");
  return reply.trim();
}

async function callOpenAiApi({
  contextText,
  apiKey,
  model = "gpt-4o",
  systemPrompt,
  conversationHistory = []
}) {
  const url = "https://api.openai.com/v1/chat/completions";
  const messages = [
    { role: "system", content: systemPrompt },
    ...buildCompactHistory(conversationHistory),
    { role: "user", content: contextText }
  ];

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model,
      messages,
      temperature: 0.4
    })
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`OpenAI API HTTP ${response.status}: ${errText}`);
  }

  const data = await response.json();
  const reply = data?.choices?.[0]?.message?.content;
  if (!reply) throw new Error("Empty response from OpenAI API.");
  return reply.trim();
}

async function runProviderCompletion({
  provider = "agy",
  apiKey = "",
  model = "",
  compiledSystem,
  contextText,
  conversationHistory = [],
  fullPromptForAgy = ""
}) {
  const activeProvider = (provider || "agy").toLowerCase();
  if (activeProvider === "gemini" && apiKey) {
    return callGeminiApi({
      contextText,
      apiKey,
      model: model || "gemini-2.0-flash",
      systemPrompt: compiledSystem,
      conversationHistory
    });
  }
  if (activeProvider === "claude" && apiKey) {
    return callClaudeApi({
      contextText,
      apiKey,
      model: model || "claude-3-5-sonnet-latest",
      systemPrompt: compiledSystem,
      conversationHistory
    });
  }
  if (activeProvider === "openai" && apiKey) {
    return callOpenAiApi({
      contextText,
      apiKey,
      model: model || "gpt-4o",
      systemPrompt: compiledSystem,
      conversationHistory
    });
  }
  const promptText = fullPromptForAgy || `[SYSTEM]\n${compiledSystem}\n[/SYSTEM]\n\n${contextText}`;
  return callAgyCli(promptText, model || DEFAULT_AGY_MODEL);
}

function anonymizeSkillText(text) {
  if (!text || typeof text !== "string") return "";
  return text
    // Replace emails
    .replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[email protected]")
    // Replace international phone numbers
    .replace(/(?:\+\d{1,3}[\s-]?)?\(?\d{2,4}\)?[\s.-]?\d{3,4}[\s.-]?\d{3,4}/g, "[Phone Number]")
    // Replace profile URLs
    .replace(/https?:\/\/(?:www\.)?(?:linkedin|github)\.com\/[^\s)]+/gi, "[Profile Link]")
    // Replace any remaining web URLs
    .replace(/https?:\/\/[^\s)\],]+/g, "[Web Link]");
}

function parseSkillfyJson(rawText, activeHost) {
  const cleaned = String(rawText || "").trim();
  const fenced = cleaned.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = fenced ? fenced[1].trim() : cleaned;
  const objMatch = candidate.match(/\{[\s\S]*\}/);
  const parsed = JSON.parse(objMatch ? objMatch[0] : candidate);

  const rawSlug = String(parsed.id || parsed.name || "custom-rule")
    .toLowerCase()
    .replace(/^learned-/, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40);
  const id = `learned-${rawSlug || Date.now()}`;
  const name = anonymizeSkillText(String(parsed.name || rawSlug || "Learned Preference").trim());
  const description = anonymizeSkillText(
    String(parsed.description || "User-taught Office formatting and editing preference.").trim()
  );
  const hosts =
    Array.isArray(parsed.hosts) && parsed.hosts.length > 0
      ? parsed.hosts.map((h) => normalizeHost(h))
      : [activeHost];
  const alwaysActive = parsed.alwaysActive !== false;
  const rules = Array.isArray(parsed.rules)
    ? parsed.rules.map((r) => anonymizeSkillText(String(r || "").trim())).filter(Boolean)
    : [];

  if (rules.length === 0) {
    throw new Error("Could not extract actionable rules from the conversation.");
  }

  const blueprint = anonymizeSkillText(
    String(parsed.blueprint || parsed.mockup || parsed.wireframe || parsed.layout || "").trim()
  );
  const expectedSteps = Array.isArray(parsed.expectedSteps || parsed.steps)
    ? (parsed.expectedSteps || parsed.steps).map((s) => anonymizeSkillText(String(s || "").trim())).filter(Boolean)
    : [];

  const antiPattern = anonymizeSkillText(
    String(parsed.antiPattern || parsed.before || parsed.incorrect || "").trim()
  );
  const goodPattern = anonymizeSkillText(
    String(parsed.goodPattern || parsed.after || parsed.correct || "").trim()
  );

  return {
    id,
    name,
    description,
    hosts: Array.from(new Set(hosts)),
    alwaysActive,
    blueprint,
    expectedSteps,
    rules,
    antiPattern,
    goodPattern
  };
}

async function handleSkillfyRequest({
  prompt = "",
  conversationHistory = [],
  host = "Word",
  provider = "agy",
  apiKey = "",
  model = "",
  selectedText = "",
  fullDocumentText = ""
}) {
  const activeHost = normalizeHost(host);
  const userNote = prompt.replace(/^\/skillfy\b/i, "").trim();
  let historyList = Array.isArray(conversationHistory) && conversationHistory.length > 0
    ? conversationHistory.slice(-8)
    : loadLastSession().slice(-8);

  if (historyList.length === 0 && !userNote && !selectedText) {
    return {
      reply:
        "No conversation history or rule note provided yet. Have a few turns in chat (or pass a rule directly, e.g. `/skillfy always keep Technical Skills as separate bold Category: list paragraphs`) so I can distill a reusable skill.",
      operations: [],
      skill: null
    };
  }

  const transcript = historyList
    .map((m) => `${m.role === "user" ? "User" : "Assistant"}: ${String(m.content || "").slice(0, 1200)}`)
    .join("\n\n");

  const skillfySystem =
    "You are an expert AI compiler for Microsoft Office skills (Claude Code / Agent Skill standard).\n" +
    "Your mission: Analyze the user's corrections, complaints, and formatting preferences to distill a permanent, high-precision, completely generic skill.\n\n" +
    "CRITICAL GENERALIZATION & PRIVACY DIRECTIVES:\n" +
    "1. ZERO HARDCODED PERSONAL DATA OR SPECIFIC ENTITIES: NEVER include real candidate names, specific company or university names, real phone numbers, real emails, or personal profile URLs.\n" +
    "2. GEOGRAPHIC & DOMAIN AGNOSTICISM: NEVER hardcode specific countries or cities. Formulate all rules generically (e.g., 'maintain uniform geographic granularity across all entries, such as [City, Region/Country]'). The skill MUST apply universally to ANY user, ANY organization, and ANY country.\n" +
    "3. STRUCTURAL BLUEPRINT / MOCKUP: Always construct a clean ASCII wireframe/mockup illustrating the expected document or section layout flow (e.g. [Header] -> [Summary Table] -> [Visual Exhibit with Caption] -> [Narrative Analysis] -> [Technical Details]).\n" +
    "4. ORDERED EXECUTION STEPS: Detail 3 to 6 sequential imperative operational steps required to achieve compliance (e.g., Step 1: Detect document coordinates, Step 2: Harmonize metrics, Step 3: Format elements, Step 4: Audit alignment).\n" +
    "5. ANONYMOUS GENERIC EXAMPLES: In antiPattern and goodPattern, use clean, anonymous placeholders showing the defect versus the compliant layout.\n\n" +
    "Output ONLY valid JSON matching this schema:\n" +
    "{\n" +
    '  "id": "learned-kebab-case-slug",\n' +
    '  "name": "Generic Descriptive Skill Name",\n' +
    '  "description": "Universal one-sentence summary of when and how this skill applies to any document",\n' +
    `  "hosts": ["${activeHost}"],\n` +
    '  "alwaysActive": true,\n' +
    '  "blueprint": "Structured ASCII wireframe / layout mockup showing the expected visual hierarchy",\n' +
    '  "expectedSteps": [\n' +
    '    "Step 1: Imperative operational step",\n' +
    '    "Step 2: Imperative operational step",\n' +
    '    "Step 3: Imperative operational step"\n' +
    '  ],\n' +
    '  "rules": [\n' +
    '    "Universal imperative rule (what MUST be done, completely generic)",\n' +
    '    "Universal imperative rule (what MUST NEVER be done, completely generic)"\n' +
    "  ],\n" +
    '  "antiPattern": "Generic illustrative bad example using anonymous placeholders showing the structural defect",\n' +
    '  "goodPattern": "Generic illustrative corrected example using anonymous placeholders showing the clean structure"\n' +
    "}\n" +
    "Do NOT invoke any tools. Output strictly valid JSON.";

  const contextParts = [`[TARGET_HOST]\n${activeHost}\n[/TARGET_HOST]`];
  if (transcript) contextParts.push(`[CONVERSATION_TRANSCRIPT]\n${transcript}\n[/CONVERSATION_TRANSCRIPT]`);
  if (userNote) contextParts.push(`[USER_NOTE]\n${userNote}\n[/USER_NOTE]`);
  if (selectedText) contextParts.push(`[SELECTED_DOCUMENT_CONTEXT]\n${selectedText.slice(0, 1500)}\n[/SELECTED_DOCUMENT_CONTEXT]`);
  else if (fullDocumentText) contextParts.push(`[DOCUMENT_SAMPLE]\n${fullDocumentText.slice(0, 1500)}\n[/DOCUMENT_SAMPLE]`);

  const rawJson = await runProviderCompletion({
    provider,
    apiKey,
    model,
    compiledSystem: skillfySystem,
    contextText: contextParts.join("\n\n"),
    conversationHistory: []
  });

  const skill = parseSkillfyJson(rawJson, activeHost);

  let markdownDoc =
    `---\n` +
    `name: ${skill.name}\n` +
    `description: ${skill.description}\n` +
    `hosts: [${skill.hosts.join(", ")}]\n` +
    `alwaysActive: ${skill.alwaysActive}\n` +
    `---\n\n`;

  if (skill.blueprint) {
    markdownDoc += `## Expected Layout Blueprint / Mockup\n\n\`\`\`\n${skill.blueprint}\n\`\`\`\n\n`;
  }

  if (skill.expectedSteps && skill.expectedSteps.length > 0) {
    markdownDoc += `## Expected Execution Steps\n\n` + skill.expectedSteps.map((s, idx) => `${idx + 1}. ${s}`).join("\n") + `\n\n`;
  }

  markdownDoc +=
    `## Enforced Rules\n\n` +
    skill.rules.map((r) => `- ${r}`).join("\n") +
    `\n`;

  if (skill.antiPattern && skill.goodPattern) {
    markdownDoc +=
      `\n## Examples\n\n` +
      `### ❌ Anti-pattern (Before / Rejected)\n\`\`\`\n${skill.antiPattern}\n\`\`\`\n\n` +
      `### ✅ Correct Pattern (After / Compliant)\n\`\`\`\n${skill.goodPattern}\n\`\`\`\n`;
  }

  // Consolidate: prune older redundant learned skills on the same theme
  for (const dir of getSkillDirectories()) {
    try {
      const existingFiles = fs.readdirSync(dir).filter((f) => f.startsWith("learned-") && f.endsWith(".md") && f !== `${skill.id}.md`);
      for (const ef of existingFiles) {
        const oldSkill = parseSkillFile(path.join(dir, ef));
        const hostOverlap = oldSkill.hosts && oldSkill.hosts.some((h) => skill.hosts.includes(h));
        const themeOverlap =
          (skill.name.toLowerCase().includes("resume") && oldSkill.name.toLowerCase().includes("resume")) ||
          (skill.name.toLowerCase().includes("align") && oldSkill.name.toLowerCase().includes("align")) ||
          (skill.name.toLowerCase().includes("format") && oldSkill.name.toLowerCase().includes("format"));
        if (hostOverlap && themeOverlap) {
          try {
            fs.unlinkSync(path.join(dir, ef));
          } catch {}
        }
      }
    } catch {}
  }

  writeSkillToDirectories(`${skill.id}.md`, markdownDoc);

  let reply =
    `🧠 **Skill Learned & Saved:** \`${skill.name}\` (\`${skill.id}.md\`)\n\n` +
    `*${skill.description}*\n\n`;

  if (skill.blueprint) {
    reply += `### 📐 Expected Layout Blueprint\n\`\`\`\n${skill.blueprint}\n\`\`\`\n\n`;
  }

  if (skill.expectedSteps && skill.expectedSteps.length > 0) {
    reply += `### 🪜 Expected Steps\n` + skill.expectedSteps.map((s, idx) => `${idx + 1}. ${s}`).join("\n") + `\n\n`;
  }

  reply += `### 📋 Enforced Rules\n` + skill.rules.map((r) => `- ${r}`).join("\n");

  if (skill.antiPattern && skill.goodPattern) {
    reply +=
      `\n\n### 🔍 Before vs. After Example\n` +
      `**❌ Before (Anti-pattern):**\n\`\`\`\n${skill.antiPattern}\n\`\`\`\n\n` +
      `**✅ After (Compliant):**\n\`\`\`\n${skill.goodPattern}\n\`\`\``;
  }

  return {
    reply,
    operations: [],
    skill: {
      id: skill.id,
      name: skill.name,
      description: skill.description,
      hosts: skill.hosts,
      alwaysActive: skill.alwaysActive,
      learned: true,
      rules: skill.rules,
      antiPattern: skill.antiPattern,
      goodPattern: skill.goodPattern
    }
  };
}

app.post("/api/skillfy", async (req, res) => {
  const startTime = Date.now();
  const {
    prompt = "/skillfy",
    conversationHistory = [],
    host = "Word",
    provider = "agy",
    apiKey = "",
    model = "",
    selectedText = "",
    fullDocumentText = ""
  } = req.body;

  const activeHost = normalizeHost(host);
  const activeProvider = (provider || "agy").toLowerCase();
  const effectiveModel =
    activeProvider === "agy" ? (model && model.trim() ? model.trim() : DEFAULT_AGY_MODEL) : model;

  try {
    const result = await handleSkillfyRequest({
      prompt,
      conversationHistory,
      host: activeHost,
      provider: activeProvider,
      apiKey,
      model: effectiveModel,
      selectedText,
      fullDocumentText
    });
    res.json({
      ...result,
      host: activeHost,
      provider: activeProvider,
      model: effectiveModel,
      durationMs: Date.now() - startTime,
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    res.status(500).json({ error: err.message, operations: [] });
  }
});

app.post("/api/chat", async (req, res) => {
  const startTime = Date.now();
  const {
    prompt = "",
    selectedText = "",
    fullDocumentText = "",
    documentScope = "selection",
    documentOutline = [],
    conversationHistory = [],
    systemPrompt = "",
    provider = "agy",
    apiKey = "",
    model = "",
    agentMode = true,
    host = "Word",
    hyperlinks = {},
    documentStyle = {}
  } = req.body;

  if (!prompt.trim()) {
    return res.status(400).json({ error: "Prompt cannot be empty." });
  }

  const activeHost = normalizeHost(host);
  const activeProvider = (provider || "agy").toLowerCase();
  const effectiveModel =
    activeProvider === "agy" ? (model && model.trim() ? model.trim() : DEFAULT_AGY_MODEL) : model;

  console.log(
    `[POST /api/chat] Host=${activeHost} | Provider=${activeProvider} | Model=${effectiveModel || "default"} | Prompt="${prompt.slice(0, 60)}"`
  );

  try {
    if (prompt.trim().toLowerCase().startsWith("/skillfy")) {
      const learnedResult = await handleSkillfyRequest({
        prompt,
        conversationHistory,
        host: activeHost,
        provider: activeProvider,
        apiKey,
        model: effectiveModel,
        selectedText,
        fullDocumentText
      });
      const durationMs = Date.now() - startTime;
      console.log(`[POST /api/chat] /skillfy completed in ${durationMs}ms`);
      return res.json({
        reply: learnedResult.reply,
        operations: [],
        skill: learnedResult.skill,
        host: activeHost,
        webUrlsUsed: [],
        provider: activeProvider,
        model: effectiveModel,
        durationMs,
        timestamp: new Date().toISOString()
      });
    }

    const webContexts = await extractUrlContextFromPrompt(prompt);
    const activeSkills = loadSkillsForRequest({
      host: activeHost,
      prompt,
      selectedText,
      fullDocumentText
    });
    const compiledSystem = buildSystemPrompt(systemPrompt, agentMode, activeHost, activeSkills);
    const contextText = buildContextPayload({
      prompt,
      selectedText,
      fullDocumentText,
      documentScope,
      documentOutline,
      webContexts,
      host: activeHost,
      hyperlinks,
      documentStyle
    });

    const fullPromptForAgy = buildFullPrompt({
      prompt,
      selectedText,
      fullDocumentText,
      documentScope,
      documentOutline,
      webContexts,
      conversationHistory,
      systemPrompt,
      agentMode,
      host: activeHost,
      activeSkills,
      hyperlinks,
      documentStyle
    });

    const rawReply = await runProviderCompletion({
      provider: activeProvider,
      apiKey,
      model: effectiveModel,
      compiledSystem,
      contextText,
      conversationHistory,
      fullPromptForAgy
    });

    const parsed = parseAgentResponse(rawReply);
    const durationMs = Date.now() - startTime;
    console.log(
      `[POST /api/chat] Completed in ${durationMs}ms | Host=${activeHost} | Skills=${activeSkills.length} | Operations=${parsed.operations.length}`
    );

    try {
      const updatedHistory = Array.isArray(conversationHistory) ? [...conversationHistory] : [];
      updatedHistory.push({ role: "user", content: prompt });
      updatedHistory.push({ role: "assistant", content: parsed.reply });
      saveLastSession(updatedHistory);
    } catch (e) {}

    res.json({
      reply: parsed.reply,
      operations: parsed.operations,
      rawReply: parsed.rawReply,
      activeSkills: activeSkills.map((s) => s.id),
      host: activeHost,
      webUrlsUsed: webContexts.map((w) => w.url),
      provider: activeProvider,
      model: effectiveModel,
      durationMs,
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    const durationMs = Date.now() - startTime;
    console.error(`[POST /api/chat] Failed after ${durationMs}ms:`, err.message);
    res.status(500).json({
      error: err.message,
      reply: `Error processing request: ${err.message}`,
      operations: []
    });
  }
});

function getHttpsCredentials() {
  const keyPath = path.join(__dirname, "key.pem");
  const certPath = path.join(__dirname, "cert.pem");

  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    return {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath),
      source: "mkcert (trusted local certificate)"
    };
  }

  const certAttrs = [{ name: "commonName", value: "localhost" }];
  const pems = selfsigned.generate(certAttrs, { days: 365, keySize: 2048 });
  return {
    key: pems.private,
    cert: pems.cert,
    source: "selfsigned-2048"
  };
}

/**
 * Terminates a conflicting listener on the target port if EADDRINUSE occurs.
 */
function freePortSync(port) {
  try {
    const out = execSync(`lsof -ti :${port}`, { encoding: "utf8" }).trim();
    if (!out) return false;
    const pids = out
      .split(/\s+/)
      .map((p) => Number(p))
      .filter((pid) => Number.isInteger(pid) && pid > 0 && pid !== process.pid);
    for (const pid of pids) {
      try {
        process.kill(pid, "SIGKILL");
        console.log(`[Port Recovery] Terminated stale process PID ${pid} on port ${port}`);
      } catch (e) {
        // Ignore if already exited
      }
    }
    return pids.length > 0;
  } catch (e) {
    return false;
  }
}

function listenWithPortRecovery(server, port, onListening) {
  let retried = false;

  // Proactively clear any conflicting listener on this port before binding
  freePortSync(port);

  server.on("error", (err) => {
    if (err.code === "EADDRINUSE" && !retried) {
      retried = true;
      console.warn(`[Port Recovery] Port ${port} is busy, clearing stale listener...`);
      freePortSync(port);
      setTimeout(() => {
        server.listen(port, onListening);
      }, 400);
      return;
    }
    console.error(`[Server Error] Port ${port}:`, err.message);
    process.exit(1);
  });

  server.listen(port, onListening);
}

const creds = getHttpsCredentials();

const server3000 = https.createServer({ key: creds.key, cert: creds.cert }, app);
listenWithPortRecovery(server3000, 3000, () => {
  console.log("==================================================");
  console.log("0x15k Office AI v4.5.0 - Ready");
  console.log("Taskpane URL: https://localhost:3000/taskpane.html");
  console.log("API Endpoint: https://localhost:3000/api/chat");
  console.log(`Certificate:  ${creds.source}`);
  console.log(`Local CLI:    ${AGY_PATH} (default: ${DEFAULT_AGY_MODEL})`);
  console.log("Hosts:        Word, Excel, PowerPoint");
  console.log("==================================================");
});

const server3001 = https.createServer({ key: creds.key, cert: creds.cert }, app);
listenWithPortRecovery(server3001, 3001, () => {
  console.log("Legacy mirror listener active on https://localhost:3001");
});

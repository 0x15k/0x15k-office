#!/usr/bin/env node
/**
 * 0x15k Office AI - Native Model Context Protocol (MCP) Server
 * Standard JSON-RPC 2.0 over stdio for Anthropic Claude Desktop, Cursor,
 * Antigravity, and any MCP-compliant AI assistant or agent.
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const https = require("https");
const readline = require("readline");
const { execSync } = require("child_process");

const chartGen = require("../lib/chart-generator.js");

const SERVER_NAME = "0x15k-office-mcp";
const SERVER_VERSION = "4.5.0";
const PROTOCOL_VERSION = "2024-11-05";

function logDebug(...args) {
  process.stderr.write(`[office-mcp] ${args.join(" ")}\n`);
}

function requestDaemon(method, urlPath, payload = null) {
  return new Promise((resolve, reject) => {
    const body = payload ? JSON.stringify(payload) : null;
    const headers = body
      ? {
          "Content-Type": "application/json",
          "Content-Length": Buffer.byteLength(body)
        }
      : {};

    const req = https.request(
      {
        hostname: "localhost",
        port: 3000,
        path: urlPath,
        method,
        rejectUnauthorized: false,
        timeout: 60000,
        headers
      },
      (res) => {
        let data = "";
        res.on("data", (chunk) => {
          data += chunk;
        });
        res.on("end", () => {
          try {
            const parsed = JSON.parse(data);
            if (res.statusCode >= 400 || parsed.error) {
              return reject(new Error(parsed.error || `HTTP ${res.statusCode}`));
            }
            resolve(parsed);
          } catch (err) {
            reject(new Error(`Invalid server response (${res.statusCode}): ${data}`));
          }
        });
      }
    );

    req.on("error", (err) => {
      reject(new Error(`Cannot connect to Office AI daemon on port 3000: ${err.message}. Ensure "office-ai start" is running.`));
    });

    if (body) req.write(body);
    req.end();
  });
}

function runWordAppleScript(scriptBody) {
  if (process.platform !== "darwin") {
    throw new Error("Word AppleScript execution is only available on macOS.");
  }
  const fullScript = `
    tell application "Microsoft Word"
      try
        ${scriptBody}
      on error errMsg number errNum
        return "ERROR: " & errMsg & " (" & errNum & ")"
      end try
    end tell
  `;
  try {
    const out = execSync("osascript", {
      input: fullScript,
      encoding: "utf8",
      timeout: 15000
    }).trim();
    if (out.startsWith("ERROR:")) {
      throw new Error(out);
    }
    return out;
  } catch (err) {
    throw new Error(`AppleScript error: ${err.message}`);
  }
}

// Tool definitions for MCP
const TOOLS = [
  {
    name: "office_status",
    description: "Check the health, version, active model, and capabilities of the 0x15k Office AI daemon and Microsoft Office host environment.",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    }
  },
  {
    name: "office_word_get_active_doc",
    description: "Inspect the active Microsoft Word document on macOS: retrieves document name, total paragraph count, word count, active selection text, and sample content.",
    inputSchema: {
      type: "object",
      properties: {
        includeSampleText: {
          type: "boolean",
          description: "Whether to return the text of the first N paragraphs (default true)."
        },
        sampleParagraphCount: {
          type: "number",
          description: "Number of leading paragraphs to sample (default 15, max 100)."
        }
      },
      required: []
    }
  },
  {
    name: "office_word_execute_applescript",
    description: "Execute arbitrary AppleScript code targeting Microsoft Word on macOS to query or manipulate the document directly.",
    inputSchema: {
      type: "object",
      properties: {
        script: {
          type: "string",
          description: "The AppleScript commands to execute inside 'tell application \"Microsoft Word\"'."
        }
      },
      required: ["script"]
    }
  },
  {
    name: "office_word_insert_chart",
    description: "Generate and insert a native DrawingML Flat OPC chart (bar, stacked bar, column, pie) directly into the active Microsoft Word document on macOS.",
    inputSchema: {
      type: "object",
      properties: {
        chartType: {
          type: "string",
          enum: ["bar", "stacked_bar", "column", "pie", "doughnut"],
          description: "Type of chart to render. Default: stacked_bar."
        },
        title: {
          type: "string",
          description: "Chart title. Leave empty for clean charts without title headers."
        },
        categories: {
          type: "array",
          items: { type: "string" },
          description: "Labels for each category/row (e.g. ['CRÍTICA', 'ALTA', 'MEDIA', 'BAJA', 'INFORMATIVA'])."
        },
        series: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              values: { type: "array", items: { type: "number" } },
              color: { type: "string", description: "6-character hex color (e.g. '548235')" }
            },
            required: ["name", "values"]
          },
          description: "Data series to plot."
        },
        showDataLabels: {
          type: "boolean",
          description: "Whether to display data labels inside bars/slices. Default false for clean bar charts."
        },
        targetParagraphIndex: {
          type: "number",
          description: "1-based paragraph index in Word document after which the chart should be inserted. Defaults to end of selection or current cursor position."
        }
      },
      required: ["categories", "series"]
    }
  },
  {
    name: "office_generate_chart_xml",
    description: "Generate a standalone Flat OPC OOXML package containing a native DrawingML chart with Aptos typography and custom corporate styling.",
    inputSchema: {
      type: "object",
      properties: {
        chartType: { type: "string", description: "bar, stacked_bar, column, pie, doughnut" },
        title: { type: "string" },
        categories: { type: "array", items: { type: "string" } },
        series: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              values: { type: "array", items: { type: "number" } },
              color: { type: "string" }
            },
            required: ["name", "values"]
          }
        },
        showDataLabels: { type: "boolean" },
        fontFamily: { type: "string", description: "Default 'Aptos'" },
        fontSize: { type: "number", description: "Default 9" }
      },
      required: ["categories", "series"]
    }
  },
  {
    name: "office_query",
    description: "Send a natural language instruction or query to the 0x15k Office AI model engine and receive natural language explanations plus atomic transactional Office operations.",
    inputSchema: {
      type: "object",
      properties: {
        prompt: {
          type: "string",
          description: "User instruction or prompt (e.g. 'Review grammar and format headers')."
        },
        host: {
          type: "string",
          enum: ["Word", "Excel", "PowerPoint"],
          description: "Active Office application. Default Word."
        },
        selectedText: {
          type: "string",
          description: "Currently selected text or cell range."
        },
        fullDocumentText: {
          type: "string",
          description: "Entire document or sheet context."
        },
        documentScope: {
          type: "string",
          enum: ["selection", "document"],
          description: "Scope of the edit."
        },
        agentMode: {
          type: "boolean",
          description: "True to generate editing operations; false for chat-only advice."
        },
        model: {
          type: "string",
          description: "Model override (e.g. 'gemini-3.8-flash-low', 'claude-3-5-sonnet', etc.)."
        }
      },
      required: ["prompt"]
    }
  },
  {
    name: "office_list_skills",
    description: "List all built-in and dynamically learned skills currently active in the Office AI runtime.",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    }
  },
  {
    name: "office_learn_skill",
    description: "Persist a new reusable skill into Office AI (saved as skills/learned-*.md) with instructions, triggers, and host targeting.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Skill display name." },
        description: { type: "string", description: "Short summary of what this skill does." },
        instructions: { type: "string", description: "Detailed markdown instructions and rules." },
        hosts: {
          type: "array",
          items: { type: "string" },
          description: "Target hosts: ['Word', 'Excel', 'PowerPoint']."
        },
        keywords: {
          type: "array",
          items: { type: "string" },
          description: "Keywords that trigger this skill."
        }
      },
      required: ["name", "description", "instructions"]
    }
  }
];

// Tool Handlers
async function handleToolCall(name, args = {}) {
  switch (name) {
    case "office_status": {
      const daemonStatus = await requestDaemon("GET", "/api/status");
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(daemonStatus, null, 2)
          }
        ]
      };
    }

    case "office_word_get_active_doc": {
      const includeSample = args.includeSampleText !== false;
      const count = Math.min(Math.max(Number(args.sampleParagraphCount) || 15, 1), 100);

      const appleScript = `
        set doc to active document
        set dName to name of doc
        set pCount to count of paragraphs of doc
        set wCount to count of words of doc
        set selText to content of text object of selection
        set docPath to ""
        try
          set docPath to POSIX path of (full name of doc as text)
        end try

        set sampleText to ""
        if ${includeSample} then
          set lim to ${count}
          if lim > pCount then set lim to pCount
          repeat with i from 1 to lim
            set sampleText to sampleText & (content of text object of paragraph i of doc) & linefeed
          end repeat
        end if

        return "NAME:" & dName & "\\nPATH:" & docPath & "\\nPARAGRAPHS:" & pCount & "\\nWORDS:" & wCount & "\\nSEL:" & selText & "\\nSAMPLE:\\n" & sampleText
      `;

      const raw = runWordAppleScript(appleScript);
      const lines = raw.split("\n");
      const result = {
        name: "",
        path: "",
        paragraphsCount: 0,
        wordsCount: 0,
        selection: "",
        sampleText: ""
      };

      let inSample = false;
      const sampleLines = [];

      for (const line of lines) {
        if (inSample) {
          sampleLines.push(line);
          continue;
        }
        if (line.startsWith("NAME:")) result.name = line.substring(5).trim();
        else if (line.startsWith("PATH:")) result.path = line.substring(5).trim();
        else if (line.startsWith("PARAGRAPHS:")) result.paragraphsCount = Number(line.substring(11).trim()) || 0;
        else if (line.startsWith("WORDS:")) result.wordsCount = Number(line.substring(6).trim()) || 0;
        else if (line.startsWith("SEL:")) result.selection = line.substring(4).trim();
        else if (line.startsWith("SAMPLE:")) inSample = true;
      }
      result.sampleText = sampleLines.join("\n").trim();

      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }]
      };
    }

    case "office_word_execute_applescript": {
      const script = String(args.script || "").trim();
      if (!script) throw new Error("Argument 'script' is required.");
      const res = runWordAppleScript(script);
      return {
        content: [{ type: "text", text: res }]
      };
    }

    case "office_generate_chart_xml": {
      const xml = chartGen.generateNativeChartFlatOpc({
        chartType: args.chartType,
        title: args.title,
        categories: args.categories,
        series: args.series,
        showDataLabels: args.showDataLabels,
        fontFamily: args.fontFamily,
        fontSize: args.fontSize
      });
      return {
        content: [{ type: "text", text: xml }]
      };
    }

    case "office_word_insert_chart": {
      const xml = chartGen.generateNativeChartFlatOpc({
        chartType: args.chartType || "stacked_bar",
        title: args.title || "",
        categories: args.categories,
        series: args.series,
        showDataLabels: args.showDataLabels === true
      });

      const tmpFile = path.join(os.tmpdir(), `office_chart_${Date.now()}.xml`);
      fs.writeFileSync(tmpFile, xml, "utf8");

      let targetPosScript = "";
      if (typeof args.targetParagraphIndex === "number" && args.targetParagraphIndex > 0) {
        targetPosScript = `
          set targetPara to paragraph ${args.targetParagraphIndex} of active document
          paste text object at end of text object of targetPara
        `;
      } else {
        targetPosScript = `
          paste text object at selection
        `;
      }

      const insertionScript = `
        set xmlFile to POSIX file "${tmpFile}"
        open xmlFile
        set tempDoc to active document
        set chartObj to text object of paragraph 1 of tempDoc
        copy text object of chartObj
        close tempDoc saving no

        ${targetPosScript}
        return "SUCCESS"
      `;

      const insertRes = runWordAppleScript(insertionScript);
      try {
        fs.unlinkSync(tmpFile);
      } catch (_) {}

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                ok: true,
                result: insertRes,
                chartType: args.chartType || "stacked_bar",
                categories: args.categories,
                seriesCount: (args.series || []).length
              },
              null,
              2
            )
          }
        ]
      };
    }

    case "office_query": {
      const chatRes = await requestDaemon("POST", "/api/chat", {
        prompt: args.prompt,
        host: args.host || "Word",
        selectedText: args.selectedText || "",
        fullDocumentText: args.fullDocumentText || "",
        documentScope: args.documentScope || "document",
        agentMode: args.agentMode !== false,
        model: args.model || ""
      });

      return {
        content: [{ type: "text", text: JSON.stringify(chatRes, null, 2) }]
      };
    }

    case "office_list_skills": {
      const skillsRes = await requestDaemon("GET", "/api/skills");
      return {
        content: [{ type: "text", text: JSON.stringify(skillsRes.skills || [], null, 2) }]
      };
    }

    case "office_learn_skill": {
      const learnRes = await requestDaemon("POST", "/api/skillfy", {
        name: args.name,
        description: args.description,
        instructions: args.instructions,
        hosts: args.hosts || ["Word", "Excel", "PowerPoint"],
        keywords: args.keywords || []
      });
      return {
        content: [{ type: "text", text: JSON.stringify(learnRes, null, 2) }]
      };
    }

    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

// JSON-RPC message dispatcher
function sendJsonRpc(obj) {
  process.stdout.write(JSON.stringify(obj) + "\n");
}

function sendResult(id, result) {
  sendJsonRpc({
    jsonrpc: "2.0",
    id,
    result
  });
}

function sendError(id, code, message, data = null) {
  sendJsonRpc({
    jsonrpc: "2.0",
    id,
    error: {
      code,
      message,
      ...(data ? { data } : {})
    }
  });
}

async function handleRequest(msg) {
  const { id, method, params } = msg;

  if (method === "initialize") {
    sendResult(id, {
      protocolVersion: PROTOCOL_VERSION,
      capabilities: {
        tools: {
          listChanged: false
        }
      },
      serverInfo: {
        name: SERVER_NAME,
        version: SERVER_VERSION
      }
    });
    return;
  }

  if (method === "notifications/initialized") {
    logDebug("Client initialized successfully.");
    return;
  }

  if (method === "ping") {
    sendResult(id, {});
    return;
  }

  if (method === "tools/list") {
    sendResult(id, { tools: TOOLS });
    return;
  }

  if (method === "tools/call") {
    const { name, arguments: toolArgs } = params || {};
    try {
      const callRes = await handleToolCall(name, toolArgs || {});
      sendResult(id, callRes);
    } catch (err) {
      logDebug(`Tool call '${name}' failed:`, err.message);
      sendResult(id, {
        content: [
          {
            type: "text",
            text: `Error executing tool '${name}': ${err.message}`
          }
        ],
        isError: true
      });
    }
    return;
  }

  // Unknown method
  if (id !== undefined) {
    sendError(id, -32601, `Method not found: ${method}`);
  }
}

function startServer() {
  logDebug(`Starting 0x15k Office AI MCP Server v${SERVER_VERSION} over stdio...`);

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
    terminal: false
  });

  rl.on("line", (line) => {
    const trimmed = line.trim();
    if (!trimmed) return;
    try {
      const msg = JSON.parse(trimmed);
      handleRequest(msg).catch((err) => {
        logDebug("Unhandled request error:", err.message);
        if (msg.id !== undefined) {
          sendError(msg.id, -32603, `Internal error: ${err.message}`);
        }
      });
    } catch (err) {
      logDebug("JSON parse error:", err.message);
      sendError(null, -32700, `Parse error: ${err.message}`);
    }
  });

  process.on("SIGINT", () => process.exit(0));
  process.on("SIGTERM", () => process.exit(0));
}

startServer();

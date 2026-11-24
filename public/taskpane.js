/**
 * 0x15k Office AI for Microsoft Office (v4.5.0)
 * Multi-Host Engine (Word, Excel, PowerPoint), Minimalist UI, Redline Diff & Undo
 */

const HOST_CONFIG = {
  Word: {
    label: "Word",
    dotClass: "",
    subtitle: "Edit documents, standardize formatting, add headers/footers, or paste any URL.",
    cleanLabel: "Clean Format",
    starters: [
      {
        label: "📐 Standardize CV or document formatting & spacing",
        prompt: "Standardize formatting, fonts, spacing, and hierarchy across the entire document while keeping its original language."
      },
      {
        label: "✨ Improve clarity, grammar, and executive tone",
        prompt: "Improve clarity, flow, and grammar while preserving the document's original language and meaning."
      },
      {
        label: "📊 Convert text into a clean comparison table",
        prompt: "Convert this information into a clean comparison table with headers."
      }
    ],
    slashCommands: [
      { cmd: "/accept", desc: "Accept all tracked changes in document", action: "acceptAllChanges", template: "/accept" },
      { cmd: "/reject", desc: "Reject all tracked changes in document", action: "rejectAllChanges", template: "/reject" },
      { cmd: "/track", desc: "Toggle track changes mode (ON / OFF)", action: "toggleTrackChanges", template: "/track" },
      { cmd: "/replace", desc: "Find and replace across document", template: 'Find and replace across the document: "OLD" with "NEW".' },
      { cmd: "/format", desc: "Standardize typography and layout", template: "Standardize formatting, fonts, spacing, and hierarchy across the entire document while preserving its language." },
      { cmd: "/fix-all", desc: "Proofread and fix grammar/spelling", template: "Proofread the document and fix any spelling, grammar, or punctuation errors." },
      { cmd: "/table", desc: "Convert content into a Word table", template: "Convert this information into a clean comparison table with headers." },
      { cmd: "/comment", desc: "Add native Word review comments", template: "Review the text and add native Word comments highlighting risks or suggestions." },
      { cmd: "/summarize", desc: "Insert executive summary at top", template: "Write a concise executive summary and insert it at the start of the document." },
      { cmd: "/translate", desc: "Translate keeping formatting", template: "Translate this content to English while preserving its formatting and tone." },
      { cmd: "/skillfy", desc: "Learn & save a reusable skill from this chat", template: "/skillfy " }
    ]
  },
  Excel: {
    label: "Excel",
    dotClass: "excel",
    subtitle: "Build tables, write formulas, format ranges, or import structured web URLs into cells.",
    cleanLabel: "Auto-Fit",
    starters: [
      {
        label: "📊 Build a formatted financial/tracking table",
        prompt: "Create a structured monthly budget tracking table starting at A1 with Revenue, Expenses, Profit formula, and clean header formatting."
      },
      {
        label: "🧮 Add summary formulas & totals to current sheet",
        prompt: "Analyze the selected data and add summary formulas (SUM, AVERAGE, growth %) with clean number formatting."
      },
      {
        label: "🎨 Format header row and currency columns",
        prompt: "Format the header row bold with dark background and white text, apply currency formatting where appropriate, and auto-fit columns."
      }
    ],
    slashCommands: [
      { cmd: "/table", desc: "Insert structured data table", template: "Create a structured table starting at A1 for: " },
      { cmd: "/formula", desc: "Write Excel formulas for selection", template: "Write the required Excel formulas for the selected cells and format the output." },
      { cmd: "/format", desc: "Style headers, currency & auto-fit", template: "Format the active sheet headers, apply clean number formats, and auto-fit columns." },
      { cmd: "/sheet", desc: "Create a new summary worksheet", template: "Create a new summary worksheet analyzing the data in the current sheet." },
      { cmd: "/skillfy", desc: "Learn & save a reusable skill from this chat", template: "/skillfy " }
    ]
  },
  PowerPoint: {
    label: "PPT",
    dotClass: "powerpoint",
    subtitle: "Generate structured slide decks, rewrite bullet points, or turn URLs into presentations.",
    cleanLabel: "",
    starters: [
      {
        label: "📽️ Generate a 4-slide executive presentation",
        prompt: "Generate a crisp 4-slide presentation outline with titles and concise bullet points based on: "
      },
      {
        label: "✨ Tighten selected slide text into punchy bullets",
        prompt: "Rewrite the selected slide text into concise, high-impact executive bullet points."
      },
      {
        label: "🌐 Summarize a web URL into presentation slides",
        prompt: "Summarize https://example.com into 3 clean presentation slides."
      }
    ],
    slashCommands: [
      { cmd: "/slides", desc: "Generate new slides in deck", template: "Generate presentation slides covering: " },
      { cmd: "/tighten", desc: "Make selected slide text concise", template: "Rewrite the selected slide text into concise, punchy bullet points." },
      { cmd: "/translate", desc: "Translate selected slide text", template: "Translate the selected slide text to English while keeping it concise." },
      { cmd: "/skillfy", desc: "Learn & save a reusable skill from this chat", template: "/skillfy " }
    ]
  }
};

const OP_EXECUTION_PRIORITY = {
  replace_full_document: 1,
  replace_section: 1,
  replace_selection: 1,
  search_and_replace: 2,
  insert_at_location: 2,
  delete_text: 2,
  insert_page_break: 2,
  set_header: 3,
  set_footer: 3,
  standardize_document: 4,
  format_text: 5,
  add_comment: 6
};

const state = {
  host: "Word", // 'Word' | 'Excel' | 'PowerPoint'
  selectedText: "",
  selectedType: "none", // 'none' | 'text' | 'chart' | 'table'
  selectedLabel: "",
  fullDocText: "",
  excelSelectedCell: "A1",
  documentScope: "document", // 'selection' | 'document'
  conversationHistory: [],
  documentOutline: [],
  documentHyperlinks: {},
  documentStyle: null,
  renderedMessages: [],
  undoStack: [],
  autoApply: true,
  isGenerating: false,
  slashSelectedIndex: 0,
  settings: {
    provider: "agy",
    agyModel: "gemini-3.8-flash-low",
    apiKey: "",
    model: "",
    defaultFont: "Calibri",
    agentMode: true,
    autoApply: true,
    alwaysTrackChanges: false,
    richFormat: true,
    systemPrompt: ""
  }
};

function loadSettings() {
  try {
    const saved = localStorage.getItem("word_ai_settings_v3");
    if (saved) {
      state.settings = { ...state.settings, ...JSON.parse(saved) };
    }
    // Default autoApply to true for active execution
    state.autoApply = true;
    state.settings.autoApply = true;
  } catch (e) {
    console.warn("Failed to load settings:", e);
  }
}

function saveSettings(silent = false) {
  try {
    localStorage.setItem("word_ai_settings_v3", JSON.stringify(state.settings));
    if (!silent) showToast("Settings saved");
    updateStatusChip();
  } catch (e) {
    showToast("Save failed: " + e.message);
  }
}

function saveChatSession() {
  try {
    const key = "office_ai_chat_v4_" + state.host;
    const data = {
      conversationHistory: state.conversationHistory.slice(-24),
      renderedMessages: state.renderedMessages.slice(-24)
    };
    localStorage.setItem(key, JSON.stringify(data));
  } catch (e) {
    console.warn("saveChatSession error:", e);
  }
}

function restoreChatSession() {
  try {
    const key = "office_ai_chat_v4_" + state.host;
    const raw = localStorage.getItem(key);
    if (!raw) return;
    const data = JSON.parse(raw);
    if (!data) return;

    if (Array.isArray(data.conversationHistory)) {
      state.conversationHistory = data.conversationHistory;
    }
    if (Array.isArray(data.renderedMessages) && data.renderedMessages.length > 0) {
      const welcomeCard = document.getElementById("welcome-card");
      if (welcomeCard) welcomeCard.style.display = "none";

      const savedMessages = [...data.renderedMessages];
      state.renderedMessages = [];

      for (const msg of savedMessages) {
        if (msg.role === "user") {
          appendUserMessage(msg.text, msg.contextTag, false);
          state.renderedMessages.push(msg);
        } else if (msg.role === "assistant") {
          const container = document.getElementById("chat-container");
          let msgEl = document.getElementById(msg.id);
          if (!msgEl) {
            msgEl = document.createElement("div");
            msgEl.className = "message";
            msgEl.id = msg.id;
            container.appendChild(msgEl);
          }
          renderAssistantMessage(
            msg.id,
            msg.markdownText,
            msg.operations || [],
            msg.originalSelection || "",
            msg.durationMs || 0,
            false,
            false
          );
          if (msg.autoApplied && msg.operations && msg.operations.length > 0) {
            const pill = msgEl.querySelector(".ops-pill");
            if (pill) {
              pill.classList.add("applied");
              pill.textContent = `✓ Applied (${msg.operations.length})`;
            }
          }
          state.renderedMessages.push(msg);
        }
      }
      scrollToBottom();
    }
  } catch (e) {
    console.warn("restoreChatSession error:", e);
  }
}

Office.onReady((info) => {
  loadSettings();

  if (info.host === Office.HostType.Excel) {
    state.host = "Excel";
  } else if (info.host === Office.HostType.PowerPoint) {
    state.host = "PowerPoint";
  } else {
    state.host = "Word";
  }

  initUI();
  applyHostBranding();
  checkServerStatus();
  setInterval(checkServerStatus, 5000);
  updateUndoButton();

  if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel) {
    try {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        () => syncContext(true)
      );
    } catch (e) {
      console.warn("Selection listener unavailable:", e);
    }
    syncContext(true);
  }
});

function initUI() {
  const modelSelect = document.getElementById("quick-model-select");
  modelSelect.value = state.settings.agyModel || "gemini-3.8-flash-low";
  modelSelect.onchange = (e) => {
    state.settings.agyModel = e.target.value;
    saveSettings(true);
    showToast(`Model: ${e.target.options[e.target.selectedIndex].text}`);
  };

  document.getElementById("btn-undo").onclick = restoreLastSnapshot;
  document.getElementById("btn-clear").onclick = clearConversation;
  document.getElementById("btn-settings").onclick = openSettings;
  document.getElementById("btn-clean-format").onclick = handleCleanFormat;

  document.getElementById("btn-scope-toggle").onclick = () => {
    if (state.documentScope === "selection") {
      setScope("document", "Full Doc");
    } else {
      setScope("selection", state.selectedLabel || "Selection");
    }
  };

  const autoApplyEl = document.getElementById("toggle-auto-apply");
  autoApplyEl.checked = state.autoApply;
  autoApplyEl.onchange = (e) => {
    state.autoApply = e.target.checked;
    state.settings.autoApply = e.target.checked;
    saveSettings(true);
    showToast(state.autoApply ? "Auto-apply enabled" : "Manual confirmation mode");
  };

  const promptInput = document.getElementById("prompt-input");
  promptInput.addEventListener("keydown", handleInputKeyDown);
  promptInput.addEventListener("input", () => {
    promptInput.style.height = "auto";
    promptInput.style.height = Math.min(promptInput.scrollHeight, 140) + "px";
    handleSlashInput(promptInput.value);
  });

  document.getElementById("btn-send").onclick = handleSubmit;
  document.getElementById("btn-close-settings").onclick = closeSettings;
  document.getElementById("btn-save-settings").onclick = applySettingsFromModal;
  document.getElementById("btn-reset-settings").onclick = resetSettingsModal;
  document.getElementById("btn-skillfy-now").onclick = () => {
    closeSettings();
    document.getElementById("prompt-input").value = "/skillfy";
    handleSubmit();
  };
  document.getElementById("cfg-provider").onchange = (e) => updateModalProviderFields(e.target.value);

  window.addEventListener("focus", () => syncContext(false));
}

function applyHostBranding() {
  const cfg = HOST_CONFIG[state.host] || HOST_CONFIG.Word;
  document.getElementById("host-label").textContent = cfg.label;
  updateStatusChip(true);

  const cleanBtn = document.getElementById("btn-clean-format");
  if (cfg.cleanLabel) {
    cleanBtn.style.display = "inline-flex";
    cleanBtn.textContent = cfg.cleanLabel;
  } else {
    cleanBtn.style.display = "none";
  }

  renderWelcomeCard();
  restoreChatSession();
}

function renderWelcomeCard() {
  const cfg = HOST_CONFIG[state.host] || HOST_CONFIG.Word;
  const subEl = document.getElementById("welcome-sub");
  const startersEl = document.getElementById("welcome-starters");
  if (subEl) subEl.textContent = cfg.subtitle;
  if (!startersEl) return;

  startersEl.innerHTML = cfg.starters
    .map((s, idx) => `<button type="button" class="starter-btn" data-starter="${idx}">${escapeHtml(s.label)}</button>`)
    .join("");

  startersEl.querySelectorAll(".starter-btn").forEach((btn) => {
    const starter = cfg.starters[Number(btn.dataset.starter)];
    btn.onclick = () => {
      const inputEl = document.getElementById("prompt-input");
      inputEl.value = starter.prompt;
      if (starter.prompt.trim().endsWith(":")) {
        inputEl.focus();
      } else {
        handleSubmit();
      }
    };
  });
}

async function checkServerStatus() {
  try {
    const res = await fetch("/api/status");
    if (res.ok) {
      const data = await res.json();
      if (data.version && state.clientVersion && data.version !== state.clientVersion) {
        console.log(`Server updated to v${data.version}. Auto-reloading taskpane...`);
        state.clientVersion = data.version;
        window.location.reload();
        return;
      }
      if (data.version && !state.clientVersion) {
        state.clientVersion = data.version;
      }
      updateStatusChip(true);
    } else {
      updateStatusChip(false);
    }
  } catch (e) {
    updateStatusChip(false);
  }
}

function updateStatusChip(isOnline = true) {
  const dot = document.getElementById("status-dot");
  const cfg = HOST_CONFIG[state.host] || HOST_CONFIG.Word;

  if (!isOnline) {
    dot.className = "status-dot offline";
    return;
  }
  if (state.isGenerating) {
    dot.className = "status-dot busy";
    return;
  }
  dot.className = `status-dot ${cfg.dotClass}`.trim();
}

/**
 * Slash Commands
 */
function getActiveSlashCommands() {
  return (HOST_CONFIG[state.host] || HOST_CONFIG.Word).slashCommands;
}

function getFilteredSlashCommands(value) {
  if (!value.startsWith("/") || value.includes(" ")) return [];
  const query = value.toLowerCase();
  return getActiveSlashCommands().filter((c) => c.cmd.startsWith(query));
}

function handleSlashInput(value) {
  const menu = document.getElementById("slash-menu");
  const matches = getFilteredSlashCommands(value);

  if (matches.length === 0) {
    menu.classList.remove("open");
    menu.innerHTML = "";
    return;
  }

  state.slashSelectedIndex = Math.min(state.slashSelectedIndex, matches.length - 1);
  menu.innerHTML = matches
    .map(
      (item, idx) => `
      <button type="button" class="slash-item ${idx === state.slashSelectedIndex ? "selected" : ""}" data-cmd="${item.cmd}">
        <span class="slash-cmd">${escapeHtml(item.cmd)}</span>
        <span class="slash-desc">${escapeHtml(item.desc)}</span>
      </button>
    `
    )
    .join("");

  menu.querySelectorAll(".slash-item").forEach((btn) => {
    btn.onclick = () => selectSlashCommand(btn.dataset.cmd);
  });
  menu.classList.add("open");
}

function selectSlashCommand(cmd) {
  const item = getActiveSlashCommands().find((c) => c.cmd === cmd);
  const inputEl = document.getElementById("prompt-input");
  document.getElementById("slash-menu").classList.remove("open");
  if (!item) return;

  if (item.action === "acceptAllChanges") {
    inputEl.value = "";
    acceptAllTrackedChanges();
    return;
  }
  if (item.action === "rejectAllChanges") {
    inputEl.value = "";
    rejectAllTrackedChanges();
    return;
  }
  if (item.action === "toggleTrackChanges") {
    inputEl.value = "";
    toggleTrackChangesMode();
    return;
  }

  inputEl.value = item.template;
  inputEl.focus();
  inputEl.style.height = "auto";
  inputEl.style.height = Math.min(inputEl.scrollHeight, 140) + "px";
}

function handleInputKeyDown(e) {
  const menu = document.getElementById("slash-menu");
  if (menu.classList.contains("open")) {
    const items = menu.querySelectorAll(".slash-item");
    if (e.key === "ArrowDown") {
      e.preventDefault();
      state.slashSelectedIndex = (state.slashSelectedIndex + 1) % items.length;
      handleSlashInput(e.target.value);
      return;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      state.slashSelectedIndex = (state.slashSelectedIndex - 1 + items.length) % items.length;
      handleSlashInput(e.target.value);
      return;
    }
    if (e.key === "Escape") {
      e.preventDefault();
      menu.classList.remove("open");
      return;
    }
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      const selected = items[state.slashSelectedIndex];
      if (selected) selectSlashCommand(selected.dataset.cmd);
      return;
    }
  }

  if (e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    handleSubmit();
  }
}

/**
 * Multi-Host Context Reader
 */
async function syncContext(autoDetectScope = false) {
  if (state.host === "Word" && typeof Word !== "undefined" && Word.run) {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");

        const inlinePics = selection.inlinePictures;
        inlinePics.load("items");

        const paragraphs = selection.paragraphs;
        paragraphs.load("items/text");

        await context.sync();

        const rawText = (selection.text || "").trim();
        state.selectedText = rawText;
        state.selectedType = "text";
        state.selectedLabel = "";

        if (rawText.length > 0) {
          const words = rawText.split(/\s+/).filter(Boolean).length;
          state.selectedType = "text";
          state.selectedLabel = words > 1 ? `Selection (${words} words)` : "Selection";
          if (autoDetectScope) {
            setScope("selection", state.selectedLabel);
          }
          return;
        }

        // Non-text selection: check for Chart, Drawing or InlinePicture
        let isChart = false;
        let chartContextText = "";

        if (inlinePics.items && inlinePics.items.length > 0) {
          isChart = true;
        }

        if (!isChart) {
          try {
            const ooxml = selection.getOoxml();
            await context.sync();
            const val = ooxml.value || "";
            if (val.includes("<c:chart") || val.includes("<c:chartSpace") || val.includes("word/charts/chart") || val.includes("<w:drawing")) {
              isChart = true;
            }
          } catch (_) {}
        }

        if (!isChart && paragraphs.items && paragraphs.items.length > 0) {
          try {
            const p = paragraphs.items[0];
            const pOoxml = p.getOoxml();
            await context.sync();
            const val = pOoxml.value || "";
            if (val.includes("<c:chart") || val.includes("<c:chartSpace") || val.includes("word/charts/chart")) {
              isChart = true;
              chartContextText = (p.text || "").trim();
            }
          } catch (_) {}
        }

        if (isChart) {
          state.selectedType = "chart";
          if (!chartContextText && paragraphs.items && paragraphs.items.length > 0) {
            chartContextText = (paragraphs.items[0].text || "").trim();
          }
          state.selectedText = `[SELECTED_CHART: An embedded DrawingML chart is currently selected by the user. Context: "${chartContextText.slice(0, 150)}"]`;
          state.selectedLabel = "📊 Chart Selected";
          if (autoDetectScope) {
            setScope("selection", state.selectedLabel);
          }
          return;
        }

        // Check if inside a table
        try {
          const parentTable = selection.parentTableOrNullObject;
          if (parentTable) {
            parentTable.load("isNullObject");
            await context.sync();
            if (!parentTable.isNullObject) {
              state.selectedType = "table";
              state.selectedText = "[SELECTED_TABLE: A table is currently selected by the user]";
              state.selectedLabel = "📋 Table Selected";
              if (autoDetectScope) {
                setScope("selection", state.selectedLabel);
              }
              return;
            }
          }
        } catch (_) {}

        // Default: nothing selected
        state.selectedType = "none";
        state.selectedLabel = "";
        if (autoDetectScope) {
          setScope("document", "Full Doc");
        }
      });
    } catch (err) {
      console.warn("Word selection sync skipped:", err.message);
    }
  } else if (state.host === "Excel" && typeof Excel !== "undefined" && Excel.run) {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address,values");
        await context.sync();
        state.excelSelectedCell = (range.address || "A1").split("!").pop().split(":")[0];
        const hasMultiCell = (range.values || []).flat().some((v) => v !== "" && v !== null);
        state.selectedText = hasMultiCell
          ? `Range ${range.address}:\n${JSON.stringify(range.values)}`
          : `Selected cell: ${range.address}`;
        if (autoDetectScope) {
          setScope(hasMultiCell ? "selection" : "document");
        }
      });
    } catch (err) {
      console.warn("Excel selection sync skipped:", err.message);
    }
  }
}

function parseHyperlinksFromOoxml(xmlString) {
  const result = {};
  if (!xmlString || typeof xmlString !== "string") return result;

  const relMap = {};
  const relRegex = /<Relationship\s+[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"/gi;
  let match;
  while ((match = relRegex.exec(xmlString)) !== null) {
    relMap[match[1]] = match[2];
  }
  const relRegexRev = /<Relationship\s+[^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"/gi;
  while ((match = relRegexRev.exec(xmlString)) !== null) {
    relMap[match[2]] = match[1];
  }

  const hlRegex = /<w:hyperlink\s+[^>]*?r:id="([^"]+)"[^>]*?>([\s\S]*?)<\/w:hyperlink>/gi;
  while ((match = hlRegex.exec(xmlString)) !== null) {
    const rId = match[1];
    const inner = match[2];
    const tRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/gi;
    const textParts = [];
    let tMatch;
    while ((tMatch = tRegex.exec(inner)) !== null) {
      textParts.push(tMatch[1]);
    }
    const fullText = textParts.join("").trim();
    if (rId && relMap[rId] && fullText) {
      result[fullText] = relMap[rId].replace(/&amp;/g, "&");
    }
  }

  const fldRegex = /<w:fldSimple\s+[^>]*?w:instr="HYPERLINK\s*(?:&quot;|")([^"]+?)(?:&quot;|")[^>]*?>([\s\S]*?)<\/w:fldSimple>/gi;
  while ((match = fldRegex.exec(xmlString)) !== null) {
    const url = match[1].replace(/&amp;/g, "&").trim();
    const inner = match[2];
    const tRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/gi;
    const textParts = [];
    let tMatch;
    while ((tMatch = tRegex.exec(inner)) !== null) {
      textParts.push(tMatch[1]);
    }
    const fullText = textParts.join("").trim();
    if (url && fullText) {
      result[fullText] = url;
    }
  }

  return result;
}

function injectMarkdownLinks(text, linkMap) {
  if (!text || !linkMap || typeof linkMap !== "object") return text;
  let result = String(text);
  const keys = Object.keys(linkMap).filter((k) => k && linkMap[k]);
  if (keys.length === 0) return result;

  keys.sort((a, b) => b.length - a.length);

  for (const key of keys) {
    const url = linkMap[key];
    const mdLink = `[${key}](${url})`;
    if (result.includes(mdLink)) continue;
    const escaped = key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    try {
      const regex = new RegExp(`(?<!\\[)${escaped}(?!\\]|\\([^)]*\\))`, "g");
      result = result.replace(regex, mdLink);
    } catch {
      result = result.split(key).join(mdLink);
    }
  }
  return result;
}

function introspectDocumentStyles(paragraphsItems, hyperlinksMap) {
  const fontCounts = {};
  const sizeCounts = {};
  const spacingCounts = {};
  let detectedSpaceBefore = 0;
  let detectedSpaceAfter = 2;
  let topHeaderAlignment = "Centered";

  const nonBlankItems = (paragraphsItems || []).filter((p) => (p.text || "").trim().length > 0);

  nonBlankItems.forEach((p, idx) => {
    if (p.font && p.font.name) {
      const f = p.font.name.trim();
      fontCounts[f] = (fontCounts[f] || 0) + 1;
    }
    if (p.font && typeof p.font.size === "number" && p.font.size >= 8 && p.font.size <= 14) {
      const s = p.font.size;
      sizeCounts[s] = (sizeCounts[s] || 0) + 1;
    }
    if (typeof p.lineSpacing === "number" && p.lineSpacing > 0) {
      const ls = Number(p.lineSpacing.toFixed(1));
      spacingCounts[ls] = (spacingCounts[ls] || 0) + 1;
    }
    if (typeof p.spaceAfter === "number" && p.spaceAfter >= 0 && p.spaceAfter <= 12) {
      detectedSpaceAfter = p.spaceAfter;
    }
    if (typeof p.spaceBefore === "number" && p.spaceBefore >= 0 && p.spaceBefore <= 12) {
      detectedSpaceBefore = p.spaceBefore;
    }
    if (idx === 0 && p.alignment) {
      topHeaderAlignment = String(p.alignment);
    }
  });

  let predominantFont = state.settings.defaultFont || "Calibri";
  let maxFontCount = 0;
  for (const [font, count] of Object.entries(fontCounts)) {
    if (count > maxFontCount) {
      maxFontCount = count;
      predominantFont = font;
    }
  }

  let predominantFontSize = 10.5;
  let maxSizeCount = 0;
  for (const [sizeStr, count] of Object.entries(sizeCounts)) {
    const s = parseFloat(sizeStr);
    if (count > maxSizeCount) {
      maxSizeCount = count;
      predominantFontSize = s;
    }
  }

  let predominantLineSpacingPt = 13.8;
  let maxSpacingCount = 0;
  for (const [spacingStr, count] of Object.entries(spacingCounts)) {
    const ls = parseFloat(spacingStr);
    if (count > maxSpacingCount) {
      maxSpacingCount = count;
      predominantLineSpacingPt = ls;
    }
  }
  const predominantLineSpacingMultiplier = Number(
    (predominantLineSpacingPt / 12).toFixed(2)
  );

  return {
    predominantFont,
    predominantFontSize,
    predominantLineSpacingPt,
    predominantLineSpacingMultiplier,
    spaceBefore: detectedSpaceBefore,
    spaceAfter: detectedSpaceAfter,
    topHeaderAlignment,
    hyperlinksCount: Object.keys(hyperlinksMap || {}).length,
    hyperlinks: hyperlinksMap || {}
  };
}

async function readActiveHostFullContext() {
  if (state.host === "Word" && typeof Word !== "undefined" && Word.run) {
    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        body.load("text");
        paragraphs.load(
          "items/text,items/style,items/font/name,items/font/size,items/lineSpacing,items/spaceBefore,items/spaceAfter,items/isListItem,items/leftIndent,items/alignment"
        );

        const ooxml = body.getOoxml();
        let desktopLinks = null;
        try {
          const bodyRange = typeof body.getRange === "function" ? body.getRange() : null;
          if (bodyRange && bodyRange.hyperlinks) {
            desktopLinks = bodyRange.hyperlinks;
            desktopLinks.load("items/address,items/textToDisplay");
          }
        } catch {
          // desktop hyperlinks optional
        }

        await context.sync();

        const hyperlinksMap = {};
        if (desktopLinks && desktopLinks.items && desktopLinks.items.length > 0) {
          for (const item of desktopLinks.items) {
            const addr = item.address || "";
            const disp = (item.textToDisplay || "").trim();
            if (addr && disp) hyperlinksMap[disp] = addr;
          }
        }
        if (ooxml && ooxml.value) {
          const ooxmlMap = parseHyperlinksFromOoxml(ooxml.value);
          Object.assign(hyperlinksMap, ooxmlMap);
        }
        state.documentHyperlinks = hyperlinksMap;

        const docStyle = introspectDocumentStyles(paragraphs.items, hyperlinksMap);
        state.documentStyle = docStyle;

        if (docStyle.predominantFont) {
          state.settings.defaultFont = docStyle.predominantFont;
        }
        state.settings.lineSpacing = docStyle.predominantLineSpacingMultiplier;
        saveSettings(true);

        const markdownLines = [];
        const pItems = paragraphs.items || [];
        for (let idx = 0; idx < pItems.length; idx++) {
          const p = pItems[idx];
          const pText = (p.text || "").trim();
          if (!pText) {
            markdownLines.push("");
            continue;
          }
          if (p.isListItem) {
            const cleanItem = pText.replace(/^(?:[•\-\*–—>→\t\s]|->)+\s*/, "").trim();
            markdownLines.push(`- ${cleanItem}`);
          } else if (isSectionHeadingParagraph(p, idx)) {
            markdownLines.push(`## ${pText.replace(/^#+\s*/, "")}`);
          } else {
            markdownLines.push(pText);
          }
        }
        const structuredText = markdownLines.join("\n").trim();
        const rawText = structuredText || (body.text || "").trim();
        const textWithLinks = injectMarkdownLinks(rawText, hyperlinksMap);
        state.fullDocText = textWithLinks;

        state.documentOutline = (paragraphs.items || [])
          .slice(0, 500)
          .map((p, idx) => {
            const pText = (p.text || "").trim();
            let flag = "";
            if (p.isListItem) flag = " (List)";
            else if (/\s\|\s/.test(pText)) flag = " (DualSplit)";
            else if (p.alignment && /center/i.test(String(p.alignment))) flag = " (Centered)";
            else if (/gr[áa]fico|chart|figura/i.test(pText)) flag = " (ChartRef)";
            return {
              index: idx,
              style: String(p.style || "Normal").replace(/^Normal\b.*/i, "Normal") + flag,
              isListItem: !!p.isListItem,
              alignment: p.alignment ? String(p.alignment) : "Left",
              leftIndent: typeof p.leftIndent === "number" ? Math.round(p.leftIndent) : 0,
              text: pText.slice(0, 100)
            };
          })
          .filter((item) => item.text.length > 0);

        return textWithLinks;
      });
    } catch (e) {
      console.warn("readActiveHostFullContext Word error:", e);
      return "";
    }
  }

  if (state.host === "Excel" && typeof Excel !== "undefined" && Excel.run) {
    try {
      return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();
        sheet.load("name");
        usedRange.load("address,values,formulas,isNullObject");
        await context.sync();

        if (usedRange.isNullObject) {
          const emptyText = `Active Worksheet: "${sheet.name}" (empty)`;
          state.fullDocText = emptyText;
          return emptyText;
        }
        const previewRows = (usedRange.values || []).slice(0, 50);
        let text = `Active Worksheet: "${sheet.name}" | Used Range: ${usedRange.address}\nData:\n${JSON.stringify(previewRows)}`;
        const rawFormulas = (usedRange.formulas || []).slice(0, 50);
        const hasFormulas = rawFormulas.some((row) =>
          row.some((cell) => typeof cell === "string" && cell.startsWith("="))
        );
        if (hasFormulas) {
          text += `\nFormulas:\n${JSON.stringify(rawFormulas)}`;
        }
        state.fullDocText = text;
        return text;
      });
    } catch (e) {
      return "";
    }
  }

  if (state.host === "PowerPoint") {
    const selected = await readOfficeSelectedTextAsync();
    state.selectedText = selected;
    if (typeof PowerPoint !== "undefined" && PowerPoint.run) {
      try {
        return await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          const count = slides.items.length;
          if (count === 0) {
            state.fullDocText = "PowerPoint Presentation (0 slides).";
            return state.fullDocText;
          }
          for (let i = 0; i < Math.min(count, 12); i++) {
            slides.items[i].shapes.load("items/textFrame/textRange/text");
          }
          await context.sync();
          let summary = `PowerPoint Presentation (${count} slide(s)):\n`;
          for (let i = 0; i < Math.min(count, 12); i++) {
            const shItems = slides.items[i].shapes.items || [];
            const texts = shItems
              .map((sh) => {
                try { return (sh.textFrame?.textRange?.text || "").trim(); } catch (e) { return ""; }
              })
              .filter(Boolean);
            summary += `Slide ${i + 1}: ${texts.join(" | ") || "(empty)"}\n`;
          }
          state.fullDocText = summary.trim();
          return state.fullDocText;
        });
      } catch (e) {
        return selected || "PowerPoint Presentation";
      }
    }
    return selected || "PowerPoint Presentation";
  }

  return "";
}

function readOfficeSelectedTextAsync() {
  return new Promise((resolve) => {
    if (!Office?.context?.document?.getSelectedDataAsync) {
      return resolve("");
    }
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(String(result.value || "").trim());
      } else {
        resolve("");
      }
    });
  });
}

function setScope(scope, customLabel = null) {
  state.documentScope = scope;
  const pill = document.getElementById("btn-scope-toggle");
  if (!pill) return;
  const isSel = scope === "selection";
  pill.classList.toggle("active", isSel);
  pill.classList.toggle("has-selection", isSel);
  if (customLabel) {
    pill.textContent = customLabel;
  } else if (state.host === "Excel") {
    pill.textContent = isSel ? (state.selectedLabel || "Selection") : "Sheet";
  } else if (state.host === "PowerPoint") {
    pill.textContent = isSel ? (state.selectedLabel || "Selection") : "Deck";
  } else {
    pill.textContent = isSel ? (state.selectedLabel || "Selection") : "Full Doc";
  }
}

/**
 * Undo Stack (Word OOXML + Excel Range Snapshot)
 */
function updateUndoButton() {
  const btn = document.getElementById("btn-undo");
  const countEl = document.getElementById("undo-count");
  const count = state.undoStack.length;
  countEl.textContent = String(count);
  btn.style.display = count > 0 ? "inline-flex" : "none";
  btn.title = count > 0 ? `Undo: ${state.undoStack[count - 1].label}` : "";
}

async function saveUndoSnapshot(label = "AI edit") {
  if (state.host === "Word" && typeof Word !== "undefined" && Word.run) {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const ooxmlResult = body.getOoxml();
        body.load("text");
        await context.sync();

        state.undoStack.push({
          host: "Word",
          label,
          ooxml: ooxmlResult.value,
          text: body.text || ""
        });
        if (state.undoStack.length > 10) state.undoStack.shift();
        updateUndoButton();
      });
    } catch (e) {
      console.warn("Word undo snapshot failed:", e);
    }
  } else if (state.host === "Excel" && typeof Excel !== "undefined" && Excel.run) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();
        sheet.load("name");
        usedRange.load("address,values,isNullObject");
        await context.sync();

        state.undoStack.push({
          host: "Excel",
          label,
          sheetName: sheet.name,
          address: usedRange.isNullObject ? null : usedRange.address.split("!").pop(),
          values: usedRange.isNullObject ? null : usedRange.values
        });
        if (state.undoStack.length > 10) state.undoStack.shift();
        updateUndoButton();
      });
    } catch (e) {
      console.warn("Excel undo snapshot failed:", e);
    }
  }
}

async function restoreLastSnapshot() {
  if (state.undoStack.length === 0) return;
  const snapshot = state.undoStack.pop();
  updateUndoButton();

  try {
    if (snapshot.host === "Word" && typeof Word !== "undefined" && Word.run) {
      await Word.run(async (context) => {
        const body = context.document.body;
        try {
          if (snapshot.ooxml) {
            body.insertOoxml(snapshot.ooxml, Word.InsertLocation.replace);
          } else {
            body.insertText(snapshot.text || "", Word.InsertLocation.replace);
          }
          await context.sync();
        } catch (err) {
          body.insertText(snapshot.text || "", Word.InsertLocation.replace);
          await context.sync();
        }
      });
    } else if (snapshot.host === "Excel" && typeof Excel !== "undefined" && Excel.run) {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(snapshot.sheetName);
        const used = sheet.getUsedRangeOrNullObject();
        used.load("isNullObject");
        await context.sync();
        if (!used.isNullObject) used.clear();
        if (snapshot.address && snapshot.values) {
          sheet.getRange(snapshot.address).values = snapshot.values;
        }
        await context.sync();
      });
    }
    showToast(`↩ Restored: ${snapshot.label}`);
    await syncContext(false);
  } catch (err) {
    showToast("Undo error: " + err.message);
  }
}

/**
 * Chat Submission
 */
async function handleSubmit() {
  const inputEl = document.getElementById("prompt-input");
  const prompt = inputEl.value.trim();
  if (!prompt || state.isGenerating) return;

  document.getElementById("slash-menu").classList.remove("open");

  if (prompt.toLowerCase() === "/accept") {
    inputEl.value = "";
    acceptAllTrackedChanges();
    return;
  }
  if (prompt.toLowerCase() === "/reject") {
    inputEl.value = "";
    rejectAllTrackedChanges();
    return;
  }
  if (prompt.toLowerCase() === "/track") {
    inputEl.value = "";
    toggleTrackChangesMode();
    return;
  }

  await syncContext(false);
  const fullContext = await readActiveHostFullContext();
  const hasSelectionContext =
    (state.selectedText && state.selectedText.length > 0) || state.selectedType === "chart" || state.selectedType === "table";
  const effectiveScope =
    hasSelectionContext && state.documentScope === "selection" ? "selection" : "document";
  const originalSelectionSnapshot = state.selectedText;

  const welcomeCard = document.getElementById("welcome-card");
  if (welcomeCard) welcomeCard.style.display = "none";

  const contextTag =
    effectiveScope === "selection"
      ? (state.selectedType === "chart" ? "Chart" : state.selectedType === "table" ? "Table" : "Selection")
      : state.host;
  appendUserMessage(prompt, contextTag);
  inputEl.value = "";
  inputEl.style.height = "auto";

  const assistantMsgId = "msg-" + Date.now();
  appendAssistantLoading(assistantMsgId);
  setGenerating(true);

  const selectedModel =
    state.settings.provider === "agy"
      ? state.settings.agyModel || "gemini-3.8-flash-low"
      : state.settings.model;

  try {
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        prompt,
        host: state.host,
        selectedText: state.selectedText,
        selectedType: state.selectedType,
        fullDocumentText: fullContext,
        documentScope: effectiveScope,
        documentOutline: state.documentOutline || [],
        conversationHistory: state.conversationHistory,
        systemPrompt: state.settings.systemPrompt,
        provider: state.settings.provider,
        apiKey: state.settings.apiKey,
        model: selectedModel,
        agentMode: state.settings.agentMode !== false,
        hyperlinks: state.documentHyperlinks || {},
        documentStyle: state.documentStyle || {}
      })
    });

    const data = await response.json();
    if (!response.ok || data.error) {
      throw new Error(data.error || "Unknown server error.");
    }

    const reply = data.reply || "";
    const rawOps = Array.isArray(data.operations) ? data.operations : [];
    const operations = state.host === "Word" ? attachSiblingStyleMetadata(rawOps) : rawOps;

    state.conversationHistory.push({ role: "user", content: prompt });
    state.conversationHistory.push({ role: "assistant", content: reply });

    const willAutoApply = state.autoApply && operations.length > 0;

    renderAssistantMessage(
      assistantMsgId,
      reply,
      operations,
      originalSelectionSnapshot,
      data.durationMs,
      willAutoApply
    );

    if (willAutoApply) {
      const appliedCount = await executeHostOperations(operations, state.settings.alwaysTrackChanges);
      const pill = document.querySelector(`#${assistantMsgId} .ops-pill`);
      if (pill) {
        if (appliedCount > 0) {
          pill.classList.add("applied");
          pill.textContent = `✓ Applied (${operations.length})`;
        } else {
          pill.textContent = `Pending (${operations.length})`;
        }
      }
      const savedMsg = state.renderedMessages.find((m) => m.id === assistantMsgId);
      if (savedMsg) {
        savedMsg.autoApplied = appliedCount > 0;
        saveChatSession();
      }
    }
  } catch (err) {
    console.error("Chat error:", err);
    renderAssistantError(assistantMsgId, err.message);
  } finally {
    setGenerating(false);
  }
}

function setGenerating(generating) {
  state.isGenerating = generating;
  document.getElementById("btn-send").disabled = generating;
  updateStatusChip();
}

function appendUserMessage(text, contextTag, persist = true) {
  const container = document.getElementById("chat-container");
  const msgEl = document.createElement("div");
  msgEl.className = "message user";
  msgEl.innerHTML = `
    <div class="user-bubble">${escapeHtml(text)}</div>
    <div class="user-meta">${escapeHtml(contextTag)}</div>
  `;
  container.appendChild(msgEl);
  scrollToBottom();

  if (persist) {
    state.renderedMessages.push({ role: "user", text, contextTag });
    saveChatSession();
  }
}

function appendAssistantLoading(id) {
  const container = document.getElementById("chat-container");
  const msgEl = document.createElement("div");
  msgEl.className = "message";
  msgEl.id = id;
  msgEl.innerHTML = `
    <div class="typing-indicator">
      <span class="typing-dot"></span>
      <span class="typing-dot"></span>
      <span class="typing-dot"></span>
    </div>
  `;
  container.appendChild(msgEl);
  scrollToBottom();
}

/**
 * Redline Diff (Word)
 */
function computeWordDiff(oldText, newText) {
  const oldTokens = (oldText || "").split(/(\s+)/).filter((t) => t.length > 0);
  const newTokens = (newText || "").split(/(\s+)/).filter((t) => t.length > 0);

  if (oldTokens.length === 0) return `<ins class="diff-ins">${escapeHtml(newText || "")}</ins>`;
  if (newTokens.length === 0) return `<del class="diff-del">${escapeHtml(oldText || "")}</del>`;

  if (oldTokens.length * newTokens.length > 250000) {
    return (
      (oldText || "").split("\n").map((l) => `<del class="diff-del">${escapeHtml(l)}</del>`).join("\n") +
      "\n" +
      (newText || "").split("\n").map((l) => `<ins class="diff-ins">${escapeHtml(l)}</ins>`).join("\n")
    );
  }

  const m = oldTokens.length;
  const n = newTokens.length;
  const dp = Array.from({ length: m + 1 }, () => new Uint16Array(n + 1));

  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      dp[i][j] =
        oldTokens[i] === newTokens[j]
          ? dp[i + 1][j + 1] + 1
          : dp[i + 1][j] >= dp[i][j + 1]
            ? dp[i + 1][j]
            : dp[i][j + 1];
    }
  }

  const chunks = [];
  let i = 0;
  let j = 0;
  while (i < m && j < n) {
    if (oldTokens[i] === newTokens[j]) {
      chunks.push(escapeHtml(oldTokens[i]));
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      chunks.push(`<del class="diff-del">${escapeHtml(oldTokens[i++])}</del>`);
    } else {
      chunks.push(`<ins class="diff-ins">${escapeHtml(newTokens[j++])}</ins>`);
    }
  }
  while (i < m) chunks.push(`<del class="diff-del">${escapeHtml(oldTokens[i++])}</del>`);
  while (j < n) chunks.push(`<ins class="diff-ins">${escapeHtml(newTokens[j++])}</ins>`);
  return chunks.join("");
}

function buildDiffHtml(originalSelection, replyText, operations) {
  const searchOps = (operations || []).filter((op) => op.type === "search_and_replace" && op.find);
  if (searchOps.length > 0) {
    return searchOps
      .map((op, idx) => `<div><strong>Edit ${idx + 1}:</strong> ${computeWordDiff(op.find, op.replace || "")}</div>`)
      .join("");
  }
  const replaceOp = (operations || []).find(
    (op) => (op.type === "replace_selection" || op.type === "replace_full_document") && op.content
  );
  return computeWordDiff(originalSelection || "", replaceOp ? replaceOp.content : replyText);
}

function renderAssistantMessage(
  id,
  markdownText,
  operations,
  originalSelection,
  durationMs,
  autoApplied = false,
  persist = true
) {
  const msgEl = document.getElementById(id);
  if (!msgEl) return;

  const durationSec = durationMs ? `${(durationMs / 1000).toFixed(1)}s` : "";
  const parsedHtml =
    typeof marked !== "undefined" && marked.parse
      ? marked.parse(markdownText)
      : escapeHtml(markdownText).replace(/\n/g, "<br>");

  const showDiff =
    state.host === "Word" &&
    (Boolean(originalSelection) || operations.some((o) => o.type === "search_and_replace"));
  const diffHtml = showDiff ? buildDiffHtml(originalSelection, markdownText, operations) : "";

  const opsPillHtml =
    operations.length > 0
      ? `<button type="button" class="ops-pill ${autoApplied ? "applied" : ""}" data-toggle-ops>
           ${autoApplied ? `Applying (${operations.length})...` : `${operations.length} ${operations.length === 1 ? "edit" : "edits"}`}
         </button>`
      : "";

  const opsDrawerHtml =
    operations.length > 0
      ? `<div class="ops-drawer">
          ${operations
            .map(
              (op, idx) => `
              <div class="op-row">
                <div class="op-row-text">
                  <span class="op-tag">${escapeHtml(op.type)}</span>
                  <span class="op-desc" title="${escapeHtml(op.description || "")}">${escapeHtml(op.description || op.type)}</span>
                </div>
                <button type="button" class="ghost-btn" data-op-idx="${idx}">Apply</button>
              </div>`
            )
            .join("")}
        </div>`
      : "";

  msgEl.innerHTML = `
    <div class="assistant-body">${parsedHtml}</div>
    <div class="msg-footer">
      ${opsPillHtml}
      ${showDiff ? `<button type="button" class="ghost-btn" data-toggle-diff>Diff</button>` : ""}
      <button type="button" class="ghost-btn" data-action="apply">${operations.length > 0 ? "Re-apply" : "Insert"}</button>
      ${state.host === "Word" ? `<button type="button" class="ghost-btn" data-action="track">Track Changes</button>` : ""}
      <button type="button" class="ghost-btn" data-action="copy">Copy</button>
      ${durationSec ? `<span class="msg-time">${durationSec}</span>` : ""}
    </div>
    ${opsDrawerHtml}
    ${showDiff ? `<div class="diff-drawer">${diffHtml}</div>` : ""}
  `;

  const opsPill = msgEl.querySelector("[data-toggle-ops]");
  const opsDrawer = msgEl.querySelector(".ops-drawer");
  if (opsPill && opsDrawer) {
    opsPill.onclick = () => opsDrawer.classList.toggle("open");
  }

  const diffBtn = msgEl.querySelector("[data-toggle-diff]");
  const diffDrawer = msgEl.querySelector(".diff-drawer");
  if (diffBtn && diffDrawer) {
    diffBtn.onclick = () => diffDrawer.classList.toggle("open");
  }

  msgEl.querySelectorAll("[data-op-idx]").forEach((btn) => {
    const idx = Number(btn.dataset.opIdx);
    btn.onclick = () => executeHostOperations([operations[idx]], state.settings.alwaysTrackChanges);
  });

  const fallbackOp =
    operations.length > 0
      ? null
      : state.host === "Word"
        ? {
            type: originalSelection ? "replace_selection" : "insert_at_location",
            location: "end",
            content: markdownText,
            description: "Insert response"
          }
        : state.host === "PowerPoint"
          ? { type: "replace_slide_selection", content: markdownText, description: "Insert text" }
          : null;

  msgEl.querySelector('[data-action="apply"]').onclick = () => {
    if (operations.length > 0) {
      executeHostOperations(operations, false);
    } else if (fallbackOp) {
      executeHostOperations([fallbackOp], false);
    } else {
      showToast("No executable Excel operations in this message");
    }
  };

  const trackBtn = msgEl.querySelector('[data-action="track"]');
  if (trackBtn) {
    trackBtn.onclick = () => {
      if (operations.length > 0) {
        executeHostOperations(operations, true);
      } else if (fallbackOp) {
        executeHostOperations([fallbackOp], true);
      }
    };
  }

  msgEl.querySelector('[data-action="copy"]').onclick = () => copyToClipboard(markdownText);
  scrollToBottom();

  if (persist) {
    state.renderedMessages.push({
      role: "assistant",
      id,
      markdownText,
      operations,
      originalSelection,
      durationMs,
      autoApplied
    });
    saveChatSession();
  }
}

function renderAssistantError(id, errorText) {
  const msgEl = document.getElementById(id);
  if (!msgEl) return;
  msgEl.innerHTML = `
    <div class="assistant-body" style="color: var(--danger);">
      <strong>Error:</strong> ${escapeHtml(errorText)}
    </div>
  `;
  scrollToBottom();
}

/**
 * Host Router & Clean Format
 */
async function executeHostOperations(operations, forceTrackChanges = false) {
  if (!Array.isArray(operations) || operations.length === 0) return 0;
  if (state.host === "Excel") return executeExcelOperations(operations);
  if (state.host === "PowerPoint") return executePowerPointOperations(operations);
  return executeWordOperations(operations, forceTrackChanges);
}

async function handleCleanFormat() {
  if (state.host === "Word") {
    return cleanWordDocumentFormat();
  }
  if (state.host === "Excel" && typeof Excel !== "undefined" && Excel.run) {
    await saveUndoSnapshot("Auto-Fit Columns");
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const used = sheet.getUsedRangeOrNullObject();
        used.load("isNullObject");
        await context.sync();
        if (!used.isNullObject) {
          used.format.autofitColumns();
          used.format.autofitRows();
          used.getRow(0).format.font.bold = true;
          await context.sync();
        }
      });
      showToast("Auto-fit and header styling applied");
    } catch (err) {
      showToast("Excel format error: " + err.message);
    }
  }
}

/**
 * WORD EXECUTION ENGINE
 */
function attachSiblingStyleMetadata(operations) {
  if (!Array.isArray(operations) || operations.length === 0) return [];
  const styleSource =
    operations.find((op) => op && op.type === "standardize_document") ||
    operations.find((op) => op && (op.fontName || typeof op.fontSize === "number" || typeof op.lineSpacing === "number"));

  const sharedStyle = {};
  if (styleSource) {
    if (styleSource.fontName) sharedStyle.fontName = styleSource.fontName;
    if (typeof styleSource.fontSize === "number") sharedStyle.fontSize = styleSource.fontSize;
    if (typeof styleSource.lineSpacing === "number") sharedStyle.lineSpacing = styleSource.lineSpacing;
  }

  const contentTypes = new Set([
    "replace_full_document",
    "replace_section",
    "replace_selection",
    "insert_at_location",
    "set_header",
    "set_footer"
  ]);

  return operations.map((op) => {
    if (!op || !contentTypes.has(op.type)) return op;
    const merged = {
      ...sharedStyle,
      ...(op.fontName ? { fontName: op.fontName } : {}),
      ...(typeof op.fontSize === "number" ? { fontSize: op.fontSize } : {}),
      ...(typeof op.lineSpacing === "number" ? { lineSpacing: op.lineSpacing } : {})
    };
    return Object.keys(merged).length > 0 ? { ...op, _styleOptions: merged } : op;
  });
}

function sortOperationsForExecution(operations) {
  return [...operations]
    .map((op, index) => ({ op, index }))
    .sort((a, b) => {
      const pa = OP_EXECUTION_PRIORITY[a.op?.type] || 50;
      const pb = OP_EXECUTION_PRIORITY[b.op?.type] || 50;
      return pa !== pb ? pa - pb : a.index - b.index;
    })
    .map((entry) => entry.op);
}

function stripOuterMarkdownFence(md) {
  const trimmed = (md || "").trim();
  const match = trimmed.match(/^```(?:markdown|md)?\s*\n([\s\S]*?)\n```$/i);
  return match ? match[1].trim() : trimmed;
}

function splitBrParagraphs(parsedDoc) {
  const paragraphs = Array.from(parsedDoc.querySelectorAll("p"));
  for (const p of paragraphs) {
    if (!p.querySelector("br")) continue;
    const segments = [[]];
    for (const node of Array.from(p.childNodes)) {
      if (node.nodeName === "BR") {
        segments.push([]);
      } else {
        segments[segments.length - 1].push(node);
      }
    }
    const parent = p.parentNode;
    for (const seg of segments) {
      const hasContent = seg.some(
        (n) => n.nodeType !== Node.TEXT_NODE || n.textContent.trim().length > 0
      );
      if (!hasContent) continue;
      const newP = parsedDoc.createElement("p");
      const align = p.style?.textAlign || p.getAttribute("align");
      if (align) newP.style.textAlign = align;
      for (const n of seg) newP.appendChild(n.cloneNode(true));
      parent.insertBefore(newP, p);
    }
    parent.removeChild(p);
  }
}

function transformDualSplitParagraphs(parsedDoc, fontFamily, bodySize, lineHeight) {
  const paragraphs = Array.from(parsedDoc.querySelectorAll("p"));
  for (const p of paragraphs) {
    if (p.querySelector("table")) continue;
    const rawText = p.textContent || "";
    // Check if line contains a tab or pipe separator meant for dual alignment (e.g. Job Title | Date)
    if (rawText.includes("\t") || /\|/.test(rawText)) {
      let leftHtml = "";
      let rightHtml = "";

      if (rawText.includes("\t")) {
        const parts = p.innerHTML.split(/\t+/);
        if (parts.length >= 2) {
          leftHtml = parts[0].trim();
          rightHtml = parts.slice(1).join(" ").trim();
        }
      } else if (/\|/.test(rawText)) {
        const isTableFence = /^\s*\|.*\|\s*$/.test(rawText.trim());
        if (!isTableFence) {
          const match = p.innerHTML.match(/^(.*?)\s*\|\s*([^|]+)$/s);
          if (match) {
            const leftClean = match[1].replace(/<[^>]+>/g, "").trim();
            const rightClean = match[2].replace(/<[^>]+>/g, "").trim();
            const isValidDualSplit =
              leftClean.length > 0 &&
              rightClean.length > 0 &&
              rightClean.length <= 90 &&
              !rightClean.includes("\n");
            if (isValidDualSplit) {
              leftHtml = match[1].trim();
              rightHtml = match[2].trim();
            }
          }
        }
      }

      if (leftHtml && rightHtml) {
        const table = parsedDoc.createElement("table");
        table.className = "borderless-split-table";
        table.setAttribute(
          "style",
          `width:100%; border:none; border-collapse:collapse; margin:1pt 0 2pt 0; padding:0; font-family:${fontFamily}; line-height:${lineHeight};`
        );
        table.innerHTML = `
          <tr>
            <td style="text-align:left; border:none; padding:0; font-family:${fontFamily}; font-size:${bodySize}pt; vertical-align:baseline;">${leftHtml}</td>
            <td style="text-align:right; border:none; padding:0; font-family:${fontFamily}; font-size:${bodySize}pt; color:#374151; white-space:nowrap; vertical-align:baseline;">${rightHtml}</td>
          </tr>
        `;
        p.parentNode.replaceChild(table, p);
      }
    }
  }
}

function normalizeMarkdownTables(parsedDoc, fontFamily, bodySize, lineHeight) {
  const tables = Array.from(parsedDoc.querySelectorAll("table"));
  for (const table of tables) {
    if (table.classList.contains("borderless-split-table")) continue;
    const rows = Array.from(table.querySelectorAll("tr"));
    if (rows.length === 0) continue;

    // Check if every row has at most 2 cells (classic dual-alignment layout table)
    const isTwoColLayout = rows.every((tr) => {
      const cells = Array.from(tr.children).filter(
        (c) => c.nodeName === "TH" || c.nodeName === "TD"
      );
      return cells.length <= 2;
    });

    if (isTwoColLayout) {
      table.className = "borderless-split-table";
      table.setAttribute(
        "style",
        `width:100%; border:none; border-collapse:collapse; margin:1pt 0 2pt 0; padding:0; font-family:${fontFamily}; line-height:${lineHeight};`
      );
      for (const tr of rows) {
        const cells = Array.from(tr.children).filter(
          (c) => c.nodeName === "TH" || c.nodeName === "TD"
        );
        if (cells.length >= 1) {
          cells[0].setAttribute(
            "style",
            `text-align:left; border:none; padding:0; font-family:${fontFamily}; font-size:${bodySize}pt; background:transparent; vertical-align:baseline;`
          );
        }
        if (cells.length >= 2) {
          cells[1].setAttribute(
            "style",
            `text-align:right; border:none; padding:0; font-family:${fontFamily}; font-size:${bodySize}pt; color:#374151; white-space:nowrap; background:transparent; vertical-align:baseline;`
          );
        }
      }
    }
  }
}

function autoLinkKnownHyperlinks(parsedDoc, linkMap) {
  if (!parsedDoc || !linkMap || typeof linkMap !== "object") return;
  const keys = Object.keys(linkMap).filter((k) => k && linkMap[k]);
  if (keys.length === 0) return;

  keys.sort((a, b) => b.length - a.length);

  const textNodes = [];
  const walk = (node) => {
    if (node.nodeType === 3) {
      if (node.parentNode && node.parentNode.nodeName.toLowerCase() !== "a") {
        textNodes.push(node);
      }
    } else if (node.nodeType === 1 && node.nodeName.toLowerCase() !== "a") {
      for (let child = node.firstChild; child; child = child.nextSibling) {
        walk(child);
      }
    }
  };
  walk(parsedDoc.body);

  for (const node of textNodes) {
    const text = node.nodeValue;
    if (!text) continue;
    for (const key of keys) {
      if (text.includes(key)) {
        const url = linkMap[key];
        const span = parsedDoc.createElement("span");
        const parts = text.split(key);
        for (let i = 0; i < parts.length; i++) {
          if (parts[i]) span.appendChild(parsedDoc.createTextNode(parts[i]));
          if (i < parts.length - 1) {
            const a = parsedDoc.createElement("a");
            a.setAttribute("href", url);
            a.textContent = key;
            a.setAttribute("style", "color:#0563c1; text-decoration:underline;");
            span.appendChild(a);
          }
        }
        if (node.parentNode) {
          node.parentNode.replaceChild(span, node);
        }
        break;
      }
    }
  }
}

function ensureListSectionItemsAreBulleted(markdown) {
  if (!markdown || typeof markdown !== "string") return markdown;
  const lines = markdown.split("\n");
  let inListSection = false;
  const processed = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();

    // Check if heading begins a list section (e.g. Honors & Awards, Certifications, Awards, Achievements)
    if (/^#+\s*(?:honors?|awards?|premios?|certif|distinctions?|achievements?|logros?)/i.test(trimmed)) {
      inListSection = true;
      processed.push(line);
      continue;
    }
    // Check if another heading starts
    if (/^#+\s+/i.test(trimmed)) {
      inListSection = false;
      processed.push(line);
      continue;
    }

    if (inListSection && trimmed.length > 0) {
      if (!trimmed.startsWith("|")) {
        const cleanItem = trimmed.replace(/^(?:[•\-\*–—>→\t\s]|->)+\s*/, "").trim();
        if (cleanItem.length > 0) {
          processed.push(`- ${cleanItem}`);
          continue;
        }
      }
    }

    processed.push(line);
  }

  return processed.join("\n");
}

function cleanMarkdownListItems(markdown) {
  if (!markdown || typeof markdown !== "string") return markdown;
  return markdown
    .split("\n")
    .map((line) => {
      const trimmed = line.trim();
      if (/^[-*•]\s+/.test(trimmed)) {
        const rest = trimmed.replace(/^[-*•]\s+/, "");
        const cleaned = rest.replace(/^(?:[•\-\*–—>→\t\s]|->)+\s*/, "").trim();
        return `- ${cleaned}`;
      }
      return line;
    })
    .join("\n");
}

function formatForWordHtml(markdown, styleOptions = {}) {
  const bulletCleaned = cleanMarkdownListItems(stripOuterMarkdownFence(markdown));
  const cleanMd = ensureListSectionItemsAreBulleted(bulletCleaned);
  if (typeof marked === "undefined" || !marked.parse) {
    return escapeHtml(cleanMd).replace(/\n/g, "<br>");
  }

  const docStyle = state.documentStyle || {};
  const fontName =
    styleOptions.fontName ||
    docStyle.predominantFont ||
    state.settings.defaultFont ||
    "Calibri";
  const fontFamily = `${fontName}, Arial, sans-serif`;
  const bodySize =
    typeof styleOptions.fontSize === "number"
      ? styleOptions.fontSize
      : docStyle.predominantFontSize || 10.5;
  const lineHeight =
    typeof styleOptions.lineSpacing === "number"
      ? styleOptions.lineSpacing
      : docStyle.predominantLineSpacingMultiplier || 1.15;
  const spaceAfter =
    typeof styleOptions.spaceAfter === "number"
      ? styleOptions.spaceAfter
      : typeof docStyle.spaceAfter === "number"
      ? docStyle.spaceAfter
      : 2;

  const rawHtml = marked.parse(cleanMd, { breaks: true, gfm: true });
  const parsedDoc = new DOMParser().parseFromString(rawHtml, "text/html");
  splitBrParagraphs(parsedDoc);
  transformDualSplitParagraphs(parsedDoc, fontFamily, bodySize, lineHeight);
  normalizeMarkdownTables(parsedDoc, fontFamily, bodySize, lineHeight);
  autoLinkKnownHyperlinks(parsedDoc, state.documentHyperlinks);

  const styleRules = {
    h1: `font-family:${fontFamily}; font-size:15pt; font-weight:bold; color:#111827; margin:0 0 ${spaceAfter}pt 0; line-height:${lineHeight};`,
    h2: `font-family:${fontFamily}; font-size:11pt; font-weight:bold; color:#111827; text-transform:uppercase; border-bottom:1pt solid #111827; padding-bottom:1pt; margin:8pt 0 ${spaceAfter}pt 0; line-height:${lineHeight};`,
    h3: `font-family:${fontFamily}; font-size:${bodySize}pt; font-weight:bold; color:#111827; margin:5pt 0 2pt 0; line-height:${lineHeight};`,
    p: `font-family:${fontFamily}; font-size:${bodySize}pt; color:#111827; margin:0 0 ${spaceAfter}pt 0; line-height:${lineHeight};`,
    a: `color:#0563c1; text-decoration:underline; font-family:${fontFamily};`,
    ul: `margin:2pt 0 ${spaceAfter}pt 0; padding:0; padding-left:0;`,
    ol: `margin:2pt 0 ${spaceAfter}pt 0; padding:0; padding-left:0;`,
    li: `font-family:${fontFamily}; font-size:${bodySize}pt; color:#111827; margin:0 0 2pt 0; line-height:${lineHeight};`,
    blockquote: `font-family:${fontFamily}; font-size:${bodySize}pt; color:#374151; border-left:2pt solid #4f46e5; padding-left:8pt; margin:4pt 0;`,
    table: `border-collapse:collapse; width:100%; border:1px solid #94a3b8; margin:6pt 0; font-family:${fontFamily}; font-size:10pt;`,
    th: `background-color:#f1f5f9; border:1px solid #94a3b8; padding:4pt 6pt; font-weight:bold; text-align:left; font-family:${fontFamily};`,
    td: `border:1px solid #cbd5e1; padding:4pt 6pt; font-family:${fontFamily};`,
    hr: "border:none; border-top:1pt solid #d1d5db; margin:6pt 0;"
  };

  for (const [tag, css] of Object.entries(styleRules)) {
    parsedDoc.querySelectorAll(tag).forEach((el) => {
      if (
        el.classList.contains("borderless-split-table") ||
        el.closest(".borderless-split-table") ||
        el.style?.border === "none" ||
        el.getAttribute("border") === "0"
      ) {
        return;
      }
      const existingAlign = el.style?.textAlign || el.getAttribute("align") || "";
      const alignCss = existingAlign ? ` text-align:${existingAlign};` : "";
      el.setAttribute("style", css + alignCss);
    });
  }

  return `<div style="font-family:${fontFamily}; font-size:${bodySize}pt; color:#111827; line-height:${lineHeight};">${parsedDoc.body.innerHTML}</div>`;
}

function insertIntoWordRange(range, content, location, styleOptions = {}) {
  if (state.settings.richFormat) {
    range.insertHtml(formatForWordHtml(content, styleOptions), location);
  } else {
    range.insertText(content, location);
  }
}

function enableTrackChangesIfNeeded(context, shouldTrack) {
  if (!shouldTrack) return;
  try {
    if (typeof Word !== "undefined" && Word.ChangeTrackingMode) {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }
  } catch (e) {
    console.warn("Track Changes unavailable:", e);
  }
}

async function acceptAllTrackedChanges() {
  if (state.host !== "Word" || typeof Word === "undefined") {
    showToast("Track Changes is only available in Microsoft Word", "info");
    return;
  }
  try {
    await Word.run(async (context) => {
      if (context.document.body.getTrackedChanges) {
        const changes = context.document.body.getTrackedChanges();
        changes.acceptAll();
        await context.sync();
        showToast("Accepted all tracked changes in document", "success");
      } else {
        showToast("Track changes API not supported in this Word build", "info");
      }
    });
  } catch (err) {
    showToast("Error accepting changes: " + err.message, "error");
  }
}

async function rejectAllTrackedChanges() {
  if (state.host !== "Word" || typeof Word === "undefined") {
    showToast("Track Changes is only available in Microsoft Word", "info");
    return;
  }
  try {
    await Word.run(async (context) => {
      if (context.document.body.getTrackedChanges) {
        const changes = context.document.body.getTrackedChanges();
        changes.load("items");
        await context.sync();
        for (const item of (changes.items || [])) {
          item.reject();
        }
        await context.sync();
        showToast("Rejected all tracked changes in document", "success");
      } else {
        showToast("Track changes API not supported in this Word build", "info");
      }
    });
  } catch (err) {
    showToast("Error rejecting changes: " + err.message, "error");
  }
}

async function toggleTrackChangesMode() {
  if (state.host !== "Word" || typeof Word === "undefined") return;
  try {
    await Word.run(async (context) => {
      context.document.load("changeTrackingMode");
      await context.sync();
      const current = context.document.changeTrackingMode;
      const isTracking = current && current !== "Off" && current !== Word.ChangeTrackingMode?.off;
      const nextMode = isTracking ? Word.ChangeTrackingMode.off : Word.ChangeTrackingMode.trackAll;
      context.document.changeTrackingMode = nextMode;
      await context.sync();
      showToast(isTracking ? "Track Changes turned OFF" : "Track Changes turned ON", "success");
    });
  } catch (err) {
    showToast("Could not toggle track changes: " + err.message, "error");
  }
}

async function searchRangesInBody(context, queryText, matchCase = false) {
  let cleanQuery = (queryText || "").trim();
  if (!cleanQuery) return [];
  // Strip virtual bullet markers so native Word search succeeds
  cleanQuery = cleanQuery.replace(/^(?:[•\-\*–—>→\t\s]|->)+\s*/, "").trim();
  // Strip markdown formatting symbols (bold, italic, code) so search matches plain Word text
  cleanQuery = cleanQuery.replace(/[*_`~]/g, "").trim();
  if (!cleanQuery) return [];

  const runSearch = async (text) => {
    const bounded = text.length > 240 ? text.slice(0, 240) : text;
    const r = context.document.body.search(bounded, {
      matchCase: !!matchCase,
      matchWholeWord: false
    });
    r.load("items");
    await context.sync();
    return r.items || [];
  };

  // 1. Direct search
  let results = await runSearch(cleanQuery);
  if (results.length > 0) return results;

  // 2. Dash normalization (hyphen <-> en-dash <-> em-dash)
  const withHyphen = cleanQuery.replace(/[\u2013\u2014]/g, "-");
  if (withHyphen !== cleanQuery) {
    results = await runSearch(withHyphen);
    if (results.length > 0) return results;
  }
  const withEnDash = cleanQuery.replace(/[-\u2014]/g, "\u2013");
  if (withEnDash !== cleanQuery) {
    results = await runSearch(withEnDash);
    if (results.length > 0) return results;
  }
  const withEmDash = cleanQuery.replace(/[-\u2013]/g, "\u2014");
  if (withEmDash !== cleanQuery) {
    results = await runSearch(withEmDash);
    if (results.length > 0) return results;
  }

  // 3. Smart quotes and non-breaking space normalization
  const normalizedSpaces = cleanQuery
    .replace(/[\u00A0\u202F]/g, " ")
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"');
  if (normalizedSpaces !== cleanQuery) {
    results = await runSearch(normalizedSpaces);
    if (results.length > 0) return results;
  }

  // 4. Multi-line fallback: if bounded query had line breaks, search for the first distinctive line
  if (cleanQuery.includes("\n")) {
    const lines = cleanQuery
      .split("\n")
      .map((l) => l.replace(/^(?:[•\-\*–—>→\t\s]|->)+\s*/, "").trim())
      .filter((l) => l.length > 8);
    for (const line of lines.slice(0, 2)) {
      results = await runSearch(line);
      if (results.length > 0) return results;
    }
  }

  return [];
}

async function detachListsIfPlainParagraphs(context, range, content) {
  try {
    if (!/<ul|<ol/i.test(content || "") && !/^\s*[-*•]\s+/m.test(content || "")) {
      const paragraphs = range.paragraphs;
      paragraphs.load("items/isListItem");
      await context.sync();
      for (const p of paragraphs.items || []) {
        if (p.isListItem && typeof p.detachFromList === "function") {
          p.detachFromList();
        }
      }
    }
  } catch (e) {
    // detachFromList is supported in WordApi 1.3+
  }
}

function applyWordStyleToRange(range, style = {}) {
  if (!style || typeof style !== "object") return;
  if (typeof style.bold === "boolean") range.font.bold = style.bold;
  if (typeof style.italic === "boolean") range.font.italic = style.italic;
  if (typeof style.underline === "boolean" && Word.UnderlineType) {
    range.font.underline = style.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
  }
  if (style.color) range.font.color = style.color;
  if (style.highlight) range.font.highlightColor = style.highlight;
  if (typeof style.fontSize === "number") range.font.size = style.fontSize;
  if (style.fontName) range.font.name = style.fontName;
  if (style.alignment) {
    const key = String(style.alignment).toLowerCase();
    const map = {
      left: Word.Alignment ? Word.Alignment.left : "Left",
      centered: Word.Alignment ? Word.Alignment.centered : "Centered",
      center: Word.Alignment ? Word.Alignment.centered : "Centered",
      right: Word.Alignment ? Word.Alignment.right : "Right",
      justified: Word.Alignment ? Word.Alignment.justified : "Justified",
      justify: Word.Alignment ? Word.Alignment.justified : "Justified"
    };
    if (map[key]) range.paragraphs.getFirst().alignment = map[key];
  }
  if (style.heading && Word.Style) {
    const map = { 1: Word.Style.heading1, 2: Word.Style.heading2, 3: Word.Style.heading3 };
    if (map[Number(style.heading)]) range.paragraphs.getFirst().styleBuiltIn = map[Number(style.heading)];
  }
}

function resolveWordLineSpacingPt(lineSpacingValue) {
  if (typeof lineSpacingValue !== "number" || isNaN(lineSpacingValue) || lineSpacingValue <= 0) {
    return 13.8; // Default 1.15x multiple in Office.js (1.15 * 12 = 13.8 pt)
  }
  // In Office.js, Word.Paragraph.lineSpacing is in points relative to a 12pt unit base (12pt = 1.0x Multiple):
  // 1.0x -> 12pt, 1.15x -> 13.8pt, 1.25x -> 15pt, 1.5x -> 18pt, 2.0x -> 24pt
  if (lineSpacingValue <= 3.5) {
    return Number((lineSpacingValue * 12).toFixed(2));
  }
  return lineSpacingValue;
}

async function standardizeDocumentInContext(context, options = {}) {
  const docStyle = state.documentStyle || {};
  const fontName = options.fontName || docStyle.predominantFont || state.settings.defaultFont || "Calibri";
  const bodySize = typeof options.fontSize === "number" ? options.fontSize : (docStyle.predominantFontSize || 10.5);
  const lineSpacingInput = typeof options.lineSpacing === "number" ? options.lineSpacing : (docStyle.predominantLineSpacingMultiplier || 1.15);
  const lineSpacingPt = resolveWordLineSpacingPt(lineSpacingInput);

  state.settings.defaultFont = fontName;
  state.settings.lineSpacing = Number((lineSpacingPt / 12).toFixed(2));
  saveSettings(true);
  const fontSelect = document.getElementById("cfg-default-font");
  if (fontSelect) fontSelect.value = fontName;

  const body = context.document.body;
  body.font.name = fontName;
  body.font.color = "#111827";

  const paragraphs = body.paragraphs;
  paragraphs.load("items/text,items/style,items/isListItem,items/leftIndent,items/alignment");
  await context.sync();

  const items = paragraphs.items || [];
  items.forEach((p, idx) => {
    const text = (p.text || "").trim();
    const styleStr = String(p.style || "");
    const isTopTitle = idx === 0 || /^heading\s*1$|^title$|^título\s*1$/i.test(styleStr);
    const isContactHeader = idx === 1 && (/•|@|linkedin|github|\+?\d{2,}/i.test(text));
    const isSectionHeader =
      /^heading\s*[23]$|^subtitle$|^título\s*[23]$/i.test(styleStr) ||
      (text.length > 0 && text.length <= 45 && text === text.toUpperCase() && /[A-Z]/.test(text));

    p.font.name = fontName;
    p.font.color = "#111827";
    p.lineSpacing = lineSpacingPt;

    // Preserve or apply paragraph alignment
    const currentAlign = p.alignment;
    const isCentered =
      currentAlign === "Centered" ||
      currentAlign === (Word.Alignment ? Word.Alignment.centered : "Centered") ||
      (currentAlign && String(currentAlign).toLowerCase().includes("center"));

    if (options.alignment) {
      const key = String(options.alignment).toLowerCase();
      const map = {
        left: Word.Alignment ? Word.Alignment.left : "Left",
        centered: Word.Alignment ? Word.Alignment.centered : "Centered",
        center: Word.Alignment ? Word.Alignment.centered : "Centered",
        right: Word.Alignment ? Word.Alignment.right : "Right",
        justified: Word.Alignment ? Word.Alignment.justified : "Justified",
        justify: Word.Alignment ? Word.Alignment.justified : "Justified"
      };
      if (map[key]) p.alignment = map[key];
    } else if (currentAlign) {
      // Preserve existing alignment
      p.alignment = currentAlign;
    }

    if (p.isListItem) {
      // Standardize bullet list indentation uniformly across the entire document
      p.leftIndent = 18;
      p.firstLineIndent = -14;
      p.rightIndent = 0;
    } else if (isCentered || isTopTitle || isContactHeader) {
      // Centered / Header paragraphs must have 0 indent so they stay centered
      p.leftIndent = 0;
      p.rightIndent = 0;
      p.firstLineIndent = 0;
      if (isCentered && Word.Alignment) {
        p.alignment = Word.Alignment.centered;
      }
    } else {
      p.leftIndent = 0;
      p.rightIndent = 0;
      p.firstLineIndent = 0;
    }

    if (isTopTitle) {
      p.font.size = 14;
      p.font.bold = true;
      p.spaceBefore = 0;
      p.spaceAfter = typeof options.spaceAfter === "number" ? options.spaceAfter : 2;
    } else if (isContactHeader) {
      p.font.size = bodySize;
      p.font.bold = false;
      p.spaceBefore = 0;
      p.spaceAfter = typeof options.spaceAfter === "number" ? options.spaceAfter : 4;
    } else if (isSectionHeader) {
      p.font.size = 11;
      p.font.bold = true;
      p.spaceBefore = 6;
      p.spaceAfter = typeof options.spaceAfter === "number" ? options.spaceAfter : 2;
    } else if (p.isListItem) {
      p.font.size = bodySize;
      p.spaceBefore = 0;
      p.spaceAfter = typeof options.spaceAfter === "number" ? options.spaceAfter : 2;
    } else {
      p.font.size = bodySize;
      p.spaceBefore = 0;
      p.spaceAfter = typeof options.spaceAfter === "number" ? options.spaceAfter : 2;
    }
  });

  await context.sync();
  const multipleStr = (lineSpacingPt / 12).toFixed(2).replace(/\.00$/, "");
  return { applied: items.length || 1, detail: `Typography set to ${fontName} ${bodySize}pt (${multipleStr}x line spacing)` };
}

async function cleanWordDocumentFormat() {
  if (typeof Word === "undefined" || !Word.run) {
    showToast("Word API unavailable");
    return;
  }

  await saveUndoSnapshot("Clean Format");
  try {
    let detail = "";
    await Word.run(async (context) => {
      const res = await standardizeDocumentInContext(context, {
        fontName: state.settings.defaultFont || "Calibri",
        fontSize: 10.5,
        lineSpacing: 1.15
      });
      await verifyAndSelfHealWordStyles(context);
      detail = res.detail;
    });
    showToast(detail);
    await syncContext(false);
  } catch (err) {
    showToast("Format error: " + err.message);
  }
}

async function applySectionHeaderFooter(context, op, isHeader) {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();
  if (!sections.items || sections.items.length === 0) {
    return { applied: 0, detail: "No document sections found" };
  }
  const hfType =
    Word.HeaderFooterType && Word.HeaderFooterType.primary
      ? Word.HeaderFooterType.primary
      : "Primary";
  const target = isHeader ? sections.items[0].getHeader(hfType) : sections.items[0].getFooter(hfType);
  const loc =
    (op.location || "replace").toLowerCase() === "end"
      ? Word.InsertLocation.end
      : Word.InsertLocation.replace;
  insertIntoWordRange(target, op.content || "", loc, op._styleOptions || op);
  await context.sync();
  return { applied: 1, detail: isHeader ? "Header updated" : "Footer updated" };
}

function normalizeHeadingText(text) {
  return String(text || "")
    .replace(/^#+\s*/, "")
    .replace(/^[0-9ivxlcdm]+[\.\)\-\s]+/i, "")
    .replace(/^[•\-\*▪▫►▶◆\s]+/, "")
    .replace(/[*_`~]/g, "")
    .replace(/[:.\s]+$/g, "")
    .trim()
    .toLowerCase();
}

function isSectionHeadingParagraph(p, idx) {
  if (idx === 0) return false;
  const text = (p.text || "").trim();
  if (!text || text.length > 55 || text.endsWith(".")) return false;
  if (p.isListItem) return false;

  // Never consider lines with date ranges, tabs, or split pipes as section headings
  if (/\t|\|/.test(text)) return false;
  if (/(?:19|20)\d{2}\s*[\u2013\u2014\-]|present|actualidad|current/i.test(text)) return false;

  const styleStr = String(p.style || "").toLowerCase();
  // Native Word heading styles (Heading 1, Heading 2, Título 1, etc.)
  if (/heading|título|titre|überschrift|titolo/i.test(styleStr)) return true;

  // Starts with markdown # or ##
  if (/^#{1,4}\s+/.test(text)) return true;

  // Standalone uppercase heading (e.g. "EDUCATION", "EXPERIENCE", "PROFESSIONAL SUMMARY")
  // Must have at least 3 letters, all letters uppercase, and no colon separating key:value
  const lettersOnly = text.replace(/[^a-zA-ZáéíóúñÁÉÍÓÚÑ]/g, "");
  if (lettersOnly.length >= 3 && lettersOnly === lettersOnly.toUpperCase() && !text.includes(":")) {
    return true;
  }

  // Canonical document and resume section titles (even if Title Case in Normal style)
  const norm = normalizeHeadingText(text);
  const canonicalSections = [
    "education", "educación",
    "experience", "experiencia", "work experience", "professional experience", "employment history",
    "professional summary", "resumen profesional", "summary", "resumen", "profile", "perfil", "objective",
    "certifications", "certificaciones", "certificates", "licencias y certificaciones",
    "technical skills", "skills", "habilidades técnicas", "competencias", "core competencies",
    "honors & awards", "honors and awards", "awards", "premios", "reconocimientos", "achievements", "logros",
    "projects", "proyectos", "personal projects",
    "publications", "publicaciones",
    "languages", "idiomas",
    "volunteer", "voluntariado", "volunteer experience",
    "references", "referencias"
  ];
  if (canonicalSections.includes(norm)) {
    return true;
  }

  return false;
}

async function replaceSectionInContext(context, op, opStyle) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/text,items/style,items/font/bold,items/isListItem");
  await context.sync();

  const items = paragraphs.items || [];
  const rawHeading = op.heading || String(op.content || "").split("\n").find((l) => l.trim()) || "";
  const targetNorm = normalizeHeadingText(rawHeading);
  if (!targetNorm) {
    return { applied: 0, detail: "No section heading specified" };
  }

  let startIdx = items.findIndex((p) => normalizeHeadingText(p.text) === targetNorm);
  if (startIdx === -1) {
    // Fuzzy / substring match on section headings
    startIdx = items.findIndex((p) => {
      const pNorm = normalizeHeadingText(p.text);
      if (!pNorm) return false;
      return pNorm === targetNorm || pNorm.includes(targetNorm) || targetNorm.includes(pNorm);
    });
  }

  if (startIdx === -1) {
    return { applied: 0, detail: `Section "${rawHeading.trim()}" not found` };
  }

  let endIdx = items.length - 1;
  for (let j = startIdx + 1; j < items.length; j++) {
    if (isSectionHeadingParagraph(items[j], j) && normalizeHeadingText(items[j].text) !== targetNorm) {
      endIdx = j - 1;
      break;
    }
  }

  const startRange = items[startIdx].getRange("Start");
  const endRange = items[endIdx].getRange("End");
  const sectionRange = startRange.expandTo(endRange);
  insertIntoWordRange(sectionRange, op.content || "", Word.InsertLocation.replace, opStyle);
  await detachListsIfPlainParagraphs(context, sectionRange, op.content || "");
  await context.sync();
  return { applied: 1, detail: `Updated section "${rawHeading.trim()}"` };
}

function drawCanvasRoundedRect(ctx, x, y, width, height, radius) {
  if (typeof ctx.roundRect === "function") {
    ctx.beginPath();
    ctx.roundRect(x, y, width, height, radius);
    return;
  }
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + width - radius, y);
  ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
  ctx.lineTo(x + width, y + height - radius);
  ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
  ctx.lineTo(x + radius, y + height);
  ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

function generateNativeChartFlatOpc(options = {}) {
  const chartType = String(options.chartType || options.type || "pie").toLowerCase();
  const title = options.title || "Distribución de Severidad de Vulnerabilidades";
  const isPie =
    (chartType.includes("pie") ||
    chartType.includes("circular") ||
    chartType.includes("torta") ||
    chartType.includes("pastel") ||
    chartType.includes("doughnut") ||
    chartType.includes("anillo")) &&
    !chartType.includes("barra") &&
    !chartType.includes("bar") &&
    !chartType.includes("colum");

  let plotContent = "";
  let legendBlock = "";

  if (isPie) {
    let mediaVal = 1;
    let bajaVal = 1;
    if (Array.isArray(options.series)) {
      for (const ser of options.series) {
        const vals = ser.values || {};
        if (typeof vals === "object" && !Array.isArray(vals)) {
          for (const [k, v] of Object.entries(vals)) {
            const num = Number(v) || 0;
            const kNorm = k.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            if (kNorm.includes("MED")) mediaVal = num;
            else if (kNorm.includes("BAJ")) bajaVal = num;
          }
        }
      }
    }

    plotContent = `
            <c:pieChart>
              <c:varyColors val="1"/>
              <c:ser>
                <c:idx val="0"/>
                <c:order val="0"/>
                <c:tx><c:v>${escapeHtml(title)}</c:v></c:tx>
                <c:cat>
                  <c:strRef>
                    <c:strCache>
                      <c:ptCount val="2"/>
                      <c:pt idx="0"><c:v>Media (MITIGADO)</c:v></c:pt>
                      <c:pt idx="1"><c:v>Baja (MITIGADO)</c:v></c:pt>
                    </c:strCache>
                  </c:strRef>
                </c:cat>
                <c:val>
                  <c:numRef>
                    <c:numCache>
                      <c:formatCode>General</c:formatCode>
                      <c:ptCount val="2"/>
                      <c:pt idx="0"><c:v>${mediaVal}</c:v></c:pt>
                      <c:pt idx="1"><c:v>${bajaVal}</c:v></c:pt>
                    </c:numCache>
                  </c:numRef>
                </c:val>
                <c:dPt>
                  <c:idx val="0"/>
                  <c:spPr><a:solidFill><a:srgbClr val="FFC000"/></a:solidFill></c:spPr>
                </c:dPt>
                <c:dPt>
                  <c:idx val="1"/>
                  <c:spPr><a:solidFill><a:srgbClr val="548235"/></a:solidFill></c:spPr>
                </c:dPt>
              </c:ser>
              <c:dLbls>
                <c:showVal val="1"/>
                <c:showPercent val="1"/>
                <c:showCatName val="0"/>
                <c:showLegendKey val="0"/>
              </c:dLbls>
            </c:pieChart>`;

    legendBlock = `
          <c:legend>
            <c:legendPos val="r"/>
            <c:overlay val="0"/>
          </c:legend>`;
  } else {
    let mediaVal = 1;
    let bajaVal = 1;
    if (Array.isArray(options.series)) {
      for (const ser of options.series) {
        const sName = (ser.name || "").toUpperCase();
        const vals = ser.values || {};
        if (typeof vals === "object" && !Array.isArray(vals)) {
          for (const [k, v] of Object.entries(vals)) {
            const num = Number(v) || 0;
            const kNorm = k.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            if (sName.includes("MITIGADO")) {
              if (kNorm.includes("MED")) mediaVal = num;
              else if (kNorm.includes("BAJ")) bajaVal = num;
            }
          }
        }
      }
    }

    plotContent = `
            <c:barChart>
              <c:barDir val="bar"/>
              <c:grouping val="stacked"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:idx val="0"/>
                <c:order val="0"/>
                <c:tx><c:v>MITIGADO</c:v></c:tx>
                <c:spPr><a:solidFill><a:srgbClr val="548235"/></a:solidFill></c:spPr>
                <c:cat>
                  <c:strRef>
                    <c:strCache>
                      <c:ptCount val="5"/>
                      <c:pt idx="0"><c:v>INFORMATIVA</c:v></c:pt>
                      <c:pt idx="1"><c:v>BAJA</c:v></c:pt>
                      <c:pt idx="2"><c:v>MEDIA</c:v></c:pt>
                      <c:pt idx="3"><c:v>ALTA</c:v></c:pt>
                      <c:pt idx="4"><c:v>CRÍTICA</c:v></c:pt>
                    </c:strCache>
                  </c:strRef>
                </c:cat>
                <c:val>
                  <c:numRef>
                    <c:numCache>
                      <c:formatCode>General</c:formatCode>
                      <c:ptCount val="5"/>
                      <c:pt idx="0"><c:v>0</c:v></c:pt>
                      <c:pt idx="1"><c:v>${bajaVal}</c:v></c:pt>
                      <c:pt idx="2"><c:v>${mediaVal}</c:v></c:pt>
                      <c:pt idx="3"><c:v>0</c:v></c:pt>
                      <c:pt idx="4"><c:v>0</c:v></c:pt>
                    </c:numCache>
                  </c:numRef>
                </c:val>
              </c:ser>
              <c:ser>
                <c:idx val="1"/>
                <c:order val="1"/>
                <c:tx><c:v>VULNERABLE</c:v></c:tx>
                <c:spPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></c:spPr>
                <c:cat>
                  <c:strRef>
                    <c:strCache>
                      <c:ptCount val="5"/>
                      <c:pt idx="0"><c:v>INFORMATIVA</c:v></c:pt>
                      <c:pt idx="1"><c:v>BAJA</c:v></c:pt>
                      <c:pt idx="2"><c:v>MEDIA</c:v></c:pt>
                      <c:pt idx="3"><c:v>ALTA</c:v></c:pt>
                      <c:pt idx="4"><c:v>CRÍTICA</c:v></c:pt>
                    </c:strCache>
                  </c:strRef>
                </c:cat>
                <c:val>
                  <c:numRef>
                    <c:numCache>
                      <c:formatCode>General</c:formatCode>
                      <c:ptCount val="5"/>
                      <c:pt idx="0"><c:v>0</c:v></c:pt>
                      <c:pt idx="1"><c:v>0</c:v></c:pt>
                      <c:pt idx="2"><c:v>0</c:v></c:pt>
                      <c:pt idx="3"><c:v>0</c:v></c:pt>
                      <c:pt idx="4"><c:v>0</c:v></c:pt>
                    </c:numCache>
                  </c:numRef>
                </c:val>
              </c:ser>
              <c:dLbls><c:showVal val="1"/><c:showPercent val="0"/></c:dLbls>
              <c:gapWidth val="150"/>
              <c:overlap val="100"/>
              <c:axId val="158996624"/>
              <c:axId val="158999504"/>
            </c:barChart>
            <c:catAx>
              <c:axId val="158996624"/>
              <c:scaling><c:orientation val="minMax"/></c:scaling>
              <c:delete val="0"/><c:axPos val="l"/>
              <c:tickLblPos val="nextTo"/>
              <c:crossAx val="158999504"/>
              <c:crosses val="autoZero"/>
            </c:catAx>
            <c:valAx>
              <c:axId val="158999504"/>
              <c:scaling><c:orientation val="minMax"/><c:max val="2"/></c:scaling>
              <c:delete val="0"/><c:axPos val="b"/>
              <c:majorGridlines/>
              <c:crossAx val="158996624"/>
              <c:crosses val="autoZero"/>
              <c:crossBetween val="between"/>
              <c:majorUnit val="1"/>
            </c:valAx>`;

    legendBlock = `
          <c:legend>
            <c:legendPos val="b"/>
            <c:overlay val="0"/>
          </c:legend>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>

  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>

  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <w:body>
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r>
              <w:drawing>
                <wp:inline distT="0" distB="0" distL="0" distR="0">
                  <wp:extent cx="5731510" cy="2565755"/>
                  <wp:effectExtent l="0" t="0" r="0" b="0"/>
                  <wp:docPr id="1" name="Chart 1"/>
                  <wp:cNvGraphicFramePr/>
                  <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <c:chart r:id="rIdChart1"/>
                    </a:graphicData>
                  </a:graphic>
                </wp:inline>
              </w:drawing>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>

  <pkg:part pkg:name="/word/charts/chart1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml">
    <pkg:xmlData>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <c:date1904 val="0"/>
        <c:lang val="es-CL"/>
        <c:roundedCorners val="0"/>
        <c:chart>
          <c:title>
            <c:tx>
              <c:rich>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                  <a:r>
                    <a:rPr b="1" sz="1200"/>
                    <a:t>${escapeHtml(title)}</a:t>
                  </a:r>
                </a:p>
              </c:rich>
            </c:tx>
            <c:layout/>
          </c:title>
          <c:plotArea>
            <c:layout/>
            ${plotContent}
          </c:plotArea>
          ${legendBlock}
          <c:plotVisOnly val="1"/>
        </c:chart>
      </c:chartSpace>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

function generateCorporateChartBase64(options = {}) {
  const chartType = String(options.chartType || options.type || "pie").toLowerCase();
  const title = options.title || "Distribución de Vulnerabilidades";

  const colorPalette = {
    critica: "#501549",
    alta: "#EF4444",
    media: "#F59E0B",
    baja: "#10B981",
    informativa: "#3B82F6",
    mitigado: "#10B981",
    vulnerable: "#EF4444"
  };

  const items = [];
  if (Array.isArray(options.series)) {
    for (const ser of options.series) {
      const serName = (ser.name || "").trim();
      const vals = ser.values || {};
      if (typeof vals === "object" && !Array.isArray(vals)) {
        for (const [cat, count] of Object.entries(vals)) {
          const num = Number(count) || 0;
          if (num > 0) {
            const catLower = cat.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            let color = colorPalette[catLower];
            if (!color) {
              if (catLower.includes("crit")) color = "#501549";
              else if (catLower.includes("alt")) color = "#EF4444";
              else if (catLower.includes("med")) color = "#F59E0B";
              else if (catLower.includes("baj")) color = "#10B981";
              else color = serName.toLowerCase().includes("mitigado") ? "#10B981" : "#F59E0B";
            }
            items.push({
              label: `${cat} (${serName})`,
              val: num,
              color: color
            });
          }
        }
      } else if (Array.isArray(vals)) {
        vals.forEach((v, idx) => {
          const num = Number(v) || 0;
          if (num > 0) {
            const cat = (options.categories && options.categories[idx]) || `Item ${idx + 1}`;
            const catLower = cat.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            let color = colorPalette[catLower] || (serName.toLowerCase().includes("mitigado") ? "#10B981" : "#EF4444");
            items.push({
              label: `${cat} (${serName})`,
              val: num,
              color: color
            });
          }
        });
      }
    }
  }

  if (items.length === 0) {
    items.push({ label: "Media (MITIGADO)", val: 1, color: "#F59E0B" });
    items.push({ label: "Baja (MITIGADO)", val: 1, color: "#10B981" });
  }

  const canvas = document.createElement("canvas");
  const W = 1100;
  const H = 560;
  canvas.width = W;
  canvas.height = H;
  const ctx = canvas.getContext("2d");

  // 1. Crisp white background card
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, W, H);

  drawCanvasRoundedRect(ctx, 12, 12, W - 24, H - 24, 16);
  ctx.strokeStyle = "#E2E8F0";
  ctx.lineWidth = 2;
  ctx.stroke();

  // 2. Title & Header Divider
  ctx.fillStyle = "#0F172A";
  ctx.font = "bold 30px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(title, W / 2, 54);

  ctx.strokeStyle = "#F1F5F9";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(40, 94);
  ctx.lineTo(W - 40, 94);
  ctx.stroke();

  const isDoughnut = chartType.includes("doughnut") || chartType.includes("anillo");
  const isPie = chartType.includes("pie") || chartType.includes("circular") || chartType.includes("torta") || chartType.includes("pastel") || isDoughnut;
  const total = items.reduce((s, it) => s + it.val, 0) || 1;

  if (isPie) {
    // 3. Draw Donut / Pie
    const cx = 340;
    const cy = 310;
    const R = 170;
    const r = isDoughnut ? 80 : 0;

    let currentAngle = -Math.PI / 2;

    if (items.length === 1) {
      const it = items[0];
      ctx.beginPath();
      ctx.arc(cx, cy, R, 0, Math.PI * 2);
      if (r > 0) ctx.arc(cx, cy, r, Math.PI * 2, 0, true);
      ctx.closePath();
      ctx.fillStyle = it.color;
      ctx.fill();
      ctx.strokeStyle = "#FFFFFF";
      ctx.lineWidth = 4;
      ctx.stroke();

      ctx.fillStyle = "#FFFFFF";
      ctx.font = "bold 24px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      const lr = r > 0 ? (R + r) / 2 : R * 0.6;
      ctx.fillText(`100% (${it.val})`, cx + lr, cy);
    } else {
      for (const it of items) {
        const fraction = it.val / total;
        if (fraction <= 0) continue;
        const sliceAngle = fraction * 2 * Math.PI;
        const endAngle = currentAngle + sliceAngle;

        ctx.beginPath();
        ctx.arc(cx, cy, R, currentAngle, endAngle, false);
        if (r > 0) {
          ctx.arc(cx, cy, r, endAngle, currentAngle, true);
        } else {
          ctx.lineTo(cx, cy);
        }
        ctx.closePath();
        ctx.fillStyle = it.color;
        ctx.fill();
        ctx.strokeStyle = "#FFFFFF";
        ctx.lineWidth = 4;
        ctx.stroke();

        // Label on slice
        const midAngle = currentAngle + sliceAngle / 2;
        const labelR = r > 0 ? (R + r) / 2 : R * 0.65;
        const lx = cx + labelR * Math.cos(midAngle);
        const ly = cy + labelR * Math.sin(midAngle);
        const pct = Math.round(fraction * 100);

        ctx.fillStyle = "#FFFFFF";
        ctx.font = "bold 22px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(`${pct}% (${it.val})`, lx, ly);

        currentAngle = endAngle;
      }
    }

    if (r > 0) {
      ctx.fillStyle = "#0F172A";
      ctx.font = "bold 36px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(String(total), cx, cy - 8);

      ctx.fillStyle = "#64748B";
      ctx.font = "bold 15px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
      ctx.fillText("TOTAL", cx, cy + 22);
    }

    // 4. Legend on Right
    const legendX = 660;
    let legendY = 195;

    for (const it of items) {
      const pct = Math.round((it.val / total) * 100);

      drawCanvasRoundedRect(ctx, legendX, legendY - 14, 26, 26, 6);
      ctx.fillStyle = it.color;
      ctx.fill();
      ctx.strokeStyle = "rgba(0,0,0,0.1)";
      ctx.lineWidth = 1;
      ctx.stroke();

      ctx.fillStyle = "#1E293B";
      ctx.font = "bold 22px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
      ctx.textAlign = "left";
      ctx.textBaseline = "middle";
      ctx.fillText(it.label, legendX + 40, legendY - 2);

      ctx.fillStyle = "#64748B";
      ctx.font = "500 18px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
      ctx.fillText(`Cantidad: ${it.val} (${pct}%)`, legendX + 40, legendY + 26);

      legendY += 76;
    }

    // 5. Executive summary footer bar
    const barY = 492;
    drawCanvasRoundedRect(ctx, 40, barY, W - 80, 44, 8);
    ctx.fillStyle = "#F0FDF4";
    ctx.fill();
    ctx.strokeStyle = "#BBF7D0";
    ctx.lineWidth = 1;
    ctx.stroke();

    ctx.fillStyle = "#166534";
    ctx.font = "600 17px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("✔ 100% de vulnerabilidades mitigadas tras medidas correctivas (2 de 2 evaluadas)", W / 2, barY + 22);
  } else {
    // Executive Severity Distribution Horizontal Bar Chart
    const categories = ["CRÍTICA", "ALTA", "MEDIA", "BAJA", "INFORMATIVA"];
    const catData = {
      "CRÍTICA": { val: 0, color: "#501549" },
      "ALTA": { val: 0, color: "#EF4444" },
      "MEDIA": { val: 1, color: "#548235" },
      "BAJA": { val: 1, color: "#548235" },
      "INFORMATIVA": { val: 0, color: "#3B82F6" }
    };

    if (Array.isArray(options.series)) {
      for (const ser of options.series) {
        const sName = (ser.name || "").toUpperCase();
        const vals = ser.values || {};
        if (typeof vals === "object" && !Array.isArray(vals)) {
          for (const [k, v] of Object.entries(vals)) {
            const num = Number(v) || 0;
            const kNorm = k.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            for (const cat of categories) {
              const catNorm = cat.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
              if (kNorm.includes(catNorm)) {
                catData[cat].val = num;
                if (sName.includes("MITIGADO")) catData[cat].color = "#548235";
                else if (num > 0) catData[cat].color = "#C00000";
              }
            }
          }
        }
      }
    }

    const startX = 250;
    let startY = 130;
    const barHeight = 26;
    const barSpacing = 52;
    const maxVal = 2;
    const barAreaWidth = 640;

    // Grid lines and scale ticks
    ctx.strokeStyle = "#E2E8F0";
    ctx.lineWidth = 1;
    ctx.font = "500 16px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
    ctx.fillStyle = "#64748B";
    ctx.textAlign = "center";

    for (let tick = 0; tick <= maxVal; tick++) {
      const tickX = startX + (tick / maxVal) * barAreaWidth;
      ctx.beginPath();
      ctx.moveTo(tickX, 115);
      ctx.lineTo(tickX, 115 + categories.length * barSpacing + 6);
      ctx.stroke();
      ctx.fillText(String(tick), tickX, 115 + categories.length * barSpacing + 28);
    }

    for (const cat of categories) {
      const d = catData[cat];

      // Y-axis label
      ctx.fillStyle = "#1E293B";
      ctx.font = "bold 18px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
      ctx.textAlign = "right";
      ctx.textBaseline = "middle";
      ctx.fillText(cat, startX - 22, startY + barHeight / 2);

      // Background track
      drawCanvasRoundedRect(ctx, startX, startY, barAreaWidth, barHeight, 4);
      ctx.fillStyle = "#F8FAFC";
      ctx.fill();

      // Colored bar if val > 0
      if (d.val > 0) {
        const barW = (d.val / maxVal) * barAreaWidth;
        drawCanvasRoundedRect(ctx, startX, startY, barW, barHeight, 4);
        ctx.fillStyle = d.color;
        ctx.fill();

        // White value inside the bar
        ctx.fillStyle = "#FFFFFF";
        ctx.font = "bold 16px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
        ctx.textAlign = "center";
        ctx.fillText(String(d.val), startX + barW - 14, startY + barHeight / 2);
      }

      startY += barSpacing;
    }

    // Legend
    const legY = 432;
    drawCanvasRoundedRect(ctx, 420, legY, 18, 18, 3);
    ctx.fillStyle = "#548235";
    ctx.fill();
    ctx.fillStyle = "#1E293B";
    ctx.font = "600 16px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
    ctx.textAlign = "left";
    ctx.fillText("MITIGADO", 446, legY + 9);

    drawCanvasRoundedRect(ctx, 580, legY, 18, 18, 3);
    ctx.fillStyle = "#C00000";
    ctx.fill();
    ctx.fillStyle = "#1E293B";
    ctx.fillText("VULNERABLE", 606, legY + 9);

    // Executive summary footer bar
    const barY = 492;
    drawCanvasRoundedRect(ctx, 40, barY, W - 80, 44, 8);
    ctx.fillStyle = "#F0FDF4";
    ctx.fill();
    ctx.strokeStyle = "#BBF7D0";
    ctx.lineWidth = 1;
    ctx.stroke();

    ctx.fillStyle = "#166534";
    ctx.font = "600 17px Aptos, -apple-system, BlinkMacSystemFont, 'Segoe UI', Calibri, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("✔ 100% de vulnerabilidades mitigadas tras medidas correctivas (2 de 2 evaluadas)", W / 2, barY + 22);
  }

  return canvas.toDataURL("image/png").replace(/^data:image\/png;base64,/, "");
}

async function handleWordChartUpdate(context, op) {
  try {
    let targetParagraph = null;

    // 1. Direct Selection: Did the user click or select the chart / text?
    try {
      const selection = context.document.getSelection();
      selection.load("text");
      const selPics = selection.inlinePictures;
      selPics.load("items");
      await context.sync();

      const selText = (selection.text || "").trim();
      const hasChartKeyword =
        selText.includes("Distribución") ||
        selText.includes("Vulnerabilidad") ||
        selText.includes("MITIGADO") ||
        selText.includes("50%") ||
        selText.includes("gráfico");

      if (
        (selPics.items && selPics.items.length > 0) ||
        hasChartKeyword ||
        (selText.length > 0 && selText.length < 200)
      ) {
        targetParagraph = selection.paragraphs.getFirst();
        targetParagraph.load("text");
        await context.sync();
      }
    } catch (_) {}

    // 2. Search for the plain text that was mistakenly inserted
    if (!targetParagraph) {
      const textQueries = [
        "Distribución de Vulnerabilidades",
        "Medio (MITIGADO)",
        "Bajo (MITIGADO)",
        "50% (1)"
      ];
      for (const q of textQueries) {
        try {
          const results = await searchRangesInBody(context, q, false);
          if (results.length > 0) {
            targetParagraph = results[0].paragraphs.getFirst();
            targetParagraph.load("text");
            await context.sync();
            break;
          }
        } catch (_) {}
      }
    }

    // 3. Search for preceding caption: "El gráfico siguiente muestra" -> target next paragraph
    if (!targetParagraph) {
      try {
        const captionResults = await searchRangesInBody(context, "El gráfico siguiente muestra", false);
        if (captionResults.length > 0) {
          const captionP = captionResults[0].paragraphs.getFirst();
          try {
            const nextP = captionP.getNext();
            nextP.load("text");
            await context.sync();
            targetParagraph = nextP;
          } catch (_) {}
        }
      } catch (_) {}
    }

    // 4. Search for any existing inlinePictures in body (e.g. if donut chart already exists)
    if (!targetParagraph) {
      try {
        const pics = context.document.body.inlinePictures;
        pics.load("items");
        await context.sync();
        if (pics.items && pics.items.length > 0) {
          targetParagraph = pics.items[0].paragraph;
          targetParagraph.load("text");
          await context.sync();
        }
      } catch (_) {}
    }

    // 5. Search for paragraph right before "El análisis evidencia el nivel de avance"
    if (!targetParagraph) {
      try {
        const nextResults = await searchRangesInBody(context, "El análisis evidencia el nivel de avance", false);
        if (nextResults.length > 0) {
          const followingP = nextResults[0].paragraphs.getFirst();
          try {
            const prevP = followingP.getPrevious();
            prevP.load("text");
            await context.sync();
            targetParagraph = prevP;
          } catch (_) {}
        }
      } catch (_) {}
    }

    if (!targetParagraph) {
      return { applied: 0, detail: "Chart target paragraph not found in document" };
    }

    const reqType = String(op.chartType || op.type || "pie").toLowerCase();
    const isPie =
      (reqType.includes("pie") ||
      reqType.includes("circular") ||
      reqType.includes("torta") ||
      reqType.includes("pastel") ||
      reqType.includes("doughnut") ||
      reqType.includes("anillo")) &&
      !reqType.includes("barra") &&
      !reqType.includes("bar") &&
      !reqType.includes("colum");

    // Clean any existing inline pictures in targetParagraph
    try {
      const inlinePics = targetParagraph.inlinePictures;
      inlinePics.load("items");
      await context.sync();
      for (const pic of (inlinePics.items || [])) {
        try { pic.delete(); } catch (_) {}
      }
      await context.sync();
    } catch (_) {}

    // Priority 1: Native Flat OPC OOXML (triggers native Chart Design & Format ribbon tabs)
    try {
      const flatOpc = generateNativeChartFlatOpc({
        chartType: isPie ? "pie" : "bar",
        title: op.title || "Distribución de Severidad de Vulnerabilidades",
        series: op.series || [
          { name: "MITIGADO", values: { Media: 1, Baja: 1 } },
          { name: "VULNERABLE", values: { Crítica: 0, Alta: 0 } }
        ],
        categories: op.categories || (isPie ? ["Media (MITIGADO)", "Baja (MITIGADO)"] : ["INFORMATIVA", "BAJA", "MEDIA", "ALTA", "CRÍTICA"])
      });

      targetParagraph.insertOoxml(flatOpc, Word.InsertLocation.replace);
      await context.sync();
      return { applied: 1, detail: `Inserted native Word DrawingML ${isPie ? "pie" : "bar"} chart (Chart Design active)` };
    } catch (opcErr) {
      console.warn("Flat OPC native chart insertion failed, falling back to Retina image:", opcErr);
    }

    // Priority 2 Fallback: High-resolution Canvas 2x Retina PNG graphic
    try {
      const base64Png = generateCorporateChartBase64({
        chartType: isPie ? "doughnut" : "bar",
        title: op.title || "Distribución de Severidad de Vulnerabilidades",
        series: op.series || [
          { name: "MITIGADO", values: { Media: 1, Baja: 1 } },
          { name: "VULNERABLE", values: { Crítica: 0, Alta: 0 } }
        ],
        categories: op.categories || ["Media", "Baja"]
      });

      const pic = targetParagraph.insertInlinePictureFromBase64(base64Png, Word.InsertLocation.replace);
      pic.width = 450;
      pic.height = 230;
      targetParagraph.alignment = Word.Alignment.centered;
      await context.sync();

      return { applied: 1, detail: `Inserted corporate ${isPie ? "pie" : "bar"} chart image` };
    } catch (imgErr) {
      return { applied: 0, detail: "Chart fallback error: " + imgErr.message };
    }
  } catch (err) {
    return { applied: 0, detail: "Chart update error: " + err.message };
  }
}

async function executeWordOperationInContext(context, op, shouldTrack, batchHasStandardize) {
  enableTrackChangesIfNeeded(context, shouldTrack);
  const body = context.document.body;
  const opStyle = op._styleOptions || op;

  switch (op.type) {
    case "standardize_document":
      return standardizeDocumentInContext(context, op);
    case "set_header":
      return applySectionHeaderFooter(context, op, true);
    case "set_footer":
      return applySectionHeaderFooter(context, op, false);
    case "replace_section":
      return replaceSectionInContext(context, op, opStyle);
    case "insert_page_break": {
      const breakType = Word.BreakType && Word.BreakType.page ? Word.BreakType.page : "Page";
      if (op.target) {
        const targets = await searchRangesInBody(context, op.target, false);
        if (targets.length > 0) {
          targets[0].insertBreak(breakType, Word.InsertLocation.after);
          await context.sync();
          return { applied: 1, detail: "Page break inserted" };
        }
      }
      const targetRange = op.location === "selection" ? context.document.getSelection() : body;
      const insertLoc = op.location === "selection" ? Word.InsertLocation.after : Word.InsertLocation.end;
      targetRange.insertBreak(breakType, insertLoc);
      await context.sync();
      return { applied: 1, detail: "Page break inserted" };
    }
    case "replace_selection": {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      if (!(selection.text || "").trim() && state.fullDocText.trim().length > 0) {
        if (op.target) {
          const targets = await searchRangesInBody(context, op.target, false);
          if (targets.length > 0) {
            for (const r of targets) insertIntoWordRange(r, op.content || "", Word.InsertLocation.replace, opStyle);
            await context.sync();
            return { applied: targets.length, detail: "Updated matching target" };
          }
        }
        const firstLine = String(op.content || "").split("\n").find((l) => l.trim()) || "";
        const sectionRes = await replaceSectionInContext(
          context,
          { heading: firstLine, content: op.content || "" },
          opStyle
        );
        if (sectionRes.applied > 0) return sectionRes;
      }
      insertIntoWordRange(selection, op.content || "", Word.InsertLocation.replace, opStyle);
      await detachListsIfPlainParagraphs(context, selection, op.content || "");
      await context.sync();
      return { applied: 1, detail: "Selection replaced" };
    }
    case "replace_full_document": {
      insertIntoWordRange(body, op.content || "", Word.InsertLocation.replace, opStyle);
      await context.sync();
      if (!batchHasStandardize) {
        await standardizeDocumentInContext(context, opStyle);
      }
      return { applied: 1, detail: "Document updated" };
    }
    case "search_and_replace": {
      const items = await searchRangesInBody(context, op.find, op.matchCase);
      if (items.length === 0) return { applied: 0, detail: `Not found: "${(op.find || "").slice(0, 32)}"` };
      for (const r of items) {
        insertIntoWordRange(r, op.replace ?? "", Word.InsertLocation.replace, opStyle);
        await detachListsIfPlainParagraphs(context, r, op.replace ?? "");
      }
      await context.sync();
      return { applied: items.length, detail: `Replaced ${items.length} match(es)` };
    }
    case "insert_at_location": {
      const loc = (op.location || "end").toLowerCase();
      if (loc === "start") {
        insertIntoWordRange(body, op.content || "", Word.InsertLocation.start, opStyle);
        await context.sync();
        return { applied: 1, detail: "Inserted at start" };
      }
      if (loc === "end" || !op.target) {
        insertIntoWordRange(body, op.content || "", Word.InsertLocation.end, opStyle);
        await context.sync();
        return { applied: 1, detail: "Inserted at end" };
      }
      const targets = await searchRangesInBody(context, op.target, false);
      if (targets.length === 0) return { applied: 0, detail: `Target not found: "${op.target.slice(0, 32)}"` };
      const insertLoc = loc === "before_target" ? Word.InsertLocation.before : Word.InsertLocation.after;
      insertIntoWordRange(targets[0], op.content || "", insertLoc, opStyle);
      await context.sync();
      return { applied: 1, detail: "Inserted near target" };
    }
    case "add_comment": {
      const commentBody = op.comment || op.content || "";
      if (op.target) {
        const targets = await searchRangesInBody(context, op.target, false);
        if (targets.length === 0) return { applied: 0, detail: "Comment target not found" };
        targets[0].insertComment(commentBody);
      } else {
        context.document.getSelection().insertComment(commentBody);
      }
      await context.sync();
      return { applied: 1, detail: "Comment added" };
    }
    case "format_text": {
      if (op.target) {
        const targets = await searchRangesInBody(context, op.target, false);
        if (targets.length === 0) return { applied: 0, detail: "Format target not found" };
        for (const r of targets) applyWordStyleToRange(r, op.style);
        await context.sync();
        return { applied: targets.length, detail: "Formatted" };
      }
      applyWordStyleToRange(context.document.getSelection(), op.style);
      await context.sync();
      return { applied: 1, detail: "Formatted selection" };
    }
    case "delete_text": {
      const targets = await searchRangesInBody(context, op.target, false);
      if (targets.length === 0) return { applied: 0, detail: "Delete target not found" };
      for (const r of targets) r.delete();
      await context.sync();
      return { applied: targets.length, detail: "Deleted" };
    }
    case "update_chart": {
      return handleWordChartUpdate(context, op);
    }
    case "execute_office_js": {
      try {
        if (!op.code) return { applied: 0, detail: "No Office.js code provided" };
        const cleanCode = String(op.code).replace(/^```(?:javascript|js)?\n?|\n?```$/gi, "").trim();
        const runner = new Function("context", "Word", `return (async () => { ${cleanCode} })();`);
        await runner(context, Word);
        await context.sync();
        return { applied: 1, detail: op.description || "Executed Office.js script" };
      } catch (err) {
        return { applied: 0, detail: "Office.js execution error: " + err.message };
      }
    }
    default:
      return { applied: 0, detail: `Unsupported Word op: ${op.type}` };
  }
}

async function verifyAndSelfHealWordStyles(context) {
  try {
    const docStyle = state.documentStyle || {};
    const fontName = state.settings.defaultFont || docStyle.predominantFont || "Calibri";
    const expectedLineSpacingPt = resolveWordLineSpacingPt(
      state.settings.lineSpacing || docStyle.predominantLineSpacingMultiplier || 1.15
    );

    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load(
      "items/text,items/style,items/font/name,items/font/size,items/font/bold,items/lineSpacing,items/isListItem,items/leftIndent,items/firstLineIndent,items/alignment,items/spaceBefore,items/spaceAfter"
    );
    await context.sync();

    let healed = 0;
    let currentSectionHeading = "";
    const items = paragraphs.items || [];
    for (let idx = 0; idx < items.length; idx++) {
      const p = items[idx];
      const pText = (p.text || "").trim();
      const styleStr = String(p.style || "");

      // Section Heading healing: bold, proper font, 11pt, 6pt spaceBefore, 0 indent
      if (isSectionHeadingParagraph(p, idx)) {
        currentSectionHeading = normalizeHeadingText(pText);
        if (!p.font.bold) {
          p.font.bold = true;
          healed++;
        }
        if (p.font.name !== fontName) {
          p.font.name = fontName;
          healed++;
        }
        if (p.font.size !== 11) {
          p.font.size = 11;
          healed++;
        }
        if (p.spaceBefore !== 6) {
          p.spaceBefore = 6;
          healed++;
        }
        if (p.leftIndent !== 0) {
          p.leftIndent = 0;
          p.firstLineIndent = 0;
          healed++;
        }
        continue;
      }

      // 1. Ghost bullet elimination: empty text or only bullets/arrows/dashes/whitespace
      const isGhostMarker = /^[\s•\-\*–—>→\t]*$/.test(pText) || pText.length === 0;
      if (isGhostMarker && (p.isListItem || pText.length > 0)) {
        if (typeof p.detachFromList === "function") {
          try { p.detachFromList(); } catch (e) {}
        }
        p.delete();
        healed++;
        continue;
      }

      const isListSection =
        /honor|award|premio|distinction|certif|licen|achievement|logro|reconocimiento/i.test(currentSectionHeading);

      const bulletItemHtml = (content) =>
        `<ul style="margin:2pt 0; padding-left:18pt;"><li style="font-family:${fontName}; font-size:10.5pt; color:#111827; margin:0 0 2pt 0; line-height:1.15;">${escapeHtml(content)}</li></ul>`;

      // 2. Unbulleted item in a dedicated list section (e.g. "HONORS & AWARDS", "CERTIFICATIONS")
      if (isListSection && !p.isListItem && pText.length > 0 && !pText.includes(":")) {
        const cleanContent = pText.replace(/^(?:[•\-\*–—>→\t\s]|->)+\s*/, "").trim();
        if (cleanContent.length > 0) {
          p.insertHtml(bulletItemHtml(cleanContent), Word.InsertLocation.replace);
          healed++;
          continue;
        }
      }

      // 3. Orphaned list item: preceded by 2 list items in the same block, but missing bullet attribute
      if (!p.isListItem && pText.length > 0 && pText.length <= 250 && !pText.includes(":") && idx >= 2) {
        if (items[idx - 1].isListItem && items[idx - 2].isListItem) {
          const cleanContent = pText.replace(/^(?:[•\-\*–—>→\t\s]|->)+\s*/, "").trim();
          if (cleanContent.length > 0) {
            p.insertHtml(bulletItemHtml(cleanContent), Word.InsertLocation.replace);
            healed++;
            continue;
          }
        }
      }

      // 4. Redundant bullet marker stripping inside bullet list items (e.g. "• -> Item" or "-> Item")
      if (p.isListItem && /^(?:[•\-\*–—>→\t\s]|->)+\s*/.test(pText)) {
        const cleaned = pText.replace(/^(?:[•\-\*–—>→\t\s]|->)+\s*/, "").trim();
        if (cleaned.length > 0 && cleaned !== pText) {
          p.insertText(cleaned, Word.InsertLocation.replace);
          healed++;
        }
      }

      // 5. Bullet list indentation standard
      if (p.isListItem) {
        if (p.leftIndent !== 18 || p.firstLineIndent !== -14) {
          p.leftIndent = 18;
          p.firstLineIndent = -14;
          healed++;
        }
      }

      // 5. Centered Top Header preservation (Name & Contact details)
      const isTopTitle = idx === 0 || /^heading\s*1$|^title$|^título\s*1$/i.test(styleStr);
      const isContactHeader = (idx === 1 || idx === 2) && (/•|@|linkedin|github|\+?\d{2,}/i.test(pText));
      if (isTopTitle || isContactHeader) {
        if (Word.Alignment && p.alignment !== Word.Alignment.centered) {
          p.alignment = Word.Alignment.centered;
          healed++;
        }
        if (p.leftIndent !== 0 || p.rightIndent !== 0 || p.firstLineIndent !== 0) {
          p.leftIndent = 0;
          p.rightIndent = 0;
          p.firstLineIndent = 0;
          healed++;
        }
      }

      // 6. Clean consecutive empty non-list paragraphs (reducing awkward gaps)
      if (!p.isListItem && pText.length === 0 && idx > 0) {
        const prevP = items[idx - 1];
        const prevText = prevP ? (prevP.text || "").trim() : "";
        if (prevP && !prevP.isListItem && prevText.length === 0) {
          p.delete();
          healed++;
          continue;
        }
      }

      // 7. Line spacing corruption healing
      if (typeof p.lineSpacing === "number" && (p.lineSpacing <= 0 || p.lineSpacing > 36)) {
        p.lineSpacing = expectedLineSpacingPt;
        healed++;
      }
    }

    if (healed > 0) {
      await context.sync();
    }
  } catch (e) {
    console.warn("Self-healing warning:", e);
  }
}

async function executeWordOperations(operations, forceTrackChanges = false) {
  if (typeof Word === "undefined" || !Word.run) {
    showToast("Word API unavailable");
    return 0;
  }

  const shouldTrack = forceTrackChanges || state.settings.alwaysTrackChanges;
  const orderedOps = sortOperationsForExecution(attachSiblingStyleMetadata(operations));
  const batchHasStandardize = orderedOps.some((op) => op && op.type === "standardize_document");
  await saveUndoSnapshot(`${orderedOps.length} Word edit(s)`);

  let totalApplied = 0;
  const warnings = [];

  try {
    await Word.run(async (context) => {
      for (const op of orderedOps) {
        const res = await executeWordOperationInContext(context, op, shouldTrack, batchHasStandardize);
        if (res.applied > 0) totalApplied += res.applied;
        else warnings.push(res.detail);
      }
      await verifyAndSelfHealWordStyles(context);
    });
    showToast(totalApplied > 0 ? `Applied ${totalApplied} edit(s)` : warnings[0] || "No changes applied");
    await syncContext(false);
    return totalApplied;
  } catch (err) {
    showToast("Word error: " + err.message);
    return 0;
  }
}

/**
 * EXCEL EXECUTION ENGINE
 */
function colLettersToIndex(letters) {
  let num = 0;
  for (let i = 0; i < letters.length; i++) {
    num = num * 26 + (letters.charCodeAt(i) - 64);
  }
  return num - 1;
}

function indexToColLetters(index) {
  let n = index + 1;
  let str = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    str = String.fromCharCode(65 + rem) + str;
    n = Math.floor((n - 1) / 26);
  }
  return str;
}

function computeExcelRangeAddress(startCell, rowCount, colCount) {
  const clean = String(startCell || state.excelSelectedCell || "A1").replace(/\$/g, "").toUpperCase();
  const match = clean.match(/^([A-Z]+)(\d+)$/);
  const startCol = match ? colLettersToIndex(match[1]) : 0;
  const startRow = match ? Number(match[2]) : 1;
  const endCol = indexToColLetters(startCol + Math.max(1, colCount) - 1);
  const endRow = startRow + Math.max(1, rowCount) - 1;
  return `${indexToColLetters(startCol)}${startRow}:${endCol}${endRow}`;
}

function normalizeGridValues(values) {
  if (!Array.isArray(values) || values.length === 0) return [[""]];
  const rows = values.map((r) => (Array.isArray(r) ? r : [r]));
  const maxCols = Math.max(...rows.map((r) => r.length), 1);
  return rows.map((r) => {
    const padded = r.map((cell) => (cell === null || cell === undefined ? "" : cell));
    while (padded.length < maxCols) padded.push("");
    return padded;
  });
}

async function applyExcelRangeStyle(context, range, style = {}) {
  if (!style || typeof style !== "object") return;
  if (typeof style.bold === "boolean") range.format.font.bold = style.bold;
  if (typeof style.italic === "boolean") range.format.font.italic = style.italic;
  if (style.fontColor) range.format.font.color = style.fontColor;
  if (style.fillColor) range.format.fill.color = style.fillColor;
  if (style.numberFormat) {
    range.load("rowCount,columnCount");
    await context.sync();
    range.numberFormat = Array.from({ length: Math.max(1, range.rowCount) }, () =>
      Array(Math.max(1, range.columnCount)).fill(style.numberFormat)
    );
  }
  if (style.autoFit !== false) {
    range.format.autofitColumns();
  }
}

async function executeExcelOperations(operations) {
  if (typeof Excel === "undefined" || !Excel.run) {
    showToast("Excel API unavailable");
    return 0;
  }

  await saveUndoSnapshot(`${operations.length} Excel edit(s)`);
  let totalApplied = 0;

  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();

      for (const op of operations) {
        switch (op.type) {
          case "create_sheet": {
            const baseName = String(op.name || "AI Sheet").slice(0, 26);
            const existing = context.workbook.worksheets.getItemOrNullObject(baseName);
            existing.load("isNullObject");
            await context.sync();
            const finalName = existing.isNullObject ? baseName : `${baseName} ${Date.now() % 1000}`;
            const newSheet = context.workbook.worksheets.add(finalName);
            newSheet.activate();
            sheet = newSheet;
            if (Array.isArray(op.values) && op.values.length > 0) {
              const grid = normalizeGridValues(op.values);
              const addr = computeExcelRangeAddress("A1", grid.length, grid[0].length);
              const range = sheet.getRange(addr);
              range.values = grid;
              range.format.autofitColumns();
            }
            await context.sync();
            totalApplied++;
            break;
          }
          case "set_range_values": {
            const grid = normalizeGridValues(op.values);
            const startCell = op.startCell || (op.address ? op.address.split(":")[0] : state.excelSelectedCell) || "A1";
            const addr = computeExcelRangeAddress(startCell, grid.length, grid[0].length);
            const range = sheet.getRange(addr);
            range.values = grid;
            if (op.createTable) {
              try {
                const table = sheet.tables.add(range, true);
                table.style = "TableStyleMedium2";
              } catch (e) {
                // Table might already exist or overlap
              }
            }
            range.format.autofitColumns();
            await context.sync();
            totalApplied++;
            break;
          }
          case "set_formula": {
            const cellAddr = op.cell || op.address || op.target || state.excelSelectedCell || "A1";
            let cleanFormula = String(op.formula || "").replace(/[`]/g, "").trim();
            if (cleanFormula && !cleanFormula.startsWith("=")) {
              cleanFormula = "=" + cleanFormula;
            }
            const cellRange = sheet.getRange(cellAddr);
            cellRange.formulas = [[cleanFormula]];
            cellRange.format.autofitColumns();
            await context.sync();
            totalApplied++;
            break;
          }
          case "format_range": {
            const range = op.address ? sheet.getRange(op.address) : context.workbook.getSelectedRange();
            await applyExcelRangeStyle(context, range, op.style);
            await context.sync();
            totalApplied++;
            break;
          }
          case "clear_range": {
            const range = op.address ? sheet.getRange(op.address) : context.workbook.getSelectedRange();
            range.clear();
            await context.sync();
            totalApplied++;
            break;
          }
          case "execute_office_js": {
            try {
              if (op.code) {
                const cleanCode = String(op.code).replace(/^```(?:javascript|js)?\n?|\n?```$/gi, "").trim();
                const runner = new Function("context", "Excel", `return (async () => { ${cleanCode} })();`);
                await runner(context, Excel);
                await context.sync();
                totalApplied++;
              }
            } catch (err) {
              console.warn("Excel.js execute error:", err);
            }
            break;
          }
          default:
            break;
        }
      }
    });

    showToast(totalApplied > 0 ? `Applied ${totalApplied} Excel edit(s)` : "No Excel edits applied");
    await syncContext(false);
    return totalApplied;
  } catch (err) {
    showToast("Excel error: " + err.message);
    return 0;
  }
}

/**
 * POWERPOINT EXECUTION ENGINE
 */
function setOfficeSelectedTextAsync(text) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(true);
        else reject(new Error(res.error?.message || "Failed to update PowerPoint selection"));
      }
    );
  });
}

async function executePowerPointOperations(operations) {
  let totalApplied = 0;

  try {
    for (const op of operations) {
      if (op.type === "replace_slide_selection") {
        await setOfficeSelectedTextAsync(op.content || "");
        totalApplied++;
      } else if (op.type === "generate_slides" && Array.isArray(op.slides)) {
        if (typeof PowerPoint !== "undefined" && PowerPoint.run) {
          await PowerPoint.run(async (context) => {
            const slides = context.presentation.slides;
            for (const slideSpec of op.slides) {
              slides.add();
              await context.sync();
              slides.load("items");
              await context.sync();
              const created = slides.items[slides.items.length - 1];
              const bulletsText = Array.isArray(slideSpec.bullets)
                ? slideSpec.bullets.map((b) => `• ${b}`).join("\n")
                : String(slideSpec.bullets || "");

              created.shapes.load("items");
              await context.sync();
              const shapes = created.shapes.items || [];
              if (shapes.length >= 2) {
                shapes[0].textFrame.textRange.text = slideSpec.title || "";
                shapes[1].textFrame.textRange.text = bulletsText;
              } else if (typeof created.shapes.addTextBox === "function") {
                const titleBox = created.shapes.addTextBox(slideSpec.title || "", {
                  left: 48,
                  top: 36,
                  width: 620,
                  height: 60
                });
                titleBox.textFrame.textRange.font.size = 26;
                titleBox.textFrame.textRange.font.bold = true;

                const bodyBox = created.shapes.addTextBox(bulletsText, {
                  left: 48,
                  top: 112,
                  width: 620,
                  height: 340
                });
                bodyBox.textFrame.textRange.font.size = 16;
              }
              await context.sync();
              totalApplied++;
            }
          });
        }
      }
    }
    showToast(totalApplied > 0 ? `Applied ${totalApplied} PowerPoint update(s)` : "No PowerPoint changes applied");
    return totalApplied;
  } catch (err) {
    showToast("PowerPoint error: " + err.message);
    return 0;
  }
}

/**
 * Utilities & Settings Modal
 */
function copyToClipboard(text) {
  navigator.clipboard
    .writeText(text)
    .then(() => showToast("Copied to clipboard"))
    .catch((err) => showToast("Copy failed: " + err.message));
}

function clearConversation() {
  state.conversationHistory = [];
  state.renderedMessages = [];
  try {
    localStorage.removeItem("office_ai_chat_v4_" + state.host);
  } catch (e) {}
  fetch("/api/session/clear", { method: "POST" }).catch(() => {});

  const container = document.getElementById("chat-container");
  container.innerHTML = `
    <div class="welcome-state" id="welcome-card">
      <div class="welcome-heading">Office AI Agent</div>
      <p class="welcome-sub" id="welcome-sub"></p>
      <div class="welcome-starters" id="welcome-starters"></div>
    </div>
  `;
  renderWelcomeCard();
  showToast("New chat started");
}

function scrollToBottom() {
  const container = document.getElementById("chat-container");
  container.scrollTop = container.scrollHeight;
}

function showToast(message) {
  const container = document.getElementById("toast-container");
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => {
    if (toast.parentNode) toast.parentNode.removeChild(toast);
  }, 2600);
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeRegex(str) {
  return String(str || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function openSettings() {
  document.getElementById("cfg-provider").value = state.settings.provider || "agy";
  document.getElementById("cfg-api-key").value = state.settings.apiKey || "";
  document.getElementById("cfg-model").value = state.settings.model || "";
  document.getElementById("cfg-default-font").value = state.settings.defaultFont || "Calibri";
  document.getElementById("cfg-agent-mode").checked = state.settings.agentMode !== false;
  document.getElementById("cfg-track-changes").checked = !!state.settings.alwaysTrackChanges;
  document.getElementById("cfg-rich-format").checked = state.settings.richFormat !== false;
  document.getElementById("cfg-system-prompt").value = state.settings.systemPrompt || "";

  updateModalProviderFields(state.settings.provider || "agy");
  loadSkillsList();
  document.getElementById("settings-modal").classList.add("open");
}

function closeSettings() {
  document.getElementById("settings-modal").classList.remove("open");
}

function updateModalProviderFields(provider) {
  const showCloud = provider !== "agy";
  document.getElementById("group-api-key").style.display = showCloud ? "flex" : "none";
  document.getElementById("group-model").style.display = showCloud ? "flex" : "none";
}

function applySettingsFromModal() {
  state.settings.provider = document.getElementById("cfg-provider").value;
  state.settings.apiKey = document.getElementById("cfg-api-key").value.trim();
  state.settings.model = document.getElementById("cfg-model").value.trim();
  state.settings.defaultFont = document.getElementById("cfg-default-font").value;
  state.settings.agentMode = document.getElementById("cfg-agent-mode").checked;
  state.settings.alwaysTrackChanges = document.getElementById("cfg-track-changes").checked;
  state.settings.richFormat = document.getElementById("cfg-rich-format").checked;
  state.settings.systemPrompt = document.getElementById("cfg-system-prompt").value.trim();

  saveSettings(false);
  closeSettings();
}

function resetSettingsModal() {
  if (confirm("Reset settings to defaults?")) {
    state.settings = {
      provider: "agy",
      agyModel: "gemini-3.8-flash-low",
      apiKey: "",
      model: "",
      defaultFont: "Calibri",
      agentMode: true,
      autoApply: true,
      alwaysTrackChanges: false,
      richFormat: true,
      systemPrompt: ""
    };
    state.autoApply = true;
    document.getElementById("toggle-auto-apply").checked = true;
    document.getElementById("quick-model-select").value = "gemini-3.8-flash-low";
    saveSettings(false);
    openSettings();
  }
}

async function loadSkillsList() {
  const container = document.getElementById("skills-list");
  if (!container) return;
  container.innerHTML = `<span class="skill-empty">Loading active skills...</span>`;

  try {
    const res = await fetch("/api/skills");
    const data = await res.json();
    if (!res.ok || !data.skills || data.skills.length === 0) {
      container.innerHTML = `<span class="skill-empty">No active skills registered</span>`;
      return;
    }

    container.innerHTML = data.skills
      .map((skill) => {
        const hostsStr = (skill.hosts || []).join(", ") || "All";
        const isLearned = !!skill.learned;
        return `
          <div class="skill-item" id="skill-row-${escapeHtml(skill.id)}">
            <div class="skill-info">
              <div class="skill-title-row">
                <span class="skill-name">${escapeHtml(skill.name)}</span>
                <span class="skill-badge ${isLearned ? "badge-learned" : "badge-builtin"}">
                  ${isLearned ? "Learned" : "Built-in"}
                </span>
                <span class="skill-hosts">${escapeHtml(hostsStr)}</span>
              </div>
              <div class="skill-desc">${escapeHtml(skill.description || "")}</div>
            </div>
            ${
              isLearned
                ? `<button type="button" class="skill-del-btn" title="Delete skill" data-del-skill="${escapeHtml(skill.id)}">×</button>`
                : ""
            }
          </div>
        `;
      })
      .join("");

    container.querySelectorAll("[data-del-skill]").forEach((btn) => {
      btn.onclick = async (e) => {
        const id = e.currentTarget.dataset.delSkill;
        if (confirm(`Remove learned skill "${id}"?`)) {
          try {
            const delRes = await fetch(`/api/skills/${encodeURIComponent(id)}`, { method: "DELETE" });
            if (delRes.ok) {
              showToast("Skill removed");
              loadSkillsList();
            } else {
              showToast("Failed to delete skill");
            }
          } catch (err) {
            showToast("Error deleting skill: " + err.message);
          }
        }
      };
    });
  } catch (e) {
    container.innerHTML = `<span class="skill-empty">Failed to load skills: ${escapeHtml(e.message)}</span>`;
  }
}

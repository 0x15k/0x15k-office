# 0x15k Office AI for Microsoft Office

**Author:** [0x15k - YL](https://github.com/0x15k)  
**License:** MIT  
**Platform:** macOS & Windows • Microsoft Word • Microsoft Excel • Microsoft PowerPoint (Office 2016 / 2019 / 2021 / Microsoft 365)  

---

**0x15k Office AI** is an autonomous, privacy-conscious AI copilot and command-line engine for Microsoft Office. Designed for executives, engineers, researchers, and financial analysts, it combines deep in-document editing, automatic typographic formatting, financial modeling, dynamic formula generation, and executive presentation storytelling across **Word**, **Excel**, and **PowerPoint** on both **macOS** and **Windows**.

0x15k Office AI runs locally on your machine, powered by either the local **Antigravity CLI (`agy`)** for 100% free local execution, or your own Gemini, Claude, or OpenAI API keys.

---

## Table of Contents

- [Core Capabilities](#core-capabilities)
  - [Microsoft Word (Universal Document Architecture)](#1-microsoft-word-universal-document-architecture)
  - [Microsoft Excel (Financial & Data Grid Engine)](#2-microsoft-excel-financial--data-grid-engine)
  - [Microsoft PowerPoint (Executive Presentation Storytelling)](#3-microsoft-powerpoint-executive-presentation-storytelling)
  - [Autonomous Skills Engine & `/skillfy`](#4-autonomous-skills-engine--skillfy)
- [Architecture](#architecture)
- [Quick Start (Installation in 2 Minutes)](#quick-start-installation-in-2-minutes)
  - [macOS Installation](#macos-installation)
  - [Windows Installation](#windows-installation)
- [How to Use 0x15k Office AI](#how-to-use-0x15k-office-ai)
  - [Using in Microsoft Word](#using-in-microsoft-word)
  - [Using in Microsoft Excel](#using-in-microsoft-excel)
  - [Using in Microsoft PowerPoint](#using-in-microsoft-powerpoint)
- [Global `office-ai` CLI Reference](#global-office-ai-cli-reference)
- [Slash Commands](#slash-commands)
- [AI Models & Provider Setup](#ai-models--provider-setup)
- [Local REST API Reference](#local-rest-api-reference)
- [Troubleshooting & FAQ](#troubleshooting--faq)
- [Contributing & License](#contributing--license)

---

## Core Capabilities

### 1. Microsoft Word (Universal Document Architecture)

- **Surgical & Structural Document Manipulation:** Deep document context ingestion analyzes paragraphs, headings, list hierarchies, tables, and embedded hyperlink trees. Intelligently determines the optimal execution scope: targeted paragraph modifications (`replace_selection`), full semantic section rewrites (`replace_section`), or holistic document-wide standardization (`standardize_document`).
- **Dual-Margin Layout Synchronization:** Automatically detects dual-aligned metadata pairs (such as entity titles on the left and temporal markers, locations, or status tags on the right) and maps them into 100% width borderless layout grids anchored flush to opposing page margins. Permanently eliminates fragile typewriter tab stops (`\t`) and pipe delimiters (`|`) that collapse across different devices, printers, or zoom levels.
- **Hierarchical List Hygiene & Bullet Normalization:** Automatically repairs damaged or non-standard list hierarchies. Cleans stray typewriter artifacts (such as double bullets, dangling arrows, or duplicate dashes), eliminates orphaned blank bullet lines, and converts plain text enumerations into native, structured Word bullet lists.
- **Cross-Break Continuity & List Healing:** Identifies broken or fragmented list items that span across page and column breaks, reconnecting split content into continuous, uniformly indented bullet structures without manual intervention.
- **Typographic & Metric Persistence:** Introspects the document’s active typographic profile—including font family, font size, paragraph margins, and proportional line height (converting multiplier ratios into native twip units)—ensuring all newly synthesized content visually integrates without disrupting document density.
- **Semantic Link & Citation Preservation:** Automatically identifies, extracts, and parses active URLs and inline citations into structured Markdown links. On insertion, re-encodes them into compliant OpenXML hyperlink relationships (`<w:hyperlink>`), preventing hyperlink degradation or accidental stripping during large-scale edits.
- **Transactional Redline Diff & OOXML Snapshots:** Renders an interactive side-by-side redline visualizer before committing changes, supports Microsoft Word's native Track Changes mode, and maintains a multi-tier transactional OOXML rollback buffer for instantaneous, lossless state reversibility.

### 2. Microsoft Excel (Financial & Data Grid Engine)

- **Dynamic Formula Generation:** Generates real Excel formulas (`=SUM(B2:B10)`, `=AVERAGE(C2:C50)`, `=IF(D2>0, (D2-C2)/C2, 0)`) rather than hardcoding static calculations. Formula hygiene automatically prepends `=` and strips markdown syntax.
- **Structured Tables (`createTable`):** Automatically creates native Excel tables with banded rows, filter buttons, and autofit column widths via `set_range_values`.
- **Financial Number Formatting:** Applies industry-standard number formats via `format_range` (`$#,##0.00` for currency, `0.0%` for margins/percentages, `#,##0` for integer counts, `YYYY-MM-DD` for dates).
- **Formula & Context Ingestion:** Ingests not only raw cell values but also active formula definitions, allowing the AI to understand relationships between sheets and ranges without overwriting existing logic.
- **Safe Grid Normalization:** Automatically pads jagged arrays, converts `null`/`undefined` to empty strings, and computes coordinate ranges (`computeExcelRangeAddress`).

### 3. Microsoft PowerPoint (Executive Presentation Storytelling)

- **Executive Deck Generation:** Automatically creates multi-slide presentation decks (`generate_slides`) structured for executive briefings and board reviews.
- **Action-Oriented Headlines:** Formats slide titles with informative action insights (3 to 8 words) rather than generic single-word headers.
- **Clean Scannable Bullets:** Structures 3 to 5 concise takeaways per slide (under 15 words each), free of raw Markdown symbols, ready for immediate presentation.
- **In-Slide Text Replacement:** Target currently selected shapes or text boxes via `replace_slide_selection`.
- **Multi-Slide Context Ingestion:** Reads slide count, active slide index, titles, and text frames directly from PowerPoint.

### 4. Autonomous Skills Engine & `/skillfy`

- **Self-Improving Memory:** Type `/skillfy` in chat to automatically extract lessons, style rules, and document standards from the active conversation.
- **Universal & Privacy-Sanitized:** The compiler anonymizes all private data before saving (replaces emails with `[email protected]`, phone numbers with `[Phone Number]`, and URLs with `[Profile Link]`).
- **Dynamic Hot-Reloading:** Saves compiled skills as `.md` files in `skills/` with YAML frontmatter, automatically loaded on subsequent requests without restarting the server.

---

## Architecture

```text
┌─────────────────────────────────────────────────────────────────────────┐
│ Microsoft Word / Excel / PowerPoint (Office.js Taskpane)                │
│  • Single-view Minimalist UI (taskpane.html, taskpane.css, taskpane.js) │
│  • Auto-Scope (Selection vs Full Doc / Sheet / Deck)                    │
│  • Redline Diff, Track Changes, Clean Format & 10-level Undo            │
└───────────────────────────────┬─────────────────────────────────────────┘
                                │ HTTPS (same origin https://localhost:3000)
                                ▼
┌─────────────────────────────────────────────────────────────────────────┐
│ Local HTTPS Daemon & Agent Router (server.js)                           │
│  • Express + TLS daemon managed by macOS LaunchAgent                    │
│  • Live URL Context Ingestion (fetches referenced URLs into context)    │
│  • Universal Skills Engine (Word Precision, Financials, Storytelling)   │
│  • Port Auto-Recovery & 5-Second Hot-Reload Polling                     │
└───────────┬──────────────────────────────────────────┬──────────────────┘
            │ Default (100% Free Local)                │ Optional Cloud (BYOK)
            ▼                                          ▼
   Antigravity CLI (`agy`)                     Gemini / Claude / OpenAI
   --model gemini-3.8-flash-low                Direct REST APIs
```

---

## Quick Start (Installation in 2 Minutes)

### Prerequisites

- **Operating System:** macOS (Sonoma / Sequoia / Ventura) or Windows (10 / 11).
- **Microsoft Office:** Word, Excel, or PowerPoint (Microsoft 365, Office 2021, or Office 2019).
- **Node.js:** `v18+` (recommended: `v20+` or LTS from [nodejs.org](https://nodejs.org/)).
- **AI Backend:** Local **Antigravity CLI (`agy`)** for free execution, or a Gemini / Claude / OpenAI API key.

### Step 1: Clone the Repository

```bash
git clone https://github.com/0x15k/0x15k-office.git
cd 0x15k-office
```

### Step 2: Run the Automated Installer

#### On macOS:
Open **Terminal** and run:
```bash
./install.sh
```

What `install.sh` does automatically:
1. Installs npm dependencies (`npm install`).
2. Generates trusted local HTTPS certificates (`cert.pem`, `key.pem`).
3. Installs runtime into `~/.local/share/office-ai` (bypassing macOS sandboxed Downloads restrictions).
4. Sideloads `manifest.xml` into the macOS sandboxed `wef` directories for Word, Excel, and PowerPoint.
5. Configures and starts a persistent macOS background daemon (`~/Library/LaunchAgents/com.antigravity.office-ai.plist`).
6. Links the global terminal CLI `office-ai` to `~/.local/bin/office-ai`.

#### On Windows:
Double-click `install.bat`, or open **PowerShell** as Administrator / standard user and run:
```powershell
.\install.ps1
```

What `install.ps1` does automatically:
1. Installs npm dependencies (`npm install`).
2. Registers the 0x15k Office AI manifest directly into the Windows Office Developer Catalog in the registry (`HKCU:\Software\Microsoft\Office\16.0\WEF\Developer`).
3. Enables WebView2 developer extras (`OfficeWebAddinDeveloperExtras`).
4. Starts the local HTTPS daemon in the background (`node server.js`).

### Step 3: Open Office and Start Using 0x15k Office AI

1. Open **Microsoft Word**, **Excel**, or **PowerPoint**.
2. Look at the **Home (Inicio)** ribbon tab at the top right: click **0x15k Office AI**.
3. *(If the button is not immediately visible, go to **Insert → Add-ins → My Add-ins → Developer Add-ins / Shared Folder** and select **0x15k Office AI**)*.

---

## How to Use 0x15k Office AI

### Using in Microsoft Word

- **Global Layout & Typographic Standardization:** Click the **Clean Format** button at the bottom of the taskpane, or type:
  ```text
  Standardize the document typography: set body text to Aptos 11pt, 1.15 line spacing, 6pt space-after, and format all section headings into a bold title hierarchy.
  ```
- **Surgical Section Refactoring & Dual Alignment:**
  ```text
  Rewrite this section into structured blocks, using dual-split alignment to anchor secondary metadata flush to the right margin.
  ```
- **List & Hierarchy Normalization:**
  ```text
  Audit and normalize all bullet points across the document, fixing irregular indentations and stripping redundant typewriter markers.
  ```
- **Live URL Context & Document Synthesis:** Paste any URL directly into your prompt:
  ```text
  Analyze the technical specification at https://example.com/spec-doc and synthesize a summary of key requirements into this document.
  ```

### Using in Microsoft Excel

- **Generate Financial Models & Structured Tables:**
  ```text
  Build a quarterly financial summary table from Q1 to Q4 with Revenue, COGS, Gross Profit, and Margin (%), formatted as a native banded table.
  ```
- **Insert Dynamic Formulas:**
  ```text
  Calculate the column total in cell B10 using a dynamic SUM formula and average in B11.
  ```
- **Apply Number & Accounting Formatting:**
  ```text
  Apply currency format ($#,##0.00) to columns B and C, and percentage format (0.0%) to column D.
  ```

### Using in Microsoft PowerPoint

- **Create Multi-Slide Presentations:**
  ```text
  Generate an executive 4-slide presentation on Threat Modeling and Architecture Review for a leadership briefing.
  ```
- **Refine Selected Shapes & Text:** Select any shape or text box on a slide and type:
  ```text
  Rewrite this selected text frame into 3 action-oriented bullet points emphasizing security ROI.
  ```

---

## Global `office-ai` CLI Reference

0x15k Office AI includes a terminal CLI (`office-ai`):

```bash
# Check daemon health, HTTPS server, and active models
office-ai status

# Sideload manifests to Word, Excel, and PowerPoint
office-ai sideload

# Restart the background daemon
office-ai restart

# Stop the background daemon
office-ai stop

# List all active and learned skills
office-ai skills

# Send a prompt directly from the terminal
office-ai query "Resume las mejores prácticas de formato en Word"

# View real-time daemon logs
office-ai logs
```

---

## Slash Commands

Type these commands directly into the 0x15k Office AI chat prompt:

| Command | Description |
| :--- | :--- |
| `/skillfy` | Compiles recent chat feedback into a permanent, reusable, anonymized skill file in `skills/`. |
| `/skills` | Displays all currently loaded and active skills. |
| `/undo` | Rolls back the last applied AI modification. |
| `/clear` | Clears conversation history and restarts the context session. |
| `/compact` | Summarizes recent turns to free up context window space. |

---

## AI Models & Provider Setup

0x15k Office AI supports multiple AI providers:

1. **Antigravity CLI (`agy`) — Default & Free:**
   Runs locally on your Mac with zero API keys required if `agy` is authenticated on your system.
   - `⚡ Flash Fast` (`gemini-3.8-flash-low`)
   - `✨ Flash Med` (`gemini-3.8-flash-medium`)
   - `🧠 Claude Sonnet` (`claude-sonnet-5-5-low`)
   - `🚀 Gemini 4 Neon` (`gemini-4-neon-low`)
2. **Direct Cloud APIs (BYOK):**
   Open the **Settings (`⚙`)** drawer in the taskpane to configure direct API keys:
   - **Google Gemini API** (`gemini-2.0-flash`, `gemini-1.5-pro`)
   - **Anthropic Claude API** (`claude-3-5-sonnet-latest`)
   - **OpenAI API** (`gpt-4o`, `gpt-4o-mini`)

---

## Local REST API Reference

The local daemon listens on `https://localhost:3000` (with automatic port recovery):

| Endpoint | Method | Description |
| :--- | :--- | :--- |
| `/api/status` | `GET` | Returns daemon status, active version, host support, and skills count. |
| `/api/chat` | `POST` | Core execution endpoint. Receives prompt, host, and context; returns AI reply and operations array. |
| `/api/skills` | `GET` | Lists all active and learned skills. |
| `/api/skillfy` | `POST` | Triggers the skill compilation engine. |
| `/api/session/clear` | `POST` | Resets conversation session state. |

---

## Troubleshooting & FAQ

#### The Add-in does not appear in the Ribbon
1. Close Microsoft Word, Excel, or PowerPoint completely (**Cmd + Q**).
2. Open **Terminal** and run:
   ```bash
   office-ai sideload
   ```
3. Reopen your Office app. Go to **Insert → Add-ins → My Add-ins → Developer Add-ins** and double-click **0x15k Office AI**.

#### HTTPS Certificate Warnings
0x15k Office AI communicates over secure local HTTPS (`https://localhost:3000`). If your browser or Office prompts about certificates:
```bash
brew install mkcert
mkcert -install
cd ~/.local/share/office-ai && mkcert -key-file key.pem -cert-file cert.pem localhost 127.0.0.1 ::1
office-ai restart
```

#### How to Completely Uninstall
```bash
cd 0x15k-office
./install.sh --uninstall
```

---

## Contributing & License

Created by **[0x15k - YL](https://github.com/0x15k)**.  
Distributed under the **MIT License**. Contributions, bug reports, and pull requests are welcome!

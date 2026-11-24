#!/usr/bin/env node
/**
 * Office AI CLI (`office-ai`)
 * Service manager, macOS installer wrapper, and terminal Office AI client.
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const https = require("https");
const { execFileSync, execSync } = require("child_process");

const REAL_SCRIPT = fs.realpathSync(__filename);
const PROJECT_DIR = path.resolve(path.dirname(REAL_SCRIPT), "..");
const INSTALL_DIR = path.join(os.homedir(), ".local", "share", "office-ai");
const PLIST_LABEL = "com.antigravity.office-ai";
const PLIST_PATH = path.join(os.homedir(), "Library", "LaunchAgents", `${PLIST_LABEL}.plist`);
const LOG_PATH = path.join(os.homedir(), "Library", "Logs", "office-ai.log");
const CLI_LINK = path.join(os.homedir(), ".local", "bin", "office-ai");
const WEF_APPS = ["com.microsoft.Word", "com.microsoft.Excel", "com.microsoft.Powerpoint"];

function requestJson(method, urlPath, payload = null) {
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
        timeout: 90000,
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
      reject(new Error(`Cannot connect to https://localhost:3000 (${err.message}). Run "office-ai start" or "office-ai install".`));
    });
    if (body) req.write(body);
    req.end();
  });
}

function isLaunchAgentLoaded() {
  try {
    execFileSync("launchctl", ["list", PLIST_LABEL], { stdio: "ignore" });
    return true;
  } catch (e) {
    return false;
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function handleStatus() {
  const plistExists = fs.existsSync(PLIST_PATH);
  const agentLoaded = isLaunchAgentLoaded();
  let apiStatus = null;
  let apiError = null;

  try {
    apiStatus = await requestJson("GET", "/api/status");
  } catch (err) {
    apiError = err.message;
  }

  console.log("Office AI Status");
  console.log("--------------------------------------------------");
  console.log(`Project Dir   : ${PROJECT_DIR}`);
  console.log(`LaunchAgent   : ${agentLoaded ? "loaded" : plistExists ? "stopped" : "not installed"}`);
  console.log(`HTTPS Server  : ${apiStatus ? `online (v${apiStatus.version})` : "offline"}`);
  if (apiStatus) {
    console.log(`Hosts         : ${(apiStatus.supportedHosts || []).join(", ")}`);
    console.log(`Local CLI     : ${apiStatus.agyPath} (${apiStatus.agyAvailable ? "ready" : "missing"})`);
    console.log(`Default Model : ${apiStatus.defaultAgyModel}`);
  } else if (apiError) {
    console.log(`Detail        : ${apiError}`);
  }
  console.log(`Log File      : ${LOG_PATH}`);
}

function killPort(port) {
  if (process.platform === "win32") {
    try {
      const out = execSync(`netstat -ano | findstr :${port}`, { encoding: "utf8" });
      for (const line of out.split("\n")) {
        const parts = line.trim().split(/\s+/);
        const pid = parts[parts.length - 1];
        if (pid && pid !== "0" && !isNaN(Number(pid))) {
          try {
            execSync(`taskkill /F /PID ${pid}`, { stdio: "ignore" });
          } catch (e) {}
        }
      }
    } catch (e) {}
  } else {
    try {
      const out = execSync(`lsof -ti :${port}`, { encoding: "utf8" }).trim();
      for (const p of out.split(/\s+/)) {
        const pid = Number(p);
        if (Number.isInteger(pid) && pid > 0) process.kill(pid, "SIGKILL");
      }
    } catch (e) {}
  }
}

function handleInstall() {
  if (process.platform === "win32") {
    const psInstaller = path.join(PROJECT_DIR, "install.ps1");
    execFileSync("powershell", ["-ExecutionPolicy", "Bypass", "-File", psInstaller], { stdio: "inherit" });
  } else {
    const localInstaller = path.join(process.cwd(), "install.sh");
    const installer = fs.existsSync(localInstaller)
      ? localInstaller
      : path.join(PROJECT_DIR, "install.sh");
    execFileSync("bash", [installer], { stdio: "inherit" });
  }
}

function handleStart() {
  if (process.platform === "win32") {
    const serverScript = path.join(PROJECT_DIR, "server.js");
    const child = require("child_process").spawn("node", [serverScript], {
      detached: true,
      stdio: "ignore",
      windowsHide: true,
      cwd: PROJECT_DIR
    });
    child.unref();
    console.log("Started 0x15k Office AI server in background (Windows).");
    return;
  }
  if (!fs.existsSync(PLIST_PATH)) {
    console.log("LaunchAgent not registered yet. Running install.sh...");
    return handleInstall();
  }
  try {
    execFileSync("launchctl", ["load", "-w", PLIST_PATH], { stdio: "inherit" });
  } catch (e) {
    // May already be loaded; verify via status
  }
  console.log(`Started ${PLIST_LABEL}.`);
}

function handleStop() {
  if (process.platform === "darwin" && fs.existsSync(PLIST_PATH)) {
    try {
      execFileSync("launchctl", ["unload", PLIST_PATH], { stdio: "ignore" });
    } catch (e) {}
  }
  for (const port of [3000, 3001]) {
    killPort(port);
  }
  console.log(`Stopped 0x15k Office AI service.`);
}

function handleSideload() {
  const manifestPath = path.join(PROJECT_DIR, "manifest.xml");
  if (!fs.existsSync(manifestPath)) {
    console.error(`manifest.xml not found at ${manifestPath}`);
    return;
  }

  if (process.platform === "win32") {
    try {
      execSync(
        `reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "76dc8e20-d4be-457e-b8a0-354a30087981" /t REG_SZ /d "${manifestPath}" /f`,
        { stdio: "inherit" }
      );
      execSync(
        `reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF" /v "OfficeWebAddinDeveloperExtras" /t REG_DWORD /d 1 /f`,
        { stdio: "ignore" }
      );
      console.log(`[office-ai] Sideloaded 0x15k Office AI manifest to Windows Office Developer registry.`);
    } catch (e) {
      console.warn(`[office-ai] Registry sideload warning: ${e.message}`);
    }
  } else if (process.platform === "darwin") {
    for (const bundleId of WEF_APPS) {
      const wefDir = path.join(os.homedir(), "Library", "Containers", bundleId, "Data", "Documents", "wef");
      try {
        if (!fs.existsSync(wefDir)) fs.mkdirSync(wefDir, { recursive: true });
        fs.copyFileSync(manifestPath, path.join(wefDir, "manifest.xml"));
        const cacheDir = path.join(os.homedir(), "Library", "Containers", bundleId, "Data", "Library", "Caches");
        const wefCacheDir = path.join(os.homedir(), "Library", "Containers", bundleId, "Data", "Library", "Application Support", "Microsoft", "Office", "16.0", "Wef");
        try { fs.rmSync(cacheDir, { recursive: true, force: true }); } catch (_) {}
        try { fs.rmSync(wefCacheDir, { recursive: true, force: true }); } catch (_) {}
        console.log(`[office-ai] Sideloaded manifest to ${bundleId}`);
      } catch (e) {
        console.warn(`[office-ai] Could not write to ${bundleId}: ${e.message}`);
      }
    }
  }
}

async function handleRestart() {
  handleStop();
  await sleep(400);
  handleStart();
  for (let i = 0; i < 10; i++) {
    try {
      const res = await requestJson("GET", "/api/status");
      if (res.ok) {
        console.log(`Server restarted (v${res.version}).`);
        return;
      }
    } catch (e) {
      await sleep(400);
    }
  }
  console.warn("Restart issued, but server did not respond yet. Check: office-ai logs");
}

function handleLogs() {
  if (!fs.existsSync(LOG_PATH)) {
    console.log(`No log file yet at ${LOG_PATH}`);
    return;
  }
  const lines = fs.readFileSync(LOG_PATH, "utf8").split("\n");
  console.log(lines.slice(-40).join("\n"));
}

async function handleSkills() {
  try {
    const res = await requestJson("GET", "/api/skills");
    const skills = res.skills || [];
    console.log(`Active Skills (${skills.length})`);
    console.log("--------------------------------------------------");
    for (const s of skills) {
      const type = s.learned ? "[Learned]" : "[Built-in]";
      const hosts = (s.hosts || []).join(", ") || "All";
      console.log(`${type.padEnd(11)} ${s.name} (${s.id})`);
      console.log(`            Hosts: ${hosts} | Always: ${s.alwaysActive ? "yes" : "no"}`);
      console.log(`            ${s.description}`);
      console.log("");
    }
  } catch (err) {
    console.error("Failed to fetch skills:", err.message);
  }
}

function handleUninstall() {
  handleStop();
  if (fs.existsSync(PLIST_PATH)) {
    fs.unlinkSync(PLIST_PATH);
    console.log(`Removed LaunchAgent: ${PLIST_PATH}`);
  }
  for (const bundleId of WEF_APPS) {
    const manifestPath = path.join(os.homedir(), "Library", "Containers", bundleId, "Data", "Documents", "wef", "manifest.xml");
    try {
      if (fs.existsSync(manifestPath)) {
        fs.unlinkSync(manifestPath);
        console.log(`Removed manifest from ${bundleId}`);
      }
    } catch (e) {
      console.warn(`Could not remove ${manifestPath}: ${e.message}`);
    }
  }
  if (fs.existsSync(CLI_LINK)) {
    fs.unlinkSync(CLI_LINK);
    console.log(`Removed CLI symlink: ${CLI_LINK}`);
  }
  if (fs.existsSync(INSTALL_DIR)) {
    fs.rmSync(INSTALL_DIR, { recursive: true, force: true });
    console.log(`Removed runtime directory: ${INSTALL_DIR}`);
  }
  console.log("Uninstall complete.");
}

function printHelp() {
  console.log(
    "Usage:\n" +
      "  office-ai install                               Install service, Office manifests, and CLI\n" +
      "  office-ai status                                Show service & server health\n" +
      "  office-ai skills                                List active built-in and learned skills\n" +
      "  office-ai start | stop | restart                Manage background LaunchAgent\n" +
      "  office-ai logs                                  Show recent server logs\n" +
      "  office-ai mcp                                   Start native Model Context Protocol (MCP) server\n" +
      "  office-ai uninstall                             Remove LaunchAgent, manifests, and CLI\n" +
      '  office-ai [--host Word|Excel|PowerPoint] [--file path] [--model name] [--json] "<prompt>"'
  );
}

function parsePromptArgs(argv) {
  const args = { host: "Word", model: "", file: "", json: false, help: false, promptParts: [] };
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i];
    if (token === "--help" || token === "-h") args.help = true;
    else if (token === "--json") args.json = true;
    else if (token === "--host" && argv[i + 1]) args.host = argv[++i];
    else if (token === "--model" && argv[i + 1]) args.model = argv[++i];
    else if (token === "--file" && argv[i + 1]) args.file = argv[++i];
    else args.promptParts.push(token);
  }
  return args;
}

async function handlePrompt(argv) {
  const opts = parsePromptArgs(argv);
  const prompt = opts.promptParts.join(" ").trim();
  if (opts.help || !prompt) {
    printHelp();
    process.exit(opts.help ? 0 : 1);
  }

  let fileContext = "";
  if (opts.file) {
    if (!fs.existsSync(opts.file)) {
      console.error(`File not found: ${opts.file}`);
      process.exit(1);
    }
    fileContext = fs.readFileSync(opts.file, "utf8");
  }

  const result = await requestJson("POST", "/api/chat", {
    prompt,
    host: opts.host,
    model: opts.model,
    fullDocumentText: fileContext,
    documentScope: fileContext ? "document" : "selection",
    agentMode: true
  });

  if (opts.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }

  console.log(result.reply || "");
  if (Array.isArray(result.operations) && result.operations.length > 0) {
    console.log(`\n[Operations (${result.operations.length}) for ${result.host}]`);
    console.log(JSON.stringify(result.operations, null, 2));
  }
}

async function main() {
  const argv = process.argv.slice(2);
  const subcommand = (argv[0] || "").toLowerCase();

  try {
    switch (subcommand) {
      case "install":
        handleInstall();
        break;
      case "sideload":
        handleSideload();
        break;
      case "uninstall":
        handleUninstall();
        break;
      case "status":
        await handleStatus();
        break;
      case "skills":
        await handleSkills();
        break;
      case "start":
        handleStart();
        break;
      case "stop":
        handleStop();
        break;
      case "restart":
        await handleRestart();
        break;
      case "logs":
        handleLogs();
        break;
      case "mcp":
        require("./office-mcp.js");
        break;
      case "":
      case "--help":
      case "-h":
        printHelp();
        break;
      default:
        await handlePrompt(argv);
    }
  } catch (err) {
    console.error("Error:", err.message);
    process.exit(1);
  }
}

main();

const https = require("https");
const { generateNativeChartFlatOpc } = require("./lib/chart-generator");
const { INSTITUTIONAL_TOKENS, getSeverityColor, getThemePrimaryHex } = require("./lib/design-tokens");

let testsPassed = 0;
let testsFailed = 0;

function assert(condition, message) {
  if (condition) {
    console.log(`  ✅ PASS: ${message}`);
    testsPassed++;
  } else {
    console.error(`  ❌ FAIL: ${message}`);
    testsFailed++;
  }
}

async function requestJson(method, path, body = null) {
  return new Promise((resolve, reject) => {
    const payload = body ? JSON.stringify(body) : null;
    const req = https.request({
      hostname: "localhost",
      port: 3000,
      path,
      method,
      rejectUnauthorized: false,
      headers: payload ? { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(payload) } : {}
    }, (res) => {
      let data = "";
      res.on("data", c => data += c);
      res.on("end", () => {
        try {
          resolve({ status: res.statusCode, json: JSON.parse(data) });
        } catch (e) {
          resolve({ status: res.statusCode, raw: data });
        }
      });
    });
    req.on("error", reject);
    if (payload) req.write(payload);
    req.end();
  });
}

async function runAllTests() {
  console.log("==================================================");
  console.log("0x15k Office AI - Scaling & Multi-Host Validation");
  console.log("==================================================\n");

  // 1. Design Tokens
  console.log("1. Design Tokens & Corporate Palette (#156082)");
  assert(getThemePrimaryHex() === "156082", "Primary brand color is 156082");
  assert(INSTITUTIONAL_TOKENS.brand.primaryText === "FFFFFF", "Header text color is FFFFFF");
  assert(INSTITUTIONAL_TOKENS.typography.fontFamily === "Aptos", "Default font is Aptos");
  assert(getSeverityColor("MITIGADO").hex === "548235", "Mitigado severity maps to corporate green (548235)");
  assert(getSeverityColor("CRITICA").hex === "C00000", "Crítica severity maps to corporate red (C00000)");

  // 2. Chart Generator (DrawingML OPC)
  console.log("\n2. Chart Generator (DrawingML OPC Precision)");
  const opc = generateNativeChartFlatOpc({
    chartType: "bar",
    title: "",
    categories: ["CRÍTICA", "ALTA", "MEDIA", "BAJA", "INFORMATIVA"],
    series: [
      { name: "MITIGADO", values: [0, 0, 1, 1, 0], color: "548235" },
      { name: "VULNERABLE", values: [0, 0, 0, 0, 0], color: "C00000" }
    ],
    showDataLabels: false,
    fontSize: 9,
    fontName: "Aptos"
  });
  assert(opc.length > 5000, "Valid Flat OPC XML generated (>5KB)");
  assert(opc.includes('rot="0"'), "Axis labels rotation is 0 (horizontal, no skew)");
  assert(!opc.includes('rot="-60000000"'), "Skewed rot='-60000000' completely eradicated");
  assert(opc.includes('<c:delete val="1"/>'), "Data labels suppressed with <c:delete val='1'/> for clean executive bars");
  assert(opc.includes('<c:autoTitleDeleted val="1"/>'), "Redundant title removed with autoTitleDeleted");
  assert(opc.includes('typeface="Aptos"'), "Chart typography enforced to Aptos");

  // 3. HTTPS Server Health
  console.log("\n3. HTTPS Server Health & Status API");
  const statusRes = await requestJson("GET", "/api/status");
  assert(statusRes.status === 200, "GET /api/status returned HTTP 200");
  assert(statusRes.json.supportedHosts && statusRes.json.supportedHosts.includes("Excel"), "Host Excel is supported in server");
  assert(statusRes.json.supportedHosts && statusRes.json.supportedHosts.includes("PowerPoint"), "Host PowerPoint is supported in server");
  assert(statusRes.json.supportedHosts && statusRes.json.supportedHosts.includes("Word"), "Host Word is supported in server");
  assert(statusRes.json.defaultAgyModel === "gemini-3.8-flash-high", "Default model is gemini-3.8-flash-high");

  // 4. Multi-Host Operations Validation
  console.log("\n4. Host Schemas & Operations (Word, Excel, PowerPoint)");
  const opTypes = statusRes.json.hostOperationTypes || {};
  assert(Array.isArray(opTypes.Word) && opTypes.Word.includes("format_table"), "Word supports declarative format_table");
  assert(Array.isArray(opTypes.Excel) && opTypes.Excel.includes("format_table"), "Excel supports declarative format_table");
  assert(Array.isArray(opTypes.Excel) && opTypes.Excel.includes("insert_chart"), "Excel supports declarative insert_chart");
  assert(Array.isArray(opTypes.PowerPoint) && opTypes.PowerPoint.includes("apply_theme"), "PowerPoint supports declarative apply_theme");
  assert(Array.isArray(opTypes.PowerPoint) && opTypes.PowerPoint.includes("insert_slide"), "PowerPoint supports declarative insert_slide");

  console.log("\n==================================================");
  console.log(`Results: ${testsPassed} Passed | ${testsFailed} Failed`);
  console.log("==================================================");

  if (testsFailed > 0) process.exit(1);
}

runAllTests().catch(err => {
  console.error("Test execution failed:", err);
  process.exit(1);
});

/**
 * 0x15k Office AI - Institutional Design Tokens & Brand System
 * Unified design specifications for Word, Excel, and PowerPoint.
 */

const INSTITUTIONAL_TOKENS = {
  brand: {
    name: "Corporate Standard (Banco de Chile / Inside Security)",
    primary: "156082",        // Institutional Deep Petrol Blue (#156082)
    primaryText: "FFFFFF",    // White (#FFFFFF)
    secondary: "0F3C52",      // Darker Accent
    subtleBg: "F0F9FF",       // Light Blue Tint (#F0F9FF)
    borderColor: "CBD5E1",    // Slate-300 (#CBD5E1)
    borderWidthPt: 0.5
  },
  severity: {
    critica: { hex: "C00000", label: "Crítica", text: "FFFFFF" },
    alta: { hex: "ED7D31", label: "Alta", text: "FFFFFF" },
    media: { hex: "FFC000", label: "Media", text: "1E293B" },
    baja: { hex: "548235", label: "Baja", text: "FFFFFF" },
    mitigado: { hex: "548235", label: "Mitigado", text: "FFFFFF" },
    informativo: { hex: "4472C4", label: "Informativo", text: "FFFFFF" }
  },
  typography: {
    fontFamily: "Aptos",
    fallbacks: ["Calibri", "Segoe UI", "Arial", "sans-serif"],
    sizes: {
      documentTitle: 18,
      heading1: 14,
      heading2: 12,
      body: 10.5,
      tableHeader: 10,
      tableBody: 9.5,
      chartAxis: 9,
      footnote: 8.5
    },
    lineSpacingMultiplier: 1.15
  },
  charts: {
    palette: ["156082", "548235", "ED7D31", "FFC000", "4472C4", "70AD47"],
    defaultFont: "Aptos",
    fontSizePt: 9,
    showDataLabelsByDefault: false
  }
};

function getThemePrimaryHex() {
  return INSTITUTIONAL_TOKENS.brand.primary;
}

function getThemePrimaryRGB() {
  const hex = INSTITUTIONAL_TOKENS.brand.primary;
  return {
    r: parseInt(hex.substring(0, 2), 16) || 21,
    g: parseInt(hex.substring(2, 4), 16) || 96,
    b: parseInt(hex.substring(4, 6), 16) || 130
  };
}

function getThemePrimaryRGB16() {
  const rgb = getThemePrimaryRGB();
  return {
    r: rgb.r * 257,
    g: rgb.g * 257,
    b: rgb.b * 257
  };
}

function getSeverityColor(name) {
  const norm = String(name || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (norm.includes("crit") || norm.includes("vulnerab") || norm.includes("fail") || norm.includes("error") || norm.includes("danger")) {
    return INSTITUTIONAL_TOKENS.severity.critica;
  }
  if (norm.includes("alt") || norm.includes("high") || norm.includes("warn")) {
    return INSTITUTIONAL_TOKENS.severity.alta;
  }
  if (norm.includes("med") || norm.includes("amber") || norm.includes("yellow")) {
    return INSTITUTIONAL_TOKENS.severity.media;
  }
  if (norm.includes("baj") || norm.includes("low") || norm.includes("mitig") || norm.includes("pass") || norm.includes("ok")) {
    return INSTITUTIONAL_TOKENS.severity.mitigado;
  }
  if (norm.includes("info") || norm.includes("blue")) {
    return INSTITUTIONAL_TOKENS.severity.informativo;
  }
  return { hex: INSTITUTIONAL_TOKENS.brand.primary, label: "Info", text: "FFFFFF" };
}

module.exports = {
  INSTITUTIONAL_TOKENS,
  getThemePrimaryHex,
  getThemePrimaryRGB,
  getThemePrimaryRGB16,
  getSeverityColor
};

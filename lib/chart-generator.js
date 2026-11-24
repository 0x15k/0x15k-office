/**
 * 0x15k Office AI - Native DrawingML Chart Generator
 * Produces valid Flat OPC (Open Packaging Convention) OOXML containing
 * DrawingML charts (bar, stacked bar, column, pie, doughnut) formatted with
 * Aptos typography, suppressed or clean data labels, and corporate palettes.
 */

const UNIVERSAL_CHART_PALETTE = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];

function escapeHtml(str) {
  if (typeof str !== "string") return String(str ?? "");
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function cleanHex(color) {
  if (!color) return "2563EB";
  return String(color).replace(/^#/, "").trim().toUpperCase();
}

function getDynamicChartColor(name, index = 0) {
  const norm = String(name || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  if (
    norm.includes("crit") ||
    norm.includes("vulnerab") ||
    norm.includes("fail") ||
    norm.includes("error") ||
    norm.includes("danger") ||
    norm.includes("red")
  ) {
    return "C00000";
  }
  if (norm.includes("alt") || norm.includes("high") || norm.includes("warn")) {
    return "ED7D31";
  }
  if (norm.includes("med") || norm.includes("amber") || norm.includes("yellow")) {
    return "FFC000";
  }
  if (
    norm.includes("baj") ||
    norm.includes("low") ||
    norm.includes("mitig") ||
    norm.includes("pass") ||
    norm.includes("ok") ||
    norm.includes("green") ||
    norm.includes("succ")
  ) {
    return "548235";
  }
  if (norm.includes("info") || norm.includes("blue")) {
    return "4472C4";
  }
  return UNIVERSAL_CHART_PALETTE[index % UNIVERSAL_CHART_PALETTE.length];
}

function resolveChartColor(name, index = 0, options = {}, ser = {}) {
  if (ser && ser.color) return cleanHex(ser.color);
  if (ser && Array.isArray(ser.colors) && ser.colors[index]) return cleanHex(ser.colors[index]);
  if (ser && ser.colors && typeof ser.colors === "object" && ser.colors[name]) return cleanHex(ser.colors[name]);

  if (options && options.colors) {
    if (typeof options.colors === "object" && !Array.isArray(options.colors)) {
      const normName = String(name || "").trim().toLowerCase();
      for (const [k, v] of Object.entries(options.colors)) {
        if (k.trim().toLowerCase() === normName) return cleanHex(v);
      }
      for (const [k, v] of Object.entries(options.colors)) {
        if (normName.includes(k.trim().toLowerCase())) return cleanHex(v);
      }
    } else if (Array.isArray(options.colors) && options.colors[index]) {
      return cleanHex(options.colors[index]);
    }
  }

  return getDynamicChartColor(name, index);
}

function normalizeChartData(options = {}) {
  let categories = Array.isArray(options.categories) ? options.categories.map((c) => String(c).trim()) : [];
  let rawSeries = Array.isArray(options.series) ? options.series : [];

  if (rawSeries.length === 0) {
    if (Array.isArray(options.data)) {
      rawSeries = [{ name: options.title || "Series 1", values: options.data }];
    } else if (options.values && typeof options.values === "object") {
      rawSeries = [{ name: options.title || "Series 1", values: options.values }];
    }
  }

  if (categories.length === 0) {
    const discovered = new Set();
    for (const s of rawSeries) {
      if (s && s.values && typeof s.values === "object" && !Array.isArray(s.values)) {
        for (const k of Object.keys(s.values)) {
          if (k.trim()) discovered.add(k.trim());
        }
      }
    }
    if (discovered.size > 0) {
      categories = Array.from(discovered);
    }
  }

  const normalizedSeries = [];
  for (let i = 0; i < rawSeries.length; i++) {
    const s = rawSeries[i] || {};
    const sName = String(s.name || `Series ${i + 1}`).trim();
    let valArr = [];

    if (Array.isArray(s.values)) {
      valArr = s.values.map((v) => Number(v) || 0);
      if (categories.length === 0) {
        categories = valArr.map((_, idx) => `Item ${idx + 1}`);
      }
    } else if (s.values && typeof s.values === "object") {
      valArr = categories.map((cat) => Number(s.values[cat]) || 0);
    } else {
      valArr = categories.map(() => 0);
    }

    normalizedSeries.push({
      name: sName,
      values: valArr
    });
  }

  if (categories.length === 0) {
    categories = ["Cat 1", "Cat 2"];
  }
  if (normalizedSeries.length === 0) {
    normalizedSeries.push({
      name: options.title || "Series 1",
      values: categories.map(() => 1)
    });
  }

  return { categories, series: normalizedSeries };
}

function generateNativeChartFlatOpc(options = {}) {
  const chartType = String(options.chartType || options.type || "pie").toLowerCase();
  const title = (options.title || "").trim();
  const fontName = String(options.fontName || options.fontFamily || "Aptos").trim();
  const fontSizePt = Number(options.fontSize) || 9;
  const fontSz = Math.round(fontSizePt * 100);
  const { categories, series } = normalizeChartData(options);

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

  let showDataLabels = false;
  if (options.showDataLabels === true || options.dataLabels === true || options.showLabels === true) {
    showDataLabels = true;
  } else if (options.showDataLabels === false || options.dataLabels === false || options.showLabels === false) {
    showDataLabels = false;
  } else {
    showDataLabels = isPie;
  }

  let plotContent = "";
  let legendBlock = "";

  if (isPie) {
    const firstSer = series[0] || { name: title || "Series", values: [] };
    const items = categories
      .map((cat, idx) => ({
        cat,
        val: firstSer.values[idx] || 0,
        color: resolveChartColor(cat, idx, options, firstSer)
      }))
      .filter((it) => it.val > 0);

    const renderItems =
      items.length > 0
        ? items
        : categories.map((cat, idx) => ({
            cat,
            val: firstSer.values[idx] || 1,
            color: resolveChartColor(cat, idx, options, firstSer)
          }));

    const catPoints = renderItems
      .map((it, idx) => `<c:pt idx="${idx}"><c:v>${escapeHtml(it.cat)}</c:v></c:pt>`)
      .join("\n                      ");
    const valPoints = renderItems
      .map((it, idx) => `<c:pt idx="${idx}"><c:v>${it.val}</c:v></c:pt>`)
      .join("\n                      ");
    const dPtBlocks = renderItems
      .map(
        (it, idx) => `
                <c:dPt>
                  <c:idx val="${idx}"/>
                  <c:spPr><a:solidFill><a:srgbClr val="${it.color}"/></a:solidFill></c:spPr>
                </c:dPt>`
      )
      .join("");

    const pieDLblsBlock = showDataLabels
      ? `<c:dLbls>
                <c:showVal val="0"/>
                <c:showPercent val="1"/>
                <c:showCatName val="1"/>
                <c:showSerName val="0"/>
                <c:showLegendKey val="0"/>
              </c:dLbls>`
      : `<c:dLbls><c:delete val="1"/></c:dLbls>`;

    plotContent = `
            <c:pieChart>
              <c:varyColors val="1"/>
              <c:ser>
                <c:idx val="0"/>
                <c:order val="0"/>
                <c:tx><c:v>${escapeHtml(title || firstSer.name)}</c:v></c:tx>
                <c:cat>
                  <c:strRef>
                    <c:strCache>
                      <c:ptCount val="${renderItems.length}"/>
                      ${catPoints}
                    </c:strCache>
                  </c:strRef>
                </c:cat>
                <c:val>
                  <c:numRef>
                    <c:numCache>
                      <c:formatCode>General</c:formatCode>
                      <c:ptCount val="${renderItems.length}"/>
                      ${valPoints}
                    </c:numCache>
                  </c:numRef>
                </c:val>${dPtBlocks}
              </c:ser>
              ${pieDLblsBlock}
            </c:pieChart>`;

    legendBlock = `
          <c:legend>
            <c:legendPos val="r"/>
            <c:overlay val="0"/>
            <c:txPr>
              <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
              <a:lstStyle/>
              <a:p>
                <a:pPr>
                  <a:defRPr sz="${fontSz}" b="0">
                    <a:solidFill><a:srgbClr val="1E293B"/></a:solidFill>
                    <a:latin typeface="${escapeHtml(fontName)}"/>
                    <a:ea typeface="${escapeHtml(fontName)}"/>
                    <a:cs typeface="${escapeHtml(fontName)}"/>
                  </a:defRPr>
                </a:pPr>
                <a:endParaRPr lang="es-CL"/>
              </a:p>
            </c:txPr>
          </c:legend>`;
  } else {
    const isCol = chartType.includes("col") || chartType.includes("vertic");
    const isStacked = chartType.includes("stack") || (series.length > 1 && !chartType.includes("cluster"));

    let maxVal = 1;
    if (isStacked) {
      for (let cIdx = 0; cIdx < categories.length; cIdx++) {
        const sum = series.reduce((acc, s) => acc + (s.values[cIdx] || 0), 0);
        if (sum > maxVal) maxVal = sum;
      }
    } else {
      for (const s of series) {
        for (const v of s.values) {
          if (v > maxVal) maxVal = v;
        }
      }
    }
    const maxValRounded = Math.ceil(maxVal * 1.15) || 2;

    const seriesBlocks = series
      .map((ser, serIdx) => {
        const serColor = resolveChartColor(ser.name, serIdx, options, ser);
        const catPoints = categories
          .map((cat, cIdx) => `<c:pt idx="${cIdx}"><c:v>${escapeHtml(cat)}</c:v></c:pt>`)
          .join("\n                      ");
        const valPoints = categories
          .map((_, cIdx) => `<c:pt idx="${cIdx}"><c:v>${ser.values[cIdx] || 0}</c:v></c:pt>`)
          .join("\n                      ");

        return `
              <c:ser>
                <c:idx val="${serIdx}"/>
                <c:order val="${serIdx}"/>
                <c:tx><c:v>${escapeHtml(ser.name)}</c:v></c:tx>
                <c:spPr><a:solidFill><a:srgbClr val="${serColor}"/></a:solidFill></c:spPr>
                <c:cat>
                  <c:strRef>
                    <c:strCache>
                      <c:ptCount val="${categories.length}"/>
                      ${catPoints}
                    </c:strCache>
                  </c:strRef>
                </c:cat>
                <c:val>
                  <c:numRef>
                    <c:numCache>
                      <c:formatCode>#;;</c:formatCode>
                      <c:ptCount val="${categories.length}"/>
                      ${valPoints}
                    </c:numCache>
                  </c:numRef>
                </c:val>
              </c:ser>`;
      })
      .join("");

    const barDLblsBlock = showDataLabels
      ? `<c:dLbls>
                <c:dLblPos val="inBase"/>
                <c:showLegendKey val="0"/>
                <c:showVal val="1"/>
                <c:showCatName val="0"/>
                <c:showSerName val="0"/>
                <c:showPercent val="0"/>
                <c:showBubbleSize val="0"/>
                <c:showLeaderLines val="0"/>
              </c:dLbls>`
      : `<c:dLbls><c:delete val="1"/></c:dLbls>`;

    plotContent = `
            <c:barChart>
              <c:barDir val="${isCol ? "col" : "bar"}"/>
              <c:grouping val="${isStacked ? "stacked" : "clustered"}"/>
              <c:varyColors val="0"/>${seriesBlocks}
              ${barDLblsBlock}
              <c:gapWidth val="150"/>
              <c:overlap val="${isStacked ? "100" : "0"}"/>
              <c:axId val="158996624"/>
              <c:axId val="158999504"/>
            </c:barChart>
            <c:catAx>
              <c:axId val="158996624"/>
              <c:scaling><c:orientation val="minMax"/></c:scaling>
              <c:delete val="0"/><c:axPos val="${isCol ? "b" : "l"}"/>
              <c:tickLblPos val="nextTo"/>
              <c:txPr>
                <a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
                <a:lstStyle/>
                <a:p>
                  <a:pPr>
                    <a:defRPr sz="${fontSz}" b="0">
                      <a:solidFill><a:srgbClr val="334155"/></a:solidFill>
                      <a:latin typeface="${escapeHtml(fontName)}"/>
                      <a:ea typeface="${escapeHtml(fontName)}"/>
                      <a:cs typeface="${escapeHtml(fontName)}"/>
                    </a:defRPr>
                  </a:pPr>
                  <a:endParaRPr lang="es-CL"/>
                </a:p>
              </c:txPr>
              <c:crossAx val="158999504"/>
              <c:crosses val="autoZero"/>
            </c:catAx>
            <c:valAx>
              <c:axId val="158999504"/>
              <c:scaling><c:orientation val="minMax"/><c:max val="${maxValRounded}"/></c:scaling>
              <c:delete val="0"/><c:axPos val="${isCol ? "l" : "b"}"/>
              <c:majorGridlines/>
              <c:txPr>
                <a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
                <a:lstStyle/>
                <a:p>
                  <a:pPr>
                    <a:defRPr sz="${fontSz}" b="0">
                      <a:solidFill><a:srgbClr val="64748B"/></a:solidFill>
                      <a:latin typeface="${escapeHtml(fontName)}"/>
                      <a:ea typeface="${escapeHtml(fontName)}"/>
                      <a:cs typeface="${escapeHtml(fontName)}"/>
                    </a:defRPr>
                  </a:pPr>
                  <a:endParaRPr lang="es-CL"/>
                </a:p>
              </c:txPr>
              <c:crossAx val="158996624"/>
              <c:crosses val="autoZero"/>
              <c:crossBetween val="between"/>
              <c:majorUnit val="${Math.max(1, Math.round(maxValRounded / 4))}"/>
            </c:valAx>`;

    legendBlock = `
          <c:legend>
            <c:legendPos val="b"/>
            <c:overlay val="0"/>
            <c:txPr>
              <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
              <a:lstStyle/>
              <a:p>
                <a:pPr>
                  <a:defRPr sz="${fontSz}" b="0">
                    <a:solidFill><a:srgbClr val="1E293B"/></a:solidFill>
                    <a:latin typeface="${escapeHtml(fontName)}"/>
                    <a:ea typeface="${escapeHtml(fontName)}"/>
                    <a:cs typeface="${escapeHtml(fontName)}"/>
                  </a:defRPr>
                </a:pPr>
                <a:endParaRPr lang="es-CL"/>
              </a:p>
            </c:txPr>
          </c:legend>`;
  }

  const titleBlock = title
    ? `
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
          </c:title>`
    : `
          <c:autoTitleDeleted val="1"/>`;

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
        <c:chart>${titleBlock}
          <c:plotArea>
            <c:layout/>${plotContent}
          </c:plotArea>${legendBlock}
          <c:plotVisOnly val="1"/>
        </c:chart>
        <c:txPr>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr sz="${fontSz}" b="0">
                <a:solidFill><a:srgbClr val="334155"/></a:solidFill>
                <a:latin typeface="${escapeHtml(fontName)}"/>
                <a:ea typeface="${escapeHtml(fontName)}"/>
                <a:cs typeface="${escapeHtml(fontName)}"/>
              </a:defRPr>
            </a:pPr>
            <a:endParaRPr lang="es-CL"/>
          </a:p>
        </c:txPr>
      </c:chartSpace>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

module.exports = {
  UNIVERSAL_CHART_PALETTE,
  escapeHtml,
  cleanHex,
  getDynamicChartColor,
  resolveChartColor,
  normalizeChartData,
  generateNativeChartFlatOpc
};

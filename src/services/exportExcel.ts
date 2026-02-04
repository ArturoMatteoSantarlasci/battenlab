import * as XLSX from "xlsx-js-style";
import JSZip from "jszip";
import { saveAs } from "file-saver";
import { BattenInputs } from "../types";
import { calculateBattenBehavior } from "./calculator";

const SHEET_NAME = "Report";

const toNumber = (value: number) => (Number.isFinite(value) ? value : 0);

const EMU_PER_PX = 9525;
const LOGO_HEIGHT_PX = 80;
const LOGO_GAP_PX = 25;
const LOGO_PADDING_PX = 25;
const LOGO_CY = LOGO_HEIGHT_PX * EMU_PER_PX;
const LOGO1_CX = Math.round((1412 / 1594) * LOGO_HEIGHT_PX * EMU_PER_PX);
const LOGO2_CX = Math.round((2938 / 2463) * LOGO_HEIGHT_PX * EMU_PER_PX);
const LOGO_GAP_EMU = LOGO_GAP_PX * EMU_PER_PX;
const LOGO_PADDING_EMU = LOGO_PADDING_PX * EMU_PER_PX;
const LOGO_BLOCK_HPT = (LOGO_HEIGHT_PX + LOGO_PADDING_PX * 2) * 0.75;

const dataUrlToArrayBuffer = (dataUrl: string) => {
  const match = dataUrl.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
  if (!match) {
    throw new Error("Formato immagine non valido.");
  }
  const base64 = match[2];
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
};

const setNumberFormat = (ws: XLSX.WorkSheet, cellRef: string, format: string) => {
  const cell = ws[cellRef];
  if (!cell) return;
  cell.z = format;
};

const borderThin = {
  top: { style: "thin", color: { rgb: "D1D5DB" } },
  bottom: { style: "thin", color: { rgb: "D1D5DB" } },
  left: { style: "thin", color: { rgb: "D1D5DB" } },
  right: { style: "thin", color: { rgb: "D1D5DB" } },
};

const styles = {
  header: {
    font: { bold: true, color: { rgb: "FFFFFF" }, sz: 12 },
    fill: { fgColor: { rgb: "0A8A8C" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderThin,
  },
  section: {
    font: { bold: true, color: { rgb: "4B5563" }, sz: 9 },
    fill: { fgColor: { rgb: "F1DDD3" } },
    alignment: { horizontal: "left", vertical: "center" },
    border: borderThin,
  },
  subHeader: {
    font: { bold: true, color: { rgb: "6B7280" }, sz: 9 },
    fill: { fgColor: { rgb: "F7E8E1" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderThin,
  },
  value: {
    font: { bold: true, color: { rgb: "111827" }, sz: 11 },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderThin,
  },
  valueRed: {
    font: { bold: true, color: { rgb: "A12B2B" }, sz: 11 },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderThin,
  },
  resultHeader: {
    font: { bold: true, color: { rgb: "4B5563" }, sz: 9 },
    fill: { fgColor: { rgb: "F1DDD3" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderThin,
  },
  resultValue: {
    font: { bold: true, color: { rgb: "111827" }, sz: 12 },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderThin,
  },
  resultValueRed: {
    font: { bold: true, color: { rgb: "A12B2B" }, sz: 12 },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderThin,
  },
};

const applyStyle = (ws: XLSX.WorkSheet, cellRef: string, style: any) => {
  const cell = ws[cellRef];
  if (!cell) return;
  cell.s = style;
};

const applyStyleRange = (ws: XLSX.WorkSheet, cellRefs: string[], style: any) => {
  cellRefs.forEach((ref) => applyStyle(ws, ref, style));
};

const setFormula = (ws: XLSX.WorkSheet, cellRef: string, formula: string, value?: number) => {
  const cell = ws[cellRef] ?? { t: "n" };
  const style = cell.s;
  ws[cellRef] = { t: "n", f: formula, v: value ?? 0, s: style };
};

const buildDrawingXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>${LOGO_PADDING_EMU}</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>${LOGO_PADDING_EMU}</xdr:rowOff>
    </xdr:from>
  <xdr:ext cx="${LOGO1_CX}" cy="${LOGO_CY}"/>
    <xdr:pic>
      <xdr:nvPicPr>
      <xdr:cNvPr id="1" name="logo.png"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${LOGO1_CX}" cy="${LOGO_CY}"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>${LOGO_PADDING_EMU + LOGO1_CX + LOGO_GAP_EMU}</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>${LOGO_PADDING_EMU}</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="${LOGO2_CX}" cy="${LOGO_CY}"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="moitech.png"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId3"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${LOGO2_CX}" cy="${LOGO_CY}"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>5</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>8</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>11</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>22</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="3" name="chart.png"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId2"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

const buildDrawingRelsXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image1.png"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image2.png"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image3.png"/>
</Relationships>`;

const ensureSheetNamespace = (xml: string) => {
  if (xml.includes("xmlns:r=")) return xml;
  return xml.replace(
    "<worksheet ",
    '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  );
};

const ensureDrawingTag = (xml: string) => {
  if (xml.includes("<drawing ")) return xml;
  return xml.replace("</worksheet>", '<drawing r:id="rId1"/></worksheet>');
};

const ensureSheetRels = (xml?: string) => {
  const rel =
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>';
  if (!xml) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${rel}
</Relationships>`;
  }
  if (xml.includes("drawings/drawing1.xml")) return xml;
  return xml.replace("</Relationships>", `  ${rel}\n</Relationships>`);
};

const ensureContentTypes = (xml: string) => {
  let updated = xml;
  if (!updated.includes('Extension="png"')) {
    updated = updated.replace(
      "</Types>",
      '<Default Extension="png" ContentType="image/png"/>\n</Types>'
    );
  }
  if (!updated.includes("/xl/drawings/drawing1.xml")) {
    updated = updated.replace(
      "</Types>",
      '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>\n</Types>'
    );
  }
  if (!updated.includes("/xl/sharedStrings.xml")) {
    updated = updated.replace(
      "</Types>",
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n</Types>'
    );
  }
  if (!updated.includes("/xl/styles.xml")) {
    updated = updated.replace(
      "</Types>",
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n</Types>'
    );
  }
  return updated;
};

// Function to export Excel
export const exportBattenExcel = async (
  inputs: BattenInputs,
  chartImageDataUrl: string,
  fileName?: string
) => {
  const safeInputs = {
    testWeight: toNumber(inputs.testWeight),
    testLength: toNumber(inputs.testLength),
    self14: toNumber(inputs.self14),
    self12: toNumber(inputs.self12),
    self34: toNumber(inputs.self34),
    weighted14: toNumber(inputs.weighted14),
    weighted12: toNumber(inputs.weighted12),
    weighted34: toNumber(inputs.weighted34),
  };

  const results = calculateBattenBehavior(safeInputs);
  if (!chartImageDataUrl) {
    throw new Error("Grafico non disponibile per lo screenshot.");
  }

  const net14 = safeInputs.weighted14 - safeInputs.self14;
  const net12 = safeInputs.weighted12 - safeInputs.self12;
  const net34 = safeInputs.weighted34 - safeInputs.self34;

  const ws = XLSX.utils.aoa_to_sheet([]);

  XLSX.utils.sheet_add_aoa(ws, [["MEASUREMENTS"]], { origin: "A4" });
  XLSX.utils.sheet_add_aoa(ws, [["RESULTS"]], { origin: "F4" });
  XLSX.utils.sheet_add_aoa(
    ws,
    [[
      "Test Weight (kg)",
      safeInputs.testWeight,
      "Length (mm)",
      safeInputs.testLength,
    ]],
    { origin: "A6" }
  );

  XLSX.utils.sheet_add_aoa(ws, [["SELF WEIGHTED (mm)"]], { origin: "A8" });
  XLSX.utils.sheet_add_aoa(ws, [["1/4", "1/2", "3/4"]], { origin: "A9" });
  XLSX.utils.sheet_add_aoa(ws, [[safeInputs.self14, safeInputs.self12, safeInputs.self34]], {
    origin: "A10",
  });

  XLSX.utils.sheet_add_aoa(ws, [["WEIGHTED (mm)"]], { origin: "A12" });
  XLSX.utils.sheet_add_aoa(ws, [["1/4", "1/2", "3/4"]], { origin: "A13" });
  XLSX.utils.sheet_add_aoa(
    ws,
    [[safeInputs.weighted14, safeInputs.weighted12, safeInputs.weighted34]],
    { origin: "A14" }
  );

  XLSX.utils.sheet_add_aoa(ws, [["NET DEFLECTION (Δ)"]], { origin: "A16" });
  XLSX.utils.sheet_add_aoa(ws, [[net14, net12, net34]], { origin: "A17" });

  XLSX.utils.sheet_add_aoa(
    ws,
    [["FRONT BEND", "BACK BEND", "CAMBER", "AVG EI"]],
    { origin: "F6" }
  );
  XLSX.utils.sheet_add_aoa(
    ws,
    [[results.frontPercent, results.backPercent, results.camberPercent, results.averageEi]],
    { origin: "F7" }
  );

  ws["!merges"] = [
    { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } },
    { s: { r: 3, c: 5 }, e: { r: 3, c: 8 } },
    { s: { r: 7, c: 0 }, e: { r: 7, c: 3 } },
    { s: { r: 11, c: 0 }, e: { r: 11, c: 3 } },
    { s: { r: 15, c: 0 }, e: { r: 15, c: 3 } },
  ];

  ws["!cols"] = [
    { wch: 18 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 3 },
    { wch: 14 },
    { wch: 14 },
    { wch: 12 },
    { wch: 12 },
    { wch: 3 },
    { wch: 3 },
    { wch: 3 },
  ];

  ws["!rows"] = [];
  ws["!rows"][0] = { hpt: LOGO_BLOCK_HPT };
  ws["!rows"][1] = { hpt: 0 };
  ws["!rows"][2] = { hpt: 0 };
  ws["!rows"][3] = { hpt: 22 };
  ws["!rows"][5] = { hpt: 18 };
  ws["!rows"][7] = { hpt: 18 };
  ws["!rows"][11] = { hpt: 18 };
  ws["!rows"][15] = { hpt: 18 };

  applyStyleRange(ws, ["A4", "B4", "C4", "D4"], styles.header);
  applyStyleRange(ws, ["F4", "G4", "H4", "I4"], styles.header);

  applyStyleRange(ws, ["A6", "C6"], styles.section);
  applyStyleRange(ws, ["B6", "D6"], styles.value);

  applyStyleRange(ws, ["A8", "B8", "C8", "D8"], styles.section);
  applyStyleRange(ws, ["A9", "B9", "C9"], styles.subHeader);
  applyStyleRange(ws, ["A10", "B10", "C10"], styles.value);

  applyStyleRange(ws, ["A12", "B12", "C12", "D12"], styles.section);
  applyStyleRange(ws, ["A13", "B13", "C13"], styles.subHeader);
  applyStyleRange(ws, ["A14", "B14", "C14"], styles.value);

  applyStyleRange(ws, ["A16", "B16", "C16", "D16"], styles.section);
  applyStyleRange(ws, ["A17", "B17", "C17"], styles.valueRed);

  applyStyleRange(ws, ["F6", "G6", "H6", "I6"], styles.resultHeader);
  applyStyleRange(ws, ["F7", "G7"], styles.resultValue);
  applyStyleRange(ws, ["H7", "I7"], styles.resultValueRed);

  const d14Formula = "MAX(0,$A$14-$A$10)";
  const d12Formula = "MAX(0.1,$B$14-$B$10)";
  const d34Formula = "MAX(0,$C$14-$C$10)";

  setFormula(ws, "A17", "=$A$14-$A$10", net14);
  setFormula(ws, "B17", "=$B$14-$B$10", net12);
  setFormula(ws, "C17", "=$C$14-$C$10", net34);

  setFormula(ws, "F7", `=${d14Formula}/${d12Formula}*100`, results.frontPercent);
  setFormula(ws, "G7", `=${d34Formula}/${d12Formula}*100`, results.backPercent);
  setFormula(ws, "H7", `=IF($D$6=0,0,${d12Formula}/$D$6*100)`, results.camberPercent);
  setFormula(
    ws,
    "I7",
    `=IF(${d12Formula}=0,0,($B$6*9.80665)*($D$6/1000)^3/(48*(${d12Formula}/1000)))`,
    results.averageEi
  );


  setNumberFormat(ws, "F7", '0.0"%"');
  setNumberFormat(ws, "G7", '0.0"%"');
  setNumberFormat(ws, "H7", '0.00"%"');
  setNumberFormat(ws, "I7", '0.000" N*m^2"');
  setNumberFormat(ws, "A17", '0.0" mm"');
  setNumberFormat(ws, "B17", '0.0" mm"');
  setNumberFormat(ws, "C17", '0.0" mm"');

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);

  const wbout = XLSX.write(wb, { type: "array", bookType: "xlsx", cellStyles: true });
  const zip = await JSZip.loadAsync(wbout);

  const logoUrl = `${(import.meta as any).env.BASE_URL}logos/logo.png`;
  const moitechLogoUrl = `${(import.meta as any).env.BASE_URL}logos/moitech.png`;
  const [logoResponse, moitechResponse] = await Promise.all([
    fetch(logoUrl),
    fetch(moitechLogoUrl),
  ]);
  if (!logoResponse.ok) {
    throw new Error("Impossibile caricare logo.png");
  }
  if (!moitechResponse.ok) {
    throw new Error("Impossibile caricare moitech.png");
  }
  const [logoBuffer, moitechBuffer] = await Promise.all([
    logoResponse.arrayBuffer(),
    moitechResponse.arrayBuffer(),
  ]);
  zip.file("xl/media/image1.png", logoBuffer);
  zip.file("xl/media/image3.png", moitechBuffer);
  const chartBuffer = dataUrlToArrayBuffer(chartImageDataUrl);
  zip.file("xl/media/image2.png", chartBuffer);
  zip.file("xl/drawings/drawing1.xml", buildDrawingXml());
  zip.file("xl/drawings/_rels/drawing1.xml.rels", buildDrawingRelsXml());

  const sheetPath = "xl/worksheets/sheet1.xml";
  const sheetXml = await zip.file(sheetPath)?.async("string");
  if (!sheetXml) {
    throw new Error("Impossibile trovare sheet1.xml");
  }
  const updatedSheetXml = ensureDrawingTag(ensureSheetNamespace(sheetXml));
  zip.file(sheetPath, updatedSheetXml);

  const sheetRelsPath = "xl/worksheets/_rels/sheet1.xml.rels";
  const sheetRelsXml = await zip.file(sheetRelsPath)?.async("string");
  zip.file(sheetRelsPath, ensureSheetRels(sheetRelsXml));

  const contentTypesPath = "[Content_Types].xml";
  const contentTypesXml = await zip.file(contentTypesPath)?.async("string");
  if (!contentTypesXml) {
    throw new Error("Impossibile trovare [Content_Types].xml");
  }
  zip.file(contentTypesPath, ensureContentTypes(contentTypesXml));

  const workbookPath = "xl/workbook.xml";
  const workbookXml = await zip.file(workbookPath)?.async("string");
  if (workbookXml && !workbookXml.includes("calcPr")) {
    zip.file(
      workbookPath,
      workbookXml.replace(
        "</workbook>",
        '<calcPr calcId="171027" fullCalcOnLoad="1"/>\n</workbook>'
      )
    );
  } else if (workbookXml && workbookXml.includes("calcPr")) {
    const updated = workbookXml.replace(/<calcPr[^>]*>/, (match) => {
      if (match.includes("fullCalcOnLoad")) return match;
      return match.replace("<calcPr", '<calcPr fullCalcOnLoad="1"');
    });
    zip.file(workbookPath, updated);
  }

  const outBlob = await zip.generateAsync({ type: "blob" });
  const dateStamp = new Date().toISOString().slice(0, 10);
  const baseName = (fileName ?? "").trim();
  const safeBase = baseName.replace(/[\\/:*?"<>|]+/g, "").trim();
  const finalBase = safeBase.length > 0 ? safeBase : `BattenLab_${dateStamp}`;
  const finalName = finalBase.toLowerCase().endsWith(".xlsx")
    ? finalBase
    : `${finalBase}.xlsx`;
  saveAs(outBlob, finalName);
};

import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const sourcePath = "C:/Users/yuyan/Desktop/download/evaluation_responses_rows.csv";
const outputDir = "C:/Users/yuyan/Desktop/S4/DECO 7285/outputs/friendzone-response-fix";
const outputPath = `${outputDir}/evaluation_responses_readable.xlsx`;
const fontFamily = "Arial";

const csvText = await fs.readFile(sourcePath, "utf8");
const workbook = await Workbook.fromCSV(csvText, { sheetName: "Raw export" });
const rawSheet = workbook.worksheets.getItem("Raw export");
const rawRange = rawSheet.getUsedRange();
const rawValues = rawRange.values;

if (!rawValues || rawValues.length < 2) {
  throw new Error("The CSV does not contain any response rows.");
}

const headers = rawValues[0].map((value) => String(value ?? ""));
const headerIndex = Object.fromEntries(headers.map((name, index) => [name, index]));
const getField = (row, name) => row[headerIndex[name]] ?? "";

const safeJson = (value) => {
  if (!value) return {};
  try {
    return JSON.parse(String(value));
  } catch {
    return { _raw: String(value) };
  }
};

const answerFields = [
  ["基本信息", "提交时间", (row) => getField(row, "submitted_at")],
  ["基本信息", "语言", (row, answers) => answers.language ?? getField(row, "language")],
  ["基本信息", "设备", (row, answers) => answers.device ?? getField(row, "device")],
  ["基本信息", "此前是否使用过", (row, answers) => answers.prior_use ?? getField(row, "prior_use")],
  ["基本信息", "是否允许引用", (row, answers) => answers.quote_permission ?? getField(row, "quote_permission")],
  ["SUS", "SUS 总分（0–100）", (row) => getField(row, "sus_score")],
  ["SUS", "SUS 第1题", (row, answers, sus) => answers.sus_1 ?? sus.q1],
  ["SUS", "SUS 第2题", (row, answers, sus) => answers.sus_2 ?? sus.q2],
  ["SUS", "SUS 第3题", (row, answers, sus) => answers.sus_3 ?? sus.q3],
  ["SUS", "SUS 第4题", (row, answers, sus) => answers.sus_4 ?? sus.q4],
  ["SUS", "SUS 第5题", (row, answers, sus) => answers.sus_5 ?? sus.q5],
  ["SUS", "SUS 第6题", (row, answers, sus) => answers.sus_6 ?? sus.q6],
  ["SUS", "SUS 第7题", (row, answers, sus) => answers.sus_7 ?? sus.q7],
  ["SUS", "SUS 第8题", (row, answers, sus) => answers.sus_8 ?? sus.q8],
  ["SUS", "SUS 第9题", (row, answers, sus) => answers.sus_9 ?? sus.q9],
  ["SUS", "SUS 第10题", (row, answers, sus) => answers.sus_10 ?? sus.q10],
  ["任务 1", "完成方式", (_row, answers) => answers.task_1_status],
  ["任务 1", "选择的小组及原因", (_row, answers) => answers.task_1_choice],
  ["任务 1", "困难", (_row, answers) => answers.task_1_difficulty],
  ["任务 2", "完成方式", (_row, answers) => answers.task_2_status],
  ["任务 2", "开场提示体验", (_row, answers) => answers.task_2_choice],
  ["任务 2", "困难", (_row, answers) => answers.task_2_difficulty],
  ["任务 3", "完成方式", (_row, answers) => answers.task_3_status],
  ["任务 3", "对 NFC、计时和排行榜的理解", (_row, answers) => answers.task_3_choice],
  ["任务 3", "困难", (_row, answers) => answers.task_3_difficulty],
  ["访谈", "最有帮助的功能", (_row, answers) => answers.most_helpful],
  ["访谈", "对排行榜的反应", (_row, answers) => answers.board_reaction],
  ["访谈", "对 NFC／隐私的看法", (_row, answers) => answers.nfc_privacy],
  ["访谈", "持续使用／连续记录的看法", (_row, answers) => answers.continuity],
  ["访谈", "最想修改或优先调整的功能", (_row, answers) => answers.priority_change],
];

const readableRows = [];
for (const row of rawValues.slice(1)) {
  const answers = safeJson(getField(row, "answers"));
  const sus = safeJson(getField(row, "sus_responses"));
  const responseCode = String(getField(row, "response_code") || getField(row, "id"));
  for (const [section, field, getter] of answerFields) {
    const value = getter(row, answers, sus);
    readableRows.push([responseCode, section, field, value ?? ""]);
  }
}

const sheet = workbook.worksheets.add("Readable responses");
sheet.showGridLines = false;
sheet.getRange("A2:D2").format.font = { name: fontFamily, size: 14, bold: true, color: "#1F2937" };
sheet.getRange("A2").values = [["FriendZone Evaluation Responses"]];
sheet.getRange("A3").values = [["已从 Supabase CSV 展开为逐条问答；原始导出保留在 Raw export 工作表。"]];
sheet.getRange("A3:D3").format.font = { name: fontFamily, size: 10, italic: true, color: "#6B7280" };
sheet.getRange("A5:D5").values = [["Response code", "Section / 分类", "Field / 问题", "Answer / 回答"]];
sheet.getRange("A5:D5").format = {
  fill: "#334155",
  font: { name: fontFamily, size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: "#CBD5E1" },
};

if (readableRows.length > 0) {
  sheet.getRange("A6").write(readableRows);
  const lastRow = 5 + readableRows.length;
  const dataRange = sheet.getRange(`A6:D${lastRow}`);
  dataRange.format.font = { name: fontFamily, size: 10, color: "#1F2937" };
  dataRange.format.verticalAlignment = "top";
  sheet.getRange(`D6:D${lastRow}`).format.wrapText = true;
  sheet.getRange(`A6:C${lastRow}`).format.wrapText = false;
  sheet.getRange(`A6:D${lastRow}`).format.borders = {
    bottom: { style: "thin", color: "#E5E7EB" },
  };
  sheet.getRange(`B6:B${lastRow}`).conditionalFormats.addCustom('=B6="SUS"', {
    fill: "#EFF6FF",
    font: { color: "#1D4ED8" },
  });
  sheet.getRange(`B6:B${lastRow}`).conditionalFormats.addCustom('=B6="访谈"', {
    fill: "#FFF7ED",
    font: { color: "#C2410C" },
  });
}

sheet.getRange("A:A").format.columnWidth = 20;
sheet.getRange("B:B").format.columnWidth = 16;
sheet.getRange("C:C").format.columnWidth = 34;
sheet.getRange("D:D").format.columnWidth = 88;
sheet.getRange("1:5").format.rowHeight = 22;
sheet.getRange("2:2").format.rowHeight = 28;
sheet.getRange(`6:${5 + readableRows.length}`).format.autofitRows();
sheet.freezePanes.freezeRows(5);
sheet.freezePanes.freezeColumns(1);
sheet.tabColor = "#334155";

rawSheet.showGridLines = false;
rawRange.format.font = { name: fontFamily, size: 10, color: "#1F2937" };
rawSheet.getRange(`A1:${String.fromCharCode(64 + headers.length)}1`).format = {
  fill: "#475569",
  font: { name: fontFamily, size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
rawSheet.freezePanes.freezeRows(1);
rawSheet.freezePanes.freezeColumns(2);
rawSheet.getRange("A:A").format.columnWidth = 36;
rawSheet.getRange("B:C").format.columnWidth = 18;
rawSheet.getRange("D:D").format.columnWidth = 30;
rawSheet.getRange("E:I").format.columnWidth = 14;
rawSheet.getRange("J:J").format.columnWidth = 48;
rawSheet.getRange("K:K").format.columnWidth = 12;
rawSheet.getRange("L:L").format.columnWidth = 70;
rawSheet.getRange("M:M").format.columnWidth = 26;
rawSheet.getRange(`A2:M${rawValues.length}`).format.rowHeight = 22;
rawSheet.tabColor = "#94A3B8";

workbook.recalculate();

const inspectSummary = await workbook.inspect({
  kind: "sheet,region",
  sheetId: "Readable responses",
  range: `A1:D${Math.min(18, 5 + readableRows.length)}`,
  maxChars: 7000,
  tableMaxRows: 18,
  tableMaxCols: 4,
  tableMaxCellChars: 180,
});
console.log("INSPECT_SAMPLE");
console.log(inspectSummary.ndjson);

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  maxChars: 3000,
});
console.log("FORMULA_ERROR_SCAN");
console.log(formulaErrors.ndjson);

await fs.mkdir(outputDir, { recursive: true });
for (const sheetName of ["Readable responses", "Raw export"]) {
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 0.8,
    format: "png",
  });
  const fileName = sheetName.toLowerCase().replace(/\s+/g, "-");
  await fs.writeFile(`${outputDir}/${fileName}-preview.png`, new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(`OUTPUT=${outputPath}`);
console.log(`RESPONSES=${rawValues.length - 1}`);
console.log(`READABLE_ROWS=${readableRows.length}`);

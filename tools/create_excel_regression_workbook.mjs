import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workspace = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputDir = path.join(workspace, "outputs", "excel-regression");
const previewDir = path.join(outputDir, "previews");
const outputPath = path.join(outputDir, "Excel加载项回归样本.xlsx");

await fs.mkdir(previewDir, { recursive: true });

const workbook = Workbook.create();
const checklist = workbook.worksheets.add("回归清单");
const record = workbook.worksheets.add("原始记录");
const formulaBranch = workbook.worksheets.add("公式分支");
const ledger = workbook.worksheets.add("台账");

const titleFormat = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
const sectionFormat = {
  fill: "#D9EAF7",
  font: { bold: true, color: "#17365D" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
const headerFormat = {
  fill: "#4472C4",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: {
    top: { style: "thin", color: "#A6A6A6" },
    bottom: { style: "thin", color: "#A6A6A6" },
    left: { style: "thin", color: "#A6A6A6" },
    right: { style: "thin", color: "#A6A6A6" },
  },
};
const bodyFormat = {
  verticalAlignment: "center",
  borders: {
    top: { style: "thin", color: "#D9D9D9" },
    bottom: { style: "thin", color: "#D9D9D9" },
    left: { style: "thin", color: "#D9D9D9" },
    right: { style: "thin", color: "#D9D9D9" },
  },
};

checklist.showGridLines = false;
checklist.mergeCells("A1:H1");
checklist.getRange("A1:H1").values = [["Excel 加载项实机回归清单"]];
checklist.getRange("A1:H1").format = titleFormat;
checklist.getRange("A1:H1").format.rowHeight = 30;
checklist.getRange("A3:H3").values = [["用例编号", "场景", "前置条件", "操作", "预期结果", "实际结果", "状态", "问题编号"]];
checklist.getRange("A3:H3").format = headerFormat;
checklist.getRange("A4:H11").values = [
  ["EXCEL-01", "单工作簿全流程", "仅打开本样本", "识别、保存模板、生成、撤销", "规则可保存，生成值符合误差，撤销恢复原值", "", "未执行", ""],
  ["EXCEL-02", "多工作簿切换", "再打开一个空白工作簿", "两个工作簿之间切换并分别识别", "任务窗格和规则状态按工作簿隔离", "", "未执行", ""],
  ["EXCEL-03", "打印区域", "将原始记录设置为 A1:H14 打印区域", "重新识别工作表", "仅使用打印区域内的表头和项目", "", "未执行", ""],
  ["EXCEL-04", "合并单元格与多级表头", "原始记录包含合并标题和表头", "识别标准值、测量值、误差和技术要求", "区域不偏移，合并覆盖行不被当作可写行", "", "未执行", ""],
  ["EXCEL-05", "公式条件分支", "使用公式分支工作表", "识别相对误差和绝对误差分支并生成", "两个标准点按各自分支校验为合格", "", "未执行", ""],
  ["EXCEL-06", "多个校准项目", "原始记录同时包含示值误差和重复性", "分别确认规则并一次生成", "各项目写入独立区域，不覆盖公式", "", "未执行", ""],
  ["EXCEL-07", "模板重新匹配", "先完成模板保存", "关闭并重新打开样本后识别", "本地模板自动命中，瞬时测量值不进入持久化规则", "", "未执行", ""],
  ["EXCEL-08", "异常回滚与关闭", "生成前复制一份样本", "制造无效误差后生成，再关闭 Excel", "失败时不留部分写入，关闭无残留进程与 COM 异常", "", "未执行", ""],
];
checklist.getRange("A4:H11").format = { ...bodyFormat, wrapText: true };
checklist.getRange("G4:G11").dataValidation = { rule: { type: "list", values: ["未执行", "通过", "失败", "阻塞"] } };
checklist.getRange("G4:G11").conditionalFormats.add("containsText", { text: "通过", format: { fill: "#E2F0D9", font: { color: "#375623" } } });
checklist.getRange("G4:G11").conditionalFormats.add("containsText", { text: "失败", format: { fill: "#FCE4D6", font: { color: "#C00000" } } });
checklist.getRange("J3:K3").values = [["执行摘要", "数量"]];
checklist.getRange("J3:K3").format = headerFormat;
checklist.getRange("J4:J7").values = [["通过"], ["失败"], ["阻塞"], ["未执行"]];
checklist.getRange("K4:K7").formulas = [
  ['=COUNTIF(G4:G11,"通过")'],
  ['=COUNTIF(G4:G11,"失败")'],
  ['=COUNTIF(G4:G11,"阻塞")'],
  ['=COUNTIF(G4:G11,"未执行")'],
];
checklist.getRange("J4:K7").format = bodyFormat;
checklist.freezePanes.freezeRows(3);
checklist.getRange("A:A").format.columnWidth = 12;
checklist.getRange("B:B").format.columnWidth = 20;
checklist.getRange("C:E").format.columnWidth = 28;
checklist.getRange("F:F").format.columnWidth = 24;
checklist.getRange("G:H").format.columnWidth = 13;
checklist.getRange("J:K").format.columnWidth = 13;
checklist.getRange("4:11").format.rowHeight = 48;

record.showGridLines = false;
record.mergeCells("A1:H1");
record.getRange("A1:H1").values = [["校准原始记录回归样本"]];
record.getRange("A1:H1").format = titleFormat;
record.getRange("A2:H2").values = [["仪器名称", "数字指示仪", "型号", "REG-001", "温度", "23 ℃", "湿度", "45 %RH"]];
record.getRange("A2:H2").format = { ...bodyFormat, fill: "#F2F2F2" };
record.mergeCells("A3:A4");
record.mergeCells("B3:B4");
record.mergeCells("C3:D3");
record.mergeCells("E3:E4");
record.mergeCells("F3:F4");
record.mergeCells("G3:G4");
record.mergeCells("H3:H4");
record.getRange("A3:H4").values = [
  ["校准项目", "标准值", "测量值", null, "平均值", "示值误差", "技术要求", "结论"],
  [null, null, "第 1 次", "第 2 次", null, null, null, null],
];
record.getRange("A3:H4").format = headerFormat;
record.getRange("A5:D7").values = [
  ["示值误差", 10, 10.12, 10.08],
  ["示值误差", 20, 20.16, 20.11],
  ["示值误差", 30, 29.82, 29.88],
];
record.getRange("E5:E7").formulas = [["=AVERAGE(C5:D5)"], ["=AVERAGE(C6:D6)"], ["=AVERAGE(C7:D7)"]];
record.getRange("F5:F7").formulas = [["=E5-B5"], ["=E6-B6"], ["=E7-B7"]];
record.mergeCells("G5:G7");
record.getRange("G5").values = [["±0.5"]];
record.getRange("H5:H7").formulas = [
  ['=IF(ABS(F5)<=0.5,"合格","不合格")'],
  ['=IF(ABS(F6)<=0.5,"合格","不合格")'],
  ['=IF(ABS(F7)<=0.5,"合格","不合格")'],
];
record.mergeCells("A9:H9");
record.getRange("A9:H9").values = [["重复性项目"]];
record.getRange("A9:H9").format = sectionFormat;
record.getRange("A10:H10").values = [["校准项目", "标准值", "测量 1", "测量 2", "测量 3", "极差", "技术要求", "结论"]];
record.getRange("A10:H10").format = headerFormat;
record.getRange("A11:E12").values = [
  ["重复性", 50, 50.03, 49.98, 50.01],
  ["重复性", 100, 100.04, 99.97, 100.01],
];
record.getRange("F11:F12").formulas = [["=MAX(C11:E11)-MIN(C11:E11)"], ["=MAX(C12:E12)-MIN(C12:E12)"]];
record.getRange("G11:G12").values = [["≤0.10"], ["≤0.10"]];
record.getRange("H11:H12").formulas = [
  ['=IF(F11<=0.1,"合格","不合格")'],
  ['=IF(F12<=0.1,"合格","不合格")'],
];
record.getRange("A5:H7").format = bodyFormat;
record.getRange("A10:H12").format = bodyFormat;
record.getRange("B5:F7").format.numberFormat = "0.00";
record.getRange("B11:F12").format.numberFormat = "0.00";
record.getRange("H5:H12").conditionalFormats.add("containsText", { text: "不合格", format: { fill: "#F4CCCC", font: { color: "#9C0006", bold: true } } });
record.freezePanes.freezeRows(4);
record.getRange("A:A").format.columnWidth = 18;
record.getRange("B:G").format.columnWidth = 14;
record.getRange("H:H").format.columnWidth = 12;
record.getRange("1:1").format.rowHeight = 30;

formulaBranch.showGridLines = false;
formulaBranch.mergeCells("A1:H1");
formulaBranch.getRange("A1:H1").values = [["相对/绝对误差公式分支样本"]];
formulaBranch.getRange("A1:H1").format = titleFormat;
formulaBranch.getRange("A3:H3").values = [["标准值", "测量 1", "测量 2", "测量 3", "平均值", "误差结果", "技术要求", "结论"]];
formulaBranch.getRange("A3:H3").format = headerFormat;
formulaBranch.getRange("A4:D5").values = [
  [50, 50.10, 50.12, 50.08],
  [200, 200.40, 200.35, 200.45],
];
formulaBranch.getRange("E4:E5").formulas = [["=AVERAGE(B4:D4)"], ["=AVERAGE(B5:D5)"]];
formulaBranch.getRange("F4:F5").formulas = [
  ["=IF(A4<=100,(E4-A4)/A4,E4-A4)"],
  ["=IF(A5<=100,(E5-A5)/A5,E5-A5)"],
];
formulaBranch.getRange("G4:G5").values = [["≤0.5%"], ["±1"]];
formulaBranch.getRange("H4:H5").formulas = [
  ['=IF(A4<=100,IF(ABS(F4)<=0.005,"合格","不合格"),IF(ABS(F4)<=1,"合格","不合格"))'],
  ['=IF(A5<=100,IF(ABS(F5)<=0.005,"合格","不合格"),IF(ABS(F5)<=1,"合格","不合格"))'],
];
formulaBranch.getRange("A4:H5").format = bodyFormat;
formulaBranch.getRange("A4:F5").format.numberFormat = "0.0000";
formulaBranch.getRange("H4:H5").conditionalFormats.add("containsText", { text: "不合格", format: { fill: "#F4CCCC", font: { color: "#9C0006", bold: true } } });
formulaBranch.getRange("A:H").format.columnWidth = 16;
formulaBranch.freezePanes.freezeRows(3);

ledger.showGridLines = false;
ledger.mergeCells("A1:D1");
ledger.getRange("A1:D1").values = [["非校准工作表（自动匹配应跳过）"]];
ledger.getRange("A1:D1").format = titleFormat;
ledger.getRange("A3:D3").values = [["编号", "设备", "状态", "备注"]];
ledger.getRange("A3:D3").format = headerFormat;
ledger.getRange("A4:D6").values = [
  [1, "REG-001", "在用", "仅用于验证无关工作表跳过"],
  [2, "REG-002", "停用", ""],
  [3, "REG-003", "在用", ""],
];
ledger.getRange("A4:D6").format = bodyFormat;
ledger.getRange("A:D").format.columnWidth = 18;
ledger.getRange("D:D").format.columnWidth = 34;

const keyCheck = await workbook.inspect({
  kind: "table",
  range: "原始记录!A1:H12",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 8,
  maxChars: 6000,
});
console.log(keyCheck.ndjson);

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(formulaErrors.ndjson);

for (const sheetName of ["回归清单", "原始记录", "公式分支", "台账"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1.2, format: "png" });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(path.join(previewDir, `${sheetName}.png`), bytes);
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);

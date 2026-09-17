import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "C:/Users/Kaylee/Documents/GitHub/Marbles-site/outputs/01a05e46-e3ae-75a0-a950-8b8e8592ec6d";
const outputPath = `${outputDir}/201179_HAAS_Offset_Sheet.xlsx`;
const previewPath = `${outputDir}/201179_HAAS_Offset_Sheet_preview.png`;

const records = [
  [1, 1, "Rough stem", "", "D31"],
  [1, 2, "Rough nose", "", "D01"],
  [1, 3, "Rough height", "", "D01"],
  [1, 4, "Head width", "", "D21"],
  [1, 5, "Bottom dish", "", "D41"],
  [2, 1, "Finish stem", "", "D42/D43"],
  [2, 2, "Finish top height", "", "D22"],
  [2, 3, "Finish nose", "", "D32"],
  [2, 4, "Finish bottom height", "", "D02"],
  [3, 1, "Finish bottom tip", "", "D03"],
  [4, 1, "Mill narrow corner", "", "D04"],
  [5, 1, "Spot pivot (S.R. + S.L.)", "", ""],
  [5, 2, "Windage marks", "", ""],
  [6, 1, "Drill pivot", "", ""],
  [7, 1, "R/B relief", "", "D07"],
  [7, 2, "Roughout dovetail", "", "D17"],
  [8, 1, "Dish radius", "", "D08"],
  [9, 1, "Top counter bore", "", "D09"],
  [10, 1, "Thru bore", "", "D10"],
  [11, 1, "Bearing hole", "", "D11"],
  [12, 1, "Mill dovetail", "", "D12"],
  [13, 1, "Open", "", ""],
  [14, 1, "Step drill", "", ""],
  [15, 1, "Windage tap", "", ""],
  [16, 1, "Top relief", "", "D16"],
  [16, 2, "Bottom relief", "", "D16"],
  [17, 1, "Top chamfer", "", "D17"],
  [17, 2, "Bottom chamfer", "", "D37"],
  [18, 1, "Ball pull", "", ""],
];

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Offset Sheet");
sheet.showGridLines = false;
sheet.freezePanes.freezeRows(4);

sheet.getRange("A1:E1").merge();
sheet.getRange("A1").values = [["201179 HAAS MACHINES - OFFSET SHEET"]];
sheet.getRange("A2:E2").merge();
sheet.getRange("A2").values = [["Editable transcription of the two-page handwritten source"]];
sheet.getRange("A4:E4").values = [["Tool", "Step", "Operation", "#", "D Offset"]];
sheet.getRange(`A5:E${4 + records.length}`).values = records;

const used = sheet.getRange(`A1:E${4 + records.length}`);
used.format.font = { name: "Aptos", size: 11, color: "#172033" };

sheet.getRange("A1:E1").format = {
  fill: "#17365D",
  font: { name: "Aptos Display", size: 18, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("A1:E1").format.rowHeight = 34;

sheet.getRange("A2:E2").format = {
  fill: "#D9EAF7",
  font: { name: "Aptos", size: 9, italic: true, color: "#36566F" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("A2:E2").format.rowHeight = 22;
sheet.getRange("A3:E3").format.rowHeight = 8;

sheet.getRange("A4:E4").format = {
  fill: "#2F75B5",
  font: { name: "Aptos", size: 11, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#17365D" },
};
sheet.getRange("A4:E4").format.rowHeight = 24;

const dataRange = sheet.getRange(`A5:E${4 + records.length}`);
dataRange.format = {
  verticalAlignment: "center",
  borders: {
    insideHorizontal: { style: "thin", color: "#C9D8E6" },
    insideVertical: { style: "thin", color: "#C9D8E6" },
    top: { style: "medium", color: "#6B7C8E" },
    bottom: { style: "medium", color: "#6B7C8E" },
    left: { style: "medium", color: "#6B7C8E" },
    right: { style: "medium", color: "#6B7C8E" },
  },
};
sheet.getRange(`A5:B${4 + records.length}`).format.horizontalAlignment = "center";
sheet.getRange(`D5:E${4 + records.length}`).format.horizontalAlignment = "center";
sheet.getRange(`C5:C${4 + records.length}`).format.horizontalAlignment = "left";
sheet.getRange(`A5:E${4 + records.length}`).format.rowHeight = 23;

const groupRanges = [];
let start = 0;
for (let i = 1; i <= records.length; i++) {
  if (i === records.length || records[i][0] !== records[start][0]) {
    groupRanges.push({ tool: records[start][0], startRow: 5 + start, endRow: 4 + i });
    start = i;
  }
}

for (const group of groupRanges) {
  const fill = group.tool % 2 === 0 ? "#F3F7FA" : "#FFFFFF";
  sheet.getRange(`A${group.startRow}:E${group.endRow}`).format.fill = fill;
  sheet.getRange(`A${group.startRow}:E${group.startRow}`).format.borders = {
    top: { style: "medium", color: "#7A8FA3" },
  };
  if (group.endRow > group.startRow) {
    sheet.getRange(`A${group.startRow}:A${group.endRow}`).merge();
  }
  sheet.getRange(`A${group.startRow}:A${group.endRow}`).format = {
    fill: group.tool % 2 === 0 ? "#E7EFF6" : "#DCEAF5",
    font: { name: "Aptos", size: 11, bold: true, color: "#17365D" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: "#8EA9C1" },
  };
}

sheet.getRange(`E5:E${4 + records.length}`).format = {
  font: { name: "Aptos", size: 11, bold: true, color: "#17365D" },
  horizontalAlignment: "center",
};
sheet.getRange(`D5:D${4 + records.length}`).format = {
  fill: "#FFF7DD",
  horizontalAlignment: "center",
};

sheet.getRange("A:A").format.columnWidth = 10;
sheet.getRange("B:B").format.columnWidth = 8;
sheet.getRange("C:C").format.columnWidth = 42;
sheet.getRange("D:D").format.columnWidth = 10;
sheet.getRange("E:E").format.columnWidth = 14;
sheet.getRange("A1:E34").format.wrapText = false;

sheet.getRange("D5:D34").dataValidation = {
  rule: { type: "whole", operator: "between", formula1: 0, formula2: 999 },
};

const inspection = await workbook.inspect({
  kind: "table",
  range: "Offset Sheet!A1:E34",
  include: "values,formulas",
  tableMaxRows: 40,
  tableMaxCols: 6,
});
console.log("INSPECTION");
console.log(inspection.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log("ERROR_SCAN");
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "Offset Sheet",
  range: "A1:E34",
  scale: 1.5,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`OUTPUT=${outputPath}`);
console.log(`PREVIEW=${previewPath}`);

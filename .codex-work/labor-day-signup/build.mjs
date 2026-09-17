import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "C:/Users/Kaylee/Documents/GitHub/Marbles-site/outputs/labor-day-signup";
const bannerPath = `${outputDir}/labor-day-banner.png`;
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Labor Day Signup");
sheet.showGridLines = false;

sheet.getRange("A1:C5").format = { fill: "#FFFFFF" };
const bannerBytes = await fs.readFile(bannerPath);
sheet.images.add({
  dataUrl: `data:image/png;base64,${bannerBytes.toString("base64")}`,
  anchor: {
    from: { row: 0, col: 0, colOffsetPx: 60 },
    extent: { widthPx: 500, heightPx: 250 },
  },
});

sheet.getRange("A6:C6").merge();
sheet.getRange("A6").values = [["LABOR DAY COVERAGE SIGNUP"]];
sheet.getRange("A7:C7").merge();
sheet.getRange("A7").values = [["Please print your name and the time you can work."]];

sheet.getRange("A9:C17").values = [
  ["Area", "Person Signing Up", "Time"],
  ["Operator L", "", ""],
  ["Operator R", "", ""],
  ["Resource", "", ""],
  ["Finishing", "", ""],
  ["Setup 1", "", ""],
  ["Setup 2", "", ""],
  ["", "", ""],
  ["", "", ""],
];

sheet.getRange("A19:C19").merge();
sheet.getRange("A19").values = [["Thank you for helping cover Labor Day!"]];

sheet.getRange("A6:C6").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("A7:C7").format = {
  fill: "#D9EAF7",
  font: { color: "#1F2937", italic: true, size: 11 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("A9:C9").format = {
  fill: "#F2C94C",
  font: { bold: true, color: "#1F2937", size: 12 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "medium", color: "#1F4E78" },
};
sheet.getRange("A10:C17").format = {
  font: { color: "#111827", size: 12 },
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#7C8A99" },
};
sheet.getRange("A10:A17").format = {
  fill: "#F4F8FB",
  font: { bold: true, color: "#1F2937", size: 12 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#7C8A99" },
};
sheet.getRange("B10:B17").format = {
  fill: "#FFFFFF",
  font: { color: "#111827", size: 12 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#7C8A99" },
};
sheet.getRange("C10:C17").format = {
  fill: "#FFFFFF",
  font: { color: "#111827", size: 12 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#7C8A99" },
};
sheet.getRange("A19:C19").format = {
  fill: "#EAF2F8",
  font: { bold: true, color: "#1F4E78", size: 11 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: "#B7C9D6" },
};

sheet.getRange("A1:C5").format.rowHeight = 37;
sheet.getRange("A6:C6").format.rowHeight = 34;
sheet.getRange("A7:C7").format.rowHeight = 24;
sheet.getRange("A8:C8").format.rowHeight = 10;
sheet.getRange("A9:C9").format.rowHeight = 28;
sheet.getRange("A10:C17").format.rowHeight = 34;
sheet.getRange("A18:C18").format.rowHeight = 10;
sheet.getRange("A19:C19").format.rowHeight = 25;
sheet.getRange("A:A").format.columnWidth = 22;
sheet.getRange("B:B").format.columnWidth = 38;
sheet.getRange("C:C").format.columnWidth = 18;

sheet.freezePanes.freezeRows(9);

const check = await workbook.inspect({
  kind: "table",
  range: "'Labor Day Signup'!A1:C19",
  include: "values,formulas",
  tableMaxRows: 19,
  tableMaxCols: 3,
});
console.log("INSPECT\n" + check.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log("ERROR_SCAN\n" + errors.ndjson);

const preview = await workbook.render({
  sheetName: "Labor Day Signup",
  range: "A1:C19",
  scale: 2,
  format: "png",
});
await fs.writeFile(`${outputDir}/labor-day-signup-preview.png`, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/Labor-Day-Coverage-Signup.xlsx`);
console.log(`SAVED ${outputDir}/Labor-Day-Coverage-Signup.xlsx`);

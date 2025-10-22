function generateAgentReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("原始資料");
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到『原始資料』工作表！");
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  const rows = data.slice(1); // 去除標題列
  const agentSales = {};

  rows.forEach(row => {
    const rawAmount = String(row[1] ?? "").replace(/[^\d.]/g, "");
    const amount = parseFloat(rawAmount);
    const agent = row[3]?.trim() || "未填";

    if (!isNaN(amount)) {
      agentSales[agent] = (agentSales[agent] || 0) + amount;
    }
  });

  const sheetName = "業務員報表";
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  sheet.appendRow(["業務員", "銷售總額"]);
  Object.entries(agentSales).forEach(([agent, total]) => {
    sheet.appendRow([agent, total]);
  });

  const lastRow = sheet.getLastRow();
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(`A2:B${lastRow}`))
    .setPosition(lastRow + 2, 1, 0, 0)
    .setOption("title", "業務員銷售總額")
    .setOption("legend", { position: "none" })
    .setOption("hAxis", { title: "業務員" })
    .setOption("vAxis", { title: "銷售總額" })
    .build();
  sheet.insertChart(chart);

  SpreadsheetApp.getUi().alert("✅ 業務員報表已完成！");
}

function generateProductReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("原始資料");
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到『原始資料』工作表！");
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  const rows = data.slice(1); // 去除標題列
  const productSales = {};

  rows.forEach(row => {
    const product = row[0]?.trim();
    const rawAmount = String(row[1] ?? "").replace(/[^\d.]/g, "");
    const amount = parseFloat(rawAmount);

    if (!isNaN(amount) && product) {
      productSales[product] = (productSales[product] || 0) + amount;
    }
  });

  const sheetName = "品項報表";
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  sheet.appendRow(["品項", "銷售總額"]);
  Object.entries(productSales).forEach(([product, total]) => {
    sheet.appendRow([product, total]);
  });

  const lastRow = sheet.getLastRow();
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(`A2:B${lastRow}`))
    .setPosition(lastRow + 2, 1, 0, 0)
    .setOption("title", "品項銷售總額")
    .setOption("legend", { position: "none" })
    .setOption("hAxis", { title: "品項" })
    .setOption("vAxis", { title: "銷售總額" })
    .build();
  sheet.insertChart(chart);

  SpreadsheetApp.getUi().alert("✅ 品項報表已完成！");
}

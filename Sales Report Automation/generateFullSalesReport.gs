function generateFullSalesReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // 取得目前的試算表
  const sourceSheet = ss.getSheetByName("原始資料"); // 取得原始資料工作表
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到『原始資料』工作表！");
    return;
  }

  const data = sourceSheet.getDataRange().getValues(); // 讀取整張表格資料
  const rows = data.slice(1); // 去除標題列

  const agentSales = {};   // 儲存業務員銷售總額
  const productSales = {}; // 儲存品項銷售總額

  rows.forEach(row => {
    const product = row[0]?.trim(); // 品項名稱
    const rawAmount = String(row[1] ?? "").replace(/[^\d.]/g, ""); // 清理金額欄位（移除符號）
    const amount = parseFloat(rawAmount); // 轉為數字
    const agent = row[3]?.trim() || "未填"; // 業務員名稱（若空則標示為「未填」）

    if (!isNaN(amount)) {
      agentSales[agent] = (agentSales[agent] || 0) + amount; // 累加業務員銷售額
      if (product) {
        productSales[product] = (productSales[product] || 0) + amount; // 累加品項銷售額
      }
    }
  });

  // 建立業務員報表工作表
  const agentSheetName = "業務員報表";
  let agentSheet = ss.getSheetByName(agentSheetName);
  if (agentSheet) ss.deleteSheet(agentSheet); // 若已存在則刪除重建
  agentSheet = ss.insertSheet(agentSheetName);

  agentSheet.appendRow(["業務員", "銷售總額"]);
  Object.entries(agentSales).forEach(([agent, total]) => {
    agentSheet.appendRow([agent, total]); // 寫入每位業務員的銷售額
  });

  // 插入業務員長條圖
  const agentLastRow = agentSheet.getLastRow();
  const agentChart = agentSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(agentSheet.getRange(`A2:B${agentLastRow}`))
    .setPosition(agentLastRow + 2, 1, 0, 0)
    .setOption("title", "業務員銷售總額")
    .setOption("legend", { position: "none" })
    .setOption("hAxis", { title: "業務員" })
    .setOption("vAxis", { title: "銷售總額" })
    .build();
  agentSheet.insertChart(agentChart);

  // 建立品項報表工作表
  const productSheetName = "品項報表";
  let productSheet = ss.getSheetByName(productSheetName);
  if (productSheet) ss.deleteSheet(productSheet);
  productSheet = ss.insertSheet(productSheetName);

  productSheet.appendRow(["品項", "銷售總額"]);
  Object.entries(productSales).forEach(([product, total]) => {
    productSheet.appendRow([product, total]); // 寫入每個品項的銷售額
  });

  // 插入品項長條圖
  const productLastRow = productSheet.getLastRow();
  const productChart = productSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(productSheet.getRange(`A2:B${productLastRow}`))
    .setPosition(productLastRow + 2, 1, 0, 0)
    .setOption("title", "品項銷售總額")
    .setOption("legend", { position: "none" })
    .setOption("hAxis", { title: "品項" })
    .setOption("vAxis", { title: "銷售總額" })
    .build();
  productSheet.insertChart(productChart);

  SpreadsheetApp.getUi().alert("✅ 全部報表與圖表已完成！");
}

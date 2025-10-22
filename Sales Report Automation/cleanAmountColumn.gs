function cleanAmountColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("原始資料");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到『原始資料』工作表！");
    return;
  }

  const range = sheet.getDataRange();
  const values = range.getValues();

  for (let i = 1; i < values.length; i++) {
    const raw = values[i][1]; // 假設金額在第 2 欄（B 欄）
    const cleaned = String(raw).replace(/[^\d.]/g, ""); // 移除非數字與小數點
    const number = parseFloat(cleaned);
    if (!isNaN(number)) {
      values[i][1] = number; // 將清理後的數字寫回陣列
    }
  }

  range.setValues(values); // 將整批資料寫回試算表
  SpreadsheetApp.getUi().alert("✅ 金額欄位已清理完成！");
}

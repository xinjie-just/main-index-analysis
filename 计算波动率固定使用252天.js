const ExcelJS = require("exceljs");
const path = require("path");

function sampleStddev(arr) {
  const n = arr.length;
  if (n < 2) return 0;
  const mean = arr.reduce((a, b) => a + b, 0) / n;
  const variance = arr.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (n - 1);
  return Math.sqrt(variance);
}

async function addVolatilityWithStyle(inputFilePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFilePath);

  // 红色字体样式（ARGB）
  const redFont = { color: { argb: "FFFF0000" } };

  for (const worksheet of workbook.worksheets) {
    // 获取所有行（从第1行开始）
    const rowCount = worksheet.rowCount;
    if (rowCount < 2) continue;

    // 列定义（Excel 列号，从1开始）
    const DATE_COL = 1; // A
    const CLOSE_COL = 10; // J
    const SAMPLE_COL = 15; // O
    const VOLATILITY_COL = 16; // P（样本数量右侧）

    // 设置表头：波动率
    worksheet.getCell(1, VOLATILITY_COL).value = "波动率";

    // 按年分组：year -> [{ excelRowIndex, dateStr, closePrice }]
    const yearlyData = {};

    for (let r = 2; r <= rowCount; r++) {
      const dateCell = worksheet.getCell(r, DATE_COL).value;
      const closeCell = worksheet.getCell(r, CLOSE_COL).value;

      if (dateCell == null || closeCell == null) continue;

      let dateStr = String(dateCell).trim();
      if (dateStr.length !== 8 || isNaN(dateStr)) continue;

      const year = dateStr.substring(0, 4);
      const closePrice = parseFloat(closeCell);
      if (isNaN(closePrice) || closePrice <= 0) continue;

      if (!yearlyData[year]) yearlyData[year] = [];
      yearlyData[year].push({ rowIndex: r, date: dateStr, close: closePrice });
    }

    // 记录需要显示的行（表头 + 每年最后一个交易日）
    const visibleRows = new Set([1]); // 表头（第1行）

    // 处理每一年
    for (const [year, list] of Object.entries(yearlyData)) {
      if (list.length < 2) continue; // 少于2天不计算

      // 按日期排序
      list.sort((a, b) => a.date.localeCompare(b.date));

      // 提取收盘价序列
      const prices = list.map((item) => item.close);
      const logReturns = [];
      for (let i = 1; i < prices.length; i++) {
        logReturns.push(Math.log(prices[i] / prices[i - 1]));
      }

      if (logReturns.length === 0) continue;

      const tradingDays = 252; // 实际交易日数
      const dailyStd = sampleStddev(logReturns);
      const annualizedVol = dailyStd * Math.sqrt(tradingDays);

      // 最后一个交易日（Excel 行号）
      const lastRowIdx = list[list.length - 1].rowIndex;
      visibleRows.add(lastRowIdx);

      // 写入波动率值（只写这一列）
      worksheet.getCell(lastRowIdx, VOLATILITY_COL).value = annualizedVol;

      // 整行标红（从A到P列）
      for (let col = 1; col <= VOLATILITY_COL; col++) {
        worksheet.getCell(lastRowIdx, col).font = redFont;
      }
    }

    // 隐藏非目标行
    for (let r = 1; r <= rowCount; r++) {
      worksheet.getRow(r).hidden = !visibleRows.has(r);
    }
  }

  // 保存文件
  const dir = path.dirname(inputFilePath);
  const baseName = path.basename(inputFilePath, path.extname(inputFilePath));
  const outputFilePath = path.join(
    dir,
    `${baseName}_波动率_固定使用252天.xlsx`
  );
  await workbook.xlsx.writeFile(outputFilePath);
  console.log(`✅ 成功生成: ${outputFilePath}`);
}

// 使用
const inputPath = "沪深300-中证500-中证1000-中证2000.xlsx";
addVolatilityWithStyle(inputPath).catch((err) => {
  console.error("❌ 错误:", err);
});

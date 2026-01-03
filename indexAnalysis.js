const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// ============== 详细计算公式说明 ==============
console.log("========================================");
console.log("中证指数关键指标计算公式说明:");
console.log("========================================");

// 日波动率 = (最高价 - 最低价) / 开盘价
console.log("1. 日波动率: (最高价 - 最低价) / 开盘价");

// 日波动幅度 = 最高价 - 最低价
console.log("2. 日波动幅度: 最高价 - 最低价");

// 收盘-开盘价差 = 收盘价 - 开盘价
console.log("3. 收盘-开盘价差: 收盘价 - 开盘价");

// 价格区间位置 = (收盘价 - 最低价) / (最高价 - 最低价)
console.log("4. 价格区间位置: (收盘价 - 最低价) / (最高价 - 最低价)");

// 成交量/价格比 = 成交量（万手） / 收盘价
console.log("5. 成交量/价格比: 成交量（万手） / 收盘价");

// 移动平均线 = 过去N日收盘价的平均值
console.log("6. 5日移动平均: 过去5日收盘价的平均值");
console.log("7. 10日移动平均: 过去10日收盘价的平均值");
console.log("8. 20日移动平均: 过去20日收盘价的平均值");
console.log("9. 60日移动平均: 过去60日收盘价的平均值");
console.log("10. 120日移动平均: 过去120日收盘价的平均值");
console.log("11. 250日移动平均: 过去250日收盘价的平均值");

// 移动平均交叉
console.log("12. 5日>10日: 5日移动平均 > 10日移动平均 ? 1 : 0");
console.log("13. 5日<10日: 5日移动平均 < 10日移动平均 ? 1 : 0");

// 年化波动率 = 日收益率标准差 × √250
console.log("14. 年化波动率: 日收益率标准差 × √250");

// 最大回撤 = (历史最高点 - 后续最低点) / 历史最高点
console.log("15. 最大回撤: (历史最高点 - 后续最低点) / 历史最高点");

// 夏普比率 = (年化收益率 - 无风险利率) / 年化波动率
console.log("16. 夏普比率: (年化收益率 - 无风险利率) / 年化波动率");
console.log("   无风险利率假设为3%（年化）");
console.log("========================================\n");

// ============== 表头模糊匹配配置 ==============
const columnMapping = {
  日期: ["日期", "交易日期", "date", "date"],
  指数代码: ["指数代码", "代码", "index_code", "code"],
  指数中文全称: ["指数中文全称", "全称", "index_full_name", "full_name"],
  指数中文简称: ["指数中文简称", "简称", "index_short_name", "short_name"],
  指数英文全称: [
    "指数英文全称",
    "英文全称",
    "index_english_full",
    "english_full",
  ],
  指数英文简称: [
    "指数英文简称",
    "英文简称",
    "index_english_short",
    "english_short",
  ],
  开盘价: ["开盘价", "open", "open_price", "open_price"],
  最高价: ["最高价", "high", "high_price", "high_price"],
  最低价: ["最低价", "low", "low_price", "low_price"],
  收盘价: ["收盘价", "close", "close_price", "close_price"],
  涨跌点数: ["涨跌点数", "change", "change_points", "change_points"],
  "涨跌幅(%)": ["涨跌幅(%)", "change_percent", "percent", "change_percent"],
  "成交量（万手）": [
    "成交量（万手）",
    "volume",
    "volume_million_shou",
    "volume",
  ],
  "成交金额（亿元）": [
    "成交金额（亿元）",
    "amount",
    "amount_billion_yuan",
    "amount",
  ],
};

async function analyzeIndexData() {
  // 1. 读取Excel文件
  const workbook = new ExcelJS.Workbook();
  const inputFileName = "中证A50-中证A100-中证A500.xlsx";

  try {
    await workbook.xlsx.readFile(inputFileName);
    console.log(`成功加载文件: ${inputFileName}`);
  } catch (error) {
    console.error(`文件加载失败: ${error.message}`);
    console.error('请确保文件名为"中证指数数据.xlsx"且位于当前目录');
    return;
  }

  // 2. 遍历每个工作表
  for (const worksheet of workbook.worksheets) {
    console.log(`\n处理工作表: ${worksheet.name}`);

    // 3. 查找匹配的列索引
    const headers = worksheet.getRow(1).values;
    const columnIndices = {};
    const missingColumns = [];

    // 3.1 匹配所有需要的列
    for (const [key, keywords] of Object.entries(columnMapping)) {
      let foundIndex = null;

      for (let i = 0; i < headers.length; i++) {
        const header = headers[i];
        if (!header) continue;

        const headerStr = header.toString().toLowerCase().trim();
        for (const keyword of keywords) {
          if (headerStr.includes(keyword.toLowerCase())) {
            foundIndex = i;
            break;
          }
        }
        if (foundIndex !== null) break;
      }

      if (foundIndex === null) {
        missingColumns.push(key);
      } else {
        columnIndices[key] = foundIndex + 1; // ExcelJS列索引是1-based
      }
    }

    // 3.2 检查缺失列
    if (missingColumns.length > 0) {
      console.error(
        `  错误: 工作表 "${worksheet.name}" 缺少必要列: ${missingColumns.join(
          ", "
        )}`
      );
      console.error(
        `  请确保表头包含以下关键词: ${missingColumns
          .map((col) => columnMapping[col].join(", "))
          .join(", ")}`
      );
      continue; // 跳过当前工作表
    }

    console.log(`  表头验证通过，匹配列索引:`);
    for (const [key, index] of Object.entries(columnIndices)) {
      console.log(`    ${key}: 列索引 ${index}`);
    }

    // 4. 添加新列标题
    const newColumns = [
      { header: "日波动率", key: "daily_volatility", width: 15 },
      { header: "日波动幅度", key: "daily_range", width: 15 },
      { header: "收盘-开盘价差", key: "close_open_diff", width: 15 },
      { header: "价格区间位置", key: "price_range_position", width: 15 },
      { header: "成交量/价格比", key: "volume_price_ratio", width: 15 },
      { header: "5日移动平均", key: "ma5", width: 15 },
      { header: "10日移动平均", key: "ma10", width: 15 },
      { header: "20日移动平均", key: "ma20", width: 15 },
      { header: "60日移动平均", key: "ma60", width: 15 },
      { header: "120日移动平均", key: "ma120", width: 15 },
      { header: "250日移动平均", key: "ma250", width: 15 },
      { header: "5日>10日", key: "ma5_gt_ma10", width: 15 },
      { header: "5日<10日", key: "ma5_lt_ma10", width: 15 },
      { header: "年化波动率", key: "annualized_volatility", width: 15 },
      { header: "最大回撤", key: "max_drawdown", width: 15 },
      { header: "夏普比率", key: "sharpe_ratio", width: 15 },
    ];

    // 添加新列
    newColumns.forEach((col, index) => {
      worksheet.getColumn(15 + index).values = [col.header];
    });

    // 5. 收集所有收盘价用于全局指标计算
    const closePrices = [];
    const rows = worksheet.getRows(2, worksheet.rowCount);

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const closePrice = parseFloat(row.getCell(columnIndices["收盘价"]).value);

      if (!isNaN(closePrice)) {
        closePrices.push(closePrice);
      }
    }

    console.log(`  数据点数量: ${closePrices.length}`);

    // 6. 计算全局指标
    const annualizedVolatility = calculateAnnualizedVolatility(closePrices);
    const maxDrawdown = calculateMaxDrawdown(closePrices);
    const sharpeRatio = calculateSharpeRatio(closePrices);

    console.log(`  年化波动率: ${annualizedVolatility.toFixed(4)}`);
    console.log(`  最大回撤: ${maxDrawdown.toFixed(4)}%`);
    console.log(`  夏普比率: ${sharpeRatio.toFixed(4)}`);

    // 7. 计算所有新列
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowIndex = i + 2; // Excel行号（从2开始）

      // 提取关键数据
      const openPrice = parseFloat(row.getCell(columnIndices["开盘价"]).value);
      const highPrice = parseFloat(row.getCell(columnIndices["最高价"]).value);
      const lowPrice = parseFloat(row.getCell(columnIndices["最低价"]).value);
      const closePrice = parseFloat(row.getCell(columnIndices["收盘价"]).value);
      const volume = parseFloat(
        row.getCell(columnIndices["成交量（万手）"]).value
      );

      // 7.1 基础指标计算
      row.getCell(15).value = (highPrice - lowPrice) / openPrice; // 日波动率
      row.getCell(16).value = highPrice - lowPrice; // 日波动幅度
      row.getCell(17).value = closePrice - openPrice; // 收盘-开盘价差
      row.getCell(18).value =
        highPrice - lowPrice > 0
          ? (closePrice - lowPrice) / (highPrice - lowPrice)
          : 0; // 价格区间位置
      row.getCell(19).value = closePrice > 0 ? volume / closePrice : 0; // 成交量/价格比

      // 7.2 移动平均线计算
      row.getCell(20).value = calculateMA(
        worksheet,
        rowIndex,
        5,
        columnIndices["收盘价"]
      ); // 5日移动平均
      row.getCell(21).value = calculateMA(
        worksheet,
        rowIndex,
        10,
        columnIndices["收盘价"]
      ); // 10日移动平均
      row.getCell(22).value = calculateMA(
        worksheet,
        rowIndex,
        20,
        columnIndices["收盘价"]
      ); // 20日移动平均
      row.getCell(23).value = calculateMA(
        worksheet,
        rowIndex,
        60,
        columnIndices["收盘价"]
      ); // 60日移动平均
      row.getCell(24).value = calculateMA(
        worksheet,
        rowIndex,
        120,
        columnIndices["收盘价"]
      ); // 120日移动平均
      row.getCell(25).value = calculateMA(
        worksheet,
        rowIndex,
        250,
        columnIndices["收盘价"]
      ); // 250日移动平均

      // 7.3 移动平均交叉
      row.getCell(26).value =
        row.getCell(20).value > row.getCell(21).value ? 1 : 0; // 5日>10日
      row.getCell(27).value =
        row.getCell(20).value < row.getCell(21).value ? 1 : 0; // 5日<10日

      // 7.4 添加全局指标（复制到每一行）
      row.getCell(28).value = annualizedVolatility; // 年化波动率
      row.getCell(29).value = maxDrawdown; // 最大回撤
      row.getCell(30).value = sharpeRatio; // 夏普比率
    }
  }

  // 8. 保存结果到新文件
  const outputFileName = `${inputFileName.split(".")[0]}_分析结果_${new Date()
    .toISOString()
    .replace(/[:.]/g, "-")}.xlsx`;
  await workbook.xlsx.writeFile(outputFileName);

  console.log(`\n分析完成！结果已保存到: ${outputFileName}`);
  console.log(`关键指标已添加: ${newColumns.map((c) => c.header).join(", ")}`);
}

// ============== 计算函数 ==============
function calculateAnnualizedVolatility(closePrices) {
  if (closePrices.length < 2) return 0;

  // 计算每日收益率
  const dailyReturns = [];
  for (let i = 1; i < closePrices.length; i++) {
    const returnRate =
      (closePrices[i] - closePrices[i - 1]) / closePrices[i - 1];
    dailyReturns.push(returnRate);
  }

  // 计算收益率标准差
  const mean =
    dailyReturns.reduce((sum, val) => sum + val, 0) / dailyReturns.length;
  const variance =
    dailyReturns.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) /
    (dailyReturns.length - 1);
  const stdDev = Math.sqrt(variance);

  // 年化波动率 = 日标准差 * sqrt(250)
  return stdDev * Math.sqrt(250);
}

function calculateMaxDrawdown(closePrices) {
  if (closePrices.length < 2) return 0;

  let maxPeak = closePrices[0];
  let maxDrawdown = 0;

  for (let i = 1; i < closePrices.length; i++) {
    if (closePrices[i] > maxPeak) {
      maxPeak = closePrices[i];
    } else {
      const drawdown = (maxPeak - closePrices[i]) / maxPeak;
      if (drawdown > maxDrawdown) {
        maxDrawdown = drawdown;
      }
    }
  }

  return maxDrawdown * 100; // 转换为百分比
}

function calculateSharpeRatio(closePrices) {
  if (closePrices.length < 2) return 0;

  // 计算每日收益率
  const dailyReturns = [];
  for (let i = 1; i < closePrices.length; i++) {
    const returnRate =
      (closePrices[i] - closePrices[i - 1]) / closePrices[i - 1];
    dailyReturns.push(returnRate);
  }

  // 计算年化收益率（简单年化）
  const totalReturn =
    (closePrices[closePrices.length - 1] - closePrices[0]) / closePrices[0];
  const annualizedReturn = totalReturn * (250 / (closePrices.length - 1));

  // 无风险利率假设为3%（年化）
  const riskFreeRate = 0.03;

  // 计算年化波动率
  const dailyReturnStd = calculateDailyReturnStd(dailyReturns);
  const annualizedVolatility = dailyReturnStd * Math.sqrt(250);

  // 夏普比率 = (年化收益率 - 无风险利率) / 年化波动率
  return (annualizedReturn - riskFreeRate) / annualizedVolatility;
}

function calculateDailyReturnStd(dailyReturns) {
  if (dailyReturns.length === 0) return 0;

  const mean =
    dailyReturns.reduce((sum, val) => sum + val, 0) / dailyReturns.length;
  const variance =
    dailyReturns.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) /
    (dailyReturns.length - 1);
  return Math.sqrt(variance);
}

function calculateMA(worksheet, currentRow, period, closeColIndex) {
  if (currentRow < period + 1) {
    return null;
  }

  let sum = 0;
  let count = 0;

  for (let i = currentRow - period; i < currentRow; i++) {
    const cell = worksheet.getRow(i).getCell(closeColIndex);
    const value = parseFloat(cell.value);

    if (!isNaN(value)) {
      sum += value;
      count++;
    }
  }

  return count > 0 ? sum / count : null;
}

// 执行分析
analyzeIndexData().catch(console.error);

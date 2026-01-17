const ExcelJS = require("exceljs");
const path = require("path");

// 尝试将值解析为日期
function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;

  // 处理 "YYYY-MM-DD" 格式的字符串
  if (typeof value === "string") {
    const match = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (match) {
      const [, y, m, d] = match;
      const date = new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
      // 验证是否有效日期
      if (
        date.getFullYear() == y &&
        date.getMonth() == m - 1 &&
        date.getDate() == d
      ) {
        return date;
      }
    }
  }
  return null;
}

// 计算样本标准差（无偏估计，除以 n-1）
function sampleStd(values) {
  if (values.length < 2) return null;
  const n = values.length;
  const mean = values.reduce((sum, val) => sum + val, 0) / n;
  const sumSquaredDiff = values.reduce(
    (sum, val) => sum + Math.pow(val - mean, 2),
    0
  );
  const variance = sumSquaredDiff / (n - 1); // 无偏估计
  return Math.sqrt(variance);
}

// 计算年化波动率
function calculateAnnualizedVolatility(dailyPrices) {
  if (dailyPrices.length < 2) return null;

  // 计算对数收益率：r_t = ln(P_t / P_{t-1})
  const logReturns = [];
  for (let i = 1; i < dailyPrices.length; i++) {
    const prev = dailyPrices[i - 1];
    const curr = dailyPrices[i];
    if (prev > 0 && curr > 0) {
      const logReturn = Math.log(curr / prev);
      logReturns.push(logReturn);
    }
  }

  if (logReturns.length < 2) return null;

  const dailyStd = sampleStd(logReturns);
  if (dailyStd === null) return null;

  // 年化波动率 = 日标准差 × √252
  const annualizedVolatility = dailyStd * Math.sqrt(252);
  return annualizedVolatility;
}

// 格式化百分比，保留2位小数；无法计算则返回 "--"
function formatPercentage(value) {
  if (value === null || value === undefined || isNaN(value)) {
    return "--";
  }
  return `${(value * 100).toFixed(2)}%`;
}

async function processExcelFile(filePath) {
  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(filePath);

    for (const worksheet of workbook.worksheets) {
      console.log(`处理工作表: ${worksheet.name}`);

      let headerRow = null;
      const allDataRows = [];

      // 读取所有数据行
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          headerRow = row;
        } else {
          const dateValue = row.getCell(1).value;
          const parsedDate = parseDate(dateValue);

          if (parsedDate) {
            const closePrice = row.getCell(3).value;
            // 只接受数字类型的收盘价，且必须 > 0
            if (typeof closePrice === "number" && closePrice > 0) {
              allDataRows.push({
                row: row,
                date: parsedDate,
                closePrice: closePrice,
              });
            }
          }
        }
      });

      if (allDataRows.length === 0) {
        console.warn(`⚠️ 工作表 ${worksheet.name} 未找到有效数据，跳过`);
        continue;
      }

      // 按日期升序排序
      allDataRows.sort((a, b) => a.date - b.date);

      // 按年分组
      const dataByYear = {};
      for (const item of allDataRows) {
        const year = item.date.getFullYear();
        if (!dataByYear[year]) dataByYear[year] = [];
        dataByYear[year].push(item);
      }

      // 找出每年最后一个交易日
      const yearEndMap = {};
      for (const [yearStr, items] of Object.entries(dataByYear)) {
        const year = parseInt(yearStr);
        items.sort((a, b) => a.date - b.date);
        yearEndMap[year] = items[items.length - 1];
      }

      const years = Object.keys(yearEndMap)
        .map(Number)
        .sort((a, b) => a - b);

      if (years.length === 0) {
        console.warn(`⚠️ 工作表 ${worksheet.name} 无有效年份数据，跳过`);
        continue;
      }

      // 添加“波动率”列到表头
      const volColIndex = worksheet.columnCount + 1;
      if (headerRow) {
        const cell = headerRow.getCell(volColIndex);
        cell.value = "波动率";
        cell.font = { color: { argb: "FFFF0000" }, bold: true };
      }

      // 隐藏所有数据行
      for (const item of allDataRows) {
        item.row.hidden = true;
      }

      // 处理每一年
      for (const year of years) {
        const yearEndItem = yearEndMap[year];

        // 显示该行
        yearEndItem.row.hidden = false;

        // 整行设为红色
        yearEndItem.row.eachCell((cell) => {
          cell.font = { color: { argb: "FFFF0000" } };
        });

        // 计算波动率
        const prices = dataByYear[year].map((i) => i.closePrice);
        const vol = calculateAnnualizedVolatility(prices);

        // 写入波动率值
        const volCell = yearEndItem.row.getCell(volColIndex);
        volCell.value = formatPercentage(vol);
        volCell.font = { color: { argb: "FFFF0000" } };
      }

      // 表头也设为红色
      if (headerRow) {
        headerRow.eachCell((cell) => {
          cell.font = { color: { argb: "FFFF0000" }, bold: true };
        });
      }
    }

    // 保存文件
    const dir = path.dirname(filePath);
    const baseName = path.basename(filePath, path.extname(filePath));
    const outputPath = path.join(dir, `${baseName}_processed.xlsx`);

    await workbook.xlsx.writeFile(outputPath);
    console.log(`✅ 处理成功！输出文件: ${outputPath}`);
  } catch (error) {
    console.error("❌ 处理失败:", error);
  }
}

// ===== 修改这里为你的实际文件路径 =====
const filePath = "创业板50-创业板指-创业200-创业板综.xlsx";
processExcelFile(filePath);

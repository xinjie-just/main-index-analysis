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

// 格式化百分比，保留2位小数
function formatPercentage(value) {
  if (value === null || value === undefined) {
    return "--";
  }
  return `${(value * 100).toFixed(2)}%`;
}

async function processExcelFile(filePath) {
  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(filePath);

    // 定义需要计算的年份跨度
    const yearSpans = [1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21];
    const columnNames = yearSpans.map((span) => `近${span}年年化收益率`);

    for (const worksheet of workbook.worksheets) {
      console.log(`处理工作表: ${worksheet.name}`);

      let headerRow = null;
      const dataRows = [];

      // 读取数据
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          headerRow = row;
        } else {
          const dateValue = row.getCell(1).value;
          const parsedDate = parseDate(dateValue);

          if (parsedDate) {
            const closePrice = row.getCell(3).value;
            const numericClose =
              typeof closePrice === "number" ? closePrice : null;

            if (numericClose !== null) {
              dataRows.push({
                row: row,
                rowNumber: rowNumber,
                date: parsedDate,
                closePrice: numericClose,
              });
            }
          }
        }
      });

      if (dataRows.length === 0) {
        console.warn(`⚠️ 工作表 ${worksheet.name} 未找到有效数据，跳过`);
        continue;
      }

      // 按日期升序排序
      dataRows.sort((a, b) => a.date - b.date);

      // 找出每年最后一个交易日
      const yearEndMap = {};
      for (const item of dataRows) {
        const year = item.date.getFullYear();
        if (!yearEndMap[year] || item.date > yearEndMap[year].date) {
          yearEndMap[year] = item;
        }
      }

      const years = Object.keys(yearEndMap)
        .map(Number)
        .sort((a, b) => a - b);

      if (years.length === 0) {
        console.warn(`⚠️ 工作表 ${worksheet.name} 未找到有效年份数据，跳过`);
        continue;
      }

      const latestYear = years[years.length - 1];

      // 添加年化收益率列到表头
      let currentColumnIndex = worksheet.columnCount + 1;
      const returnColumnIndices = {};

      if (headerRow) {
        for (const columnName of columnNames) {
          const headerCell = headerRow.getCell(currentColumnIndex);
          headerCell.value = columnName;
          headerCell.font = { color: { argb: "FFFF0000" }, bold: true };
          returnColumnIndices[columnName] = currentColumnIndex;
          currentColumnIndex++;
        }
      }

      // 创建第二行用于存放年化收益率（如果不存在的话）
      let summaryRow;
      if (worksheet.rowCount >= 2) {
        summaryRow = worksheet.getRow(2);
      } else {
        summaryRow = worksheet.addRow([]);
      }

      // 计算各年化收益率
      const returnValues = {};
      for (const span of yearSpans) {
        const startYear = latestYear - span;

        // 检查起始年和结束年是否有数据
        if (yearEndMap[startYear] && yearEndMap[latestYear]) {
          const startPrice = yearEndMap[startYear].closePrice;
          const endPrice = yearEndMap[latestYear].closePrice;

          if (startPrice > 0 && endPrice > 0) {
            const annualizedReturn =
              Math.pow(endPrice / startPrice, 1 / span) - 1;
            returnValues[`近${span}年年化收益率`] = annualizedReturn;
          } else {
            returnValues[`近${span}年年化收益率`] = null;
          }
        } else {
          returnValues[`近${span}年年化收益率`] = null;
        }
      }

      // 在第二行设置年化收益率值
      for (const [columnName, columnIndex] of Object.entries(
        returnColumnIndices
      )) {
        const cell = summaryRow.getCell(columnIndex);
        cell.value = formatPercentage(returnValues[columnName]);
        cell.font = { color: { argb: "FFFF0000" }, bold: true };
      }

      // 隐藏所有数据行
      for (const item of dataRows) {
        item.row.hidden = true;
      }

      // 显示每年最后一天的行并设置为红色
      for (const year of years) {
        const yearEndItem = yearEndMap[year];
        yearEndItem.row.hidden = false;

        // 整行设为红色
        yearEndItem.row.eachCell((cell) => {
          cell.font = { color: { argb: "FFFF0000" } };
        });
      }

      // 确保表头为红色
      if (headerRow) {
        headerRow.eachCell((cell) => {
          cell.font = { color: { argb: "FFFF0000" }, bold: true };
        });
      }
    }

    // 保存结果
    const dir = path.dirname(filePath);
    const baseName = path.basename(filePath, path.extname(filePath));
    const outputPath = path.join(dir, `${baseName}_processed.xlsx`);

    await workbook.xlsx.writeFile(outputPath);
    console.log(`✅ 处理成功！输出文件: ${outputPath}`);
  } catch (error) {
    console.error("❌ 处理失败:", error);
  }
}

// 使用方式
const filePath = "创业板50-创业板指-创业200-创业板综.xlsx";
processExcelFile(filePath);

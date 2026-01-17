const ExcelJS = require("exceljs");
const path = require("path");

// 尝试将值解析为日期
function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;

  // 如果是数字（Excel 序列号），也尝试处理（可选）
  if (typeof value === "number") {
    // Excel 日期序列号转 JS Date（假设已正确设置）
    const date = new Date((value - 25569) * 86400 * 1000);
    if (date.getFullYear() > 1900 && date.getFullYear() < 2100) {
      return date;
    }
  }

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

async function processExcelFile(filePath) {
  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(filePath);

    for (const worksheet of workbook.worksheets) {
      console.log(`处理工作表: ${worksheet.name}`);

      let headerRow = null;
      const dataRows = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          headerRow = row;
        } else {
          const dateValue = row.getCell(1).value;
          const parsedDate = parseDate(dateValue);

          if (parsedDate) {
            const closePrice = row.getCell(3).value;
            // 确保收盘价是数字
            const numericClose =
              typeof closePrice === "number" ? closePrice : null;

            dataRows.push({
              row: row,
              rowNumber: rowNumber,
              date: parsedDate,
              closePrice: numericClose,
            });
          }
        }
      });

      if (dataRows.length === 0) {
        console.warn(
          `⚠️ 工作表 ${worksheet.name} 未找到有效日期或数据格式不匹配，跳过`
        );
        continue;
      }

      // 按日期升序排序（确保时间顺序）
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

      // 添加“年收益(%)”列到表头
      const newColIndex = worksheet.columnCount + 1;
      if (headerRow) {
        const headerCell = headerRow.getCell(newColIndex);
        headerCell.value = "年收益(%)";
        headerCell.font = { color: { argb: "FFFF0000" }, bold: true };
      }

      // 隐藏所有数据行
      for (const item of dataRows) {
        item.row.hidden = true;
      }

      // 处理每年最后一天
      for (let i = 0; i < years.length; i++) {
        const year = years[i];
        const yearEndItem = yearEndMap[year];

        // 显示该行
        yearEndItem.row.hidden = false;

        // 整行设为红色
        yearEndItem.row.eachCell((cell, colNumber) => {
          cell.font = { color: { argb: "FFFF0000" } };
        });

        // 计算年收益率（从第二年开始）
        if (i > 0) {
          const prevYearItem = yearEndMap[years[i - 1]];
          const currentClose = yearEndItem.closePrice;
          const prevClose = prevYearItem.closePrice;

          if (currentClose !== null && prevClose !== null && prevClose !== 0) {
            const returnRate = ((currentClose - prevClose) / prevClose) * 100;
            const formattedReturn = returnRate.toFixed(2);

            const returnCell = yearEndItem.row.getCell(newColIndex);
            returnCell.value = `${formattedReturn}%`;
            returnCell.font = { color: { argb: "FFFF0000" } };
          }
        }
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

// ===== 使用方式 =====
const filePath = "创业板50-创业板指-创业200-创业板综.xlsx"; // 👈 修改为你的实际文件路径
processExcelFile(filePath);

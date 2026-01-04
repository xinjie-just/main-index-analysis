const ExcelJS = require("exceljs");
const fs = require("fs");

async function calculateAnnualizedReturns() {
  const inputFileName = "沪深300-中证500-中证1000-中证2000_处理结果.xlsx";
  // 读取输入 Excel 文件
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFileName);

  // 处理每个工作表
  for (const worksheet of workbook.worksheets) {
    // 1. 检查工作表是否有数据
    if (worksheet.rowCount < 2) {
      console.log(`⚠️ 跳过空工作表: ${worksheet.name}`);
      continue;
    }

    // 2. 提取日期列、标记列和收盘价列
    const dateCol = 1; // 日期列 (A列)
    const lastTradeCol = 2; // 标记列 (B列)
    const priceCol = 3; // 收盘价列 (C列)

    // 3. 收集所有最后一个交易日的日期和收盘价
    const yearToDate = {}; // { year: { date: string, price: number, row: number } }

    for (let row = 2; row <= worksheet.rowCount; row++) {
      const dateCell = worksheet.getCell(row, dateCol);
      const lastTradeCell = worksheet.getCell(row, lastTradeCol);
      const priceCell = worksheet.getCell(row, priceCol);

      // 跳过空行
      if (!dateCell.value || !lastTradeCell.value || !priceCell.value) continue;

      // 检查标记列是否为 "Y" (表示最后一个交易日)
      if (lastTradeCell.value.toString().toUpperCase() === "Y") {
        // 提取年份
        let year;
        if (dateCell.value instanceof Date) {
          year = dateCell.value.getFullYear();
        } else if (typeof dateCell.value === "string") {
          year = parseInt(dateCell.value.substring(0, 4));
        } else {
          continue; // 跳过无效日期
        }

        // 保存数据
        yearToDate[year] = {
          date: dateCell.value,
          price: priceCell.value,
          row: row,
        };
      }
    }

    // 4. 检查2025年是否存在（作为基准年）
    if (!yearToDate[2025]) {
      console.log(
        `⚠️ 工作表 ${worksheet.name} 缺少2025年最后一个交易日数据，跳过计算`
      );
      continue;
    }

    // 5. 定义需要计算的年数（3,5,7,...,21）
    const nValues = [3, 5, 7, 9, 11, 13, 15, 17, 19, 21];
    const newColumnTitles = nValues.map((n) => `近${n}年年化收益率`);

    // 6. 添加新列标题
    const newColStart = worksheet.columnCount + 1;
    for (let i = 0; i < newColumnTitles.length; i++) {
      worksheet.getColumn(newColStart + i).values = [newColumnTitles[i]];
    }

    // 7. 计算每个新列的公式
    for (let i = 0; i < nValues.length; i++) {
      const n = nValues[i];
      const startYear = 2025 - n; // 起始年份

      // 检查起始年份是否存在
      if (!yearToDate[startYear] || !yearToDate[2025]) {
        worksheet.getCell(2, newColStart + i).value = "--";
      } else {
        // 获取单元格引用 (C2, C3等)
        const endRow = yearToDate[2025].row;
        const startRow = yearToDate[startYear].row;
        const endCell = `C${endRow}`;
        const startCell = `C${startRow}`;

        // 生成Excel公式: =(C[endRow]/C[startRow])^(1/n)-1
        const formula = `=(${endCell}/${startCell})^(1/${n})-1`;
        worksheet.getCell(2, newColStart + i).value = formula;
      }
    }
  }

  // 8. 保存结果到新文件
  await workbook.xlsx.writeFile(`${inputFileName}_output.xlsx`);
  console.log(`✅ 计算完成！结果已保存到 ${inputFileName}_output.xlsx`);
  console.log("📌 点击单元格可查看公式（Excel会自动计算）");
}

// 执行主逻辑
calculateAnnualizedReturns().catch((err) => {
  console.error("❌ 错误:", err);
  process.exit(1);
});

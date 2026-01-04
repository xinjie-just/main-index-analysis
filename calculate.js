const ExcelJS = require("exceljs");
const fs = require("fs");

async function calculateAnnualizedReturns() {
  // 读取输入 Excel 文件
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(
    "沪深300-中证500-中证1000-中证2000_处理结果.xlsx"
  );

  // 处理每个工作表
  for (const worksheet of workbook.worksheets) {
    console.log(`\n🔍 处理工作表: ${worksheet.name}`);

    // 1. 检查工作表是否有数据
    if (worksheet.rowCount < 2) {
      console.log(`⚠️ 跳过空工作表: ${worksheet.name}`);
      continue;
    }

    // 2. 提取日期列(A列)和收盘价列(J列)
    const dateCol = 1; // 日期列 (A列)
    const priceCol = 10; // 收盘价列 (J列)

    // 3. 收集所有最后一个交易日的日期和收盘价
    const yearToDate = {}; // { year: { date: string, price: number, row: number } }

    for (let row = 2; row <= worksheet.rowCount; row++) {
      const dateCell = worksheet.getCell(row, dateCol);
      const priceCell = worksheet.getCell(row, priceCol);

      // 跳过空行
      if (!dateCell.value || !priceCell.value) continue;

      let year = null;
      let isLastDay = false;

      // 处理日期值（现在只处理Date对象和字符串）
      if (dateCell.value instanceof Date) {
        // Excel 日期对象
        const month = dateCell.value.getMonth(); // 0-11 (11=12月)
        const day = dateCell.value.getDate();
        if (month === 11 && day === 31) {
          // 12月31日
          year = dateCell.value.getFullYear();
          isLastDay = true;
        }
      } else if (typeof dateCell.value === "string") {
        // 字符串日期 (如 "20251231")
        if (dateCell.value.endsWith("1231")) {
          const yearMatch = dateCell.value.match(/(\d{4})/);
          if (yearMatch) {
            year = parseInt(yearMatch[0]);
            isLastDay = true;
          }
        }
      }

      // 识别到最后一个交易日
      if (isLastDay && year) {
        yearToDate[year] = {
          date: dateCell.value,
          price: priceCell.value,
          row: row,
        };
        console.log(
          `✅ 识别到年份: ${year} (行 ${row}, 日期: ${dateCell.value})`
        );
      }
    }

    // 4. 检查2025年是否存在（作为基准年）
    if (!yearToDate[2025]) {
      console.log(`❌ 工作表 ${worksheet.name} 缺少2025年最后一个交易日数据`);
      console.log("📌 请检查以下关键点:");
      console.log("1. 2025年数据行的日期是否为12月31日（12月31日）");
      console.log("2. 2025年数据行的日期格式:");
      console.log("   - 期望: 20251231 或 Excel日期格式（显示为2025-12-31）");
      console.log("   - 实际: ", get2025DateValue(worksheet, dateCol));
      console.log("3. 2025年数据行的J列收盘价是否为数值");
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
        console.log(
          `⚠️ 工作表 ${worksheet.name} 缺少 ${startYear}年数据，近${n}年收益率显示为--`
        );
        worksheet.getCell(2, newColStart + i).value = "--";
      } else {
        // 获取单元格引用 (J2, J3等)
        const endRow = yearToDate[2025].row;
        const startRow = yearToDate[startYear].row;
        const endCell = `J${endRow}`;
        const startCell = `J${startRow}`;

        // 生成Excel公式: =(J[endRow]/J[startRow])^(1/n)-1
        const formula = `=(${endCell}/${startCell})^(1/${n})-1`;
        worksheet.getCell(2, newColStart + i).value = formula;
        console.log(`✅ 工作表 ${worksheet.name} 添加公式: ${formula}`);
      }
    }
  }

  // 8. 保存结果到新文件
  await workbook.xlsx.writeFile("output.xlsx");
  console.log("\n✅ 计算完成！结果已保存到 output.xlsx");
  console.log("📌 点击单元格可查看公式（Excel会自动计算）");
}

// 辅助函数：获取2025年日期列的值（用于诊断）
function get2025DateValue(worksheet, dateCol) {
  for (let row = 2; row <= worksheet.rowCount; row++) {
    const dateCell = worksheet.getCell(row, dateCol);
    if (
      dateCell.value &&
      typeof dateCell.value === "string" &&
      dateCell.value.includes("2025")
    ) {
      return dateCell.value;
    }
  }
  return "未找到2025年数据";
}

// 执行主逻辑
calculateAnnualizedReturns().catch((err) => {
  console.error("❌ 错误:", err);
  process.exit(1);
});

const ExcelJS = require("exceljs");
const path = require("path");
const dayjs = require("dayjs");

async function processExcel(filePath) {
  // 创建工作簿
  const workbook = new ExcelJS.Workbook();

  try {
    // 加载Excel文件
    await workbook.xlsx.readFile(filePath);

    console.log(`找到 ${workbook.worksheets.length} 个工作表`);

    // 遍历所有工作表
    for (const worksheet of workbook.worksheets) {
      console.log(`处理工作表: ${worksheet.name}`);
      await processWorksheet(worksheet);
    }

    // 保存文件
    const timestamp = new Date().getTime();
    const fileName = path.basename(filePath, ".xlsx");
    const outputPath = path.join(
      path.dirname(filePath),
      `${fileName}_processed_${timestamp}.xlsx`
    );
    await workbook.xlsx.writeFile(outputPath);

    console.log("\n处理完成！");
    console.log(`输出文件: ${outputPath}`);
  } catch (error) {
    console.error("处理Excel文件时出错:", error);
  }
}

// 处理单个工作表的函数
async function processWorksheet(worksheet) {
  // 如果工作表没有数据行，跳过
  if (worksheet.rowCount <= 1) {
    console.log(`工作表 ${worksheet.name} 没有数据行，跳过处理`);
    return;
  }

  // 存储每年最后一个交易日的数据
  const yearEndData = new Map(); // year -> {rowNumber, closePrice, date}
  const yearEndRows = new Set(); // 存储每年最后一个交易日行的行号

  // 获取表头行（第一行）
  const headerRow = worksheet.getRow(1);

  // 确保J列存在（添加年收益列）
  const jCol = 10; // J是第10列

  // 检查J列是否已经有标题
  const jCell = headerRow.getCell(jCol);
  if (!jCell.value || jCell.value.toString().trim() === "") {
    jCell.value = "年收益";
    jCell.font = { bold: true };
  }

  // 添加K列显示计算公式
  const kCol = 11; // K是第11列
  const kCell = headerRow.getCell(kCol);
  kCell.value = "计算公式";
  kCell.font = { bold: true };

  // 从第2行开始遍历数据（跳过表头）
  let totalRows = 0;

  console.log(`  解析工作表 ${worksheet.name} 的数据...`);

  // 第一遍：找出每年最后一个交易日
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // 跳过表头

    totalRows = rowNumber; // 更新总行数

    const dateCell = row.getCell("A");
    const dateValue = dateCell.value;

    if (dateValue === null || dateValue === undefined || dateValue === "") {
      return;
    }

    // 将日期转换为dayjs对象 - 支持"2015-01-04"格式
    let date;

    // 情况1: Date对象
    if (dateValue instanceof Date) {
      date = dayjs(dateValue);
    }
    // 情况2: 字符串格式的日期，如"2015-01-04"
    else if (typeof dateValue === "string") {
      // 移除可能的前后空格
      const trimmedDate = dateValue.trim();

      // 尝试多种日期格式
      if (/^\d{4}-\d{2}-\d{2}$/.test(trimmedDate)) {
        // "2015-01-04"
        date = dayjs(trimmedDate, "YYYY-MM-DD");
      } else if (/^\d{8}$/.test(trimmedDate)) {
        // 8位数字字符串，如"20150104"
        date = dayjs(trimmedDate, "YYYYMMDD");
      } else if (/^\d{4}\/\d{2}\/\d{2}$/.test(trimmedDate)) {
        // "2015/01/04"
        date = dayjs(trimmedDate, "YYYY/MM/DD");
      } else {
        // 尝试dayjs自动解析
        date = dayjs(trimmedDate);
      }
    }
    // 情况3: 数字格式的日期（Excel日期序列号）
    else if (typeof dateValue === "number") {
      // Excel日期序列号从1899-12-30开始
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const utcDate = new Date(excelEpoch.getTime() + dateValue * 86400000);
      date = dayjs(utcDate);
    } else {
      return;
    }

    if (!date.isValid()) {
      console.warn(`  第 ${rowNumber} 行日期格式无法解析: ${dateValue}`);
      return;
    }

    const year = date.year();

    // 获取收盘价（C列）
    const closePriceCell = row.getCell("C");
    let closePrice = closePriceCell.value;

    // 如果是字符串，尝试转换为数字
    if (typeof closePrice === "string") {
      closePrice = parseFloat(closePrice);
    }

    // 检查收盘价是否有效
    if (closePrice === null || closePrice === undefined || isNaN(closePrice)) {
      console.warn(`  第 ${rowNumber} 行收盘价无效: ${closePriceCell.value}`);
      return;
    }

    // 更新当年最后交易日数据
    const currentYearData = yearEndData.get(year);

    // 如果当前行日期更晚，则更新
    if (!currentYearData || date.isAfter(currentYearData.date)) {
      yearEndData.set(year, {
        rowNumber: rowNumber,
        closePrice: closePrice,
        date: date,
      });
    }
  });

  // 如果没有找到任何有效数据
  if (yearEndData.size === 0) {
    console.log(`  工作表 ${worksheet.name} 没有找到有效的日期数据，跳过处理`);
    return;
  }

  // 按年份排序
  const sortedYears = Array.from(yearEndData.keys()).sort((a, b) => a - b);
  console.log(
    `  找到 ${yearEndData.size} 年的数据，年份范围: ${sortedYears[0]} 到 ${
      sortedYears[sortedYears.length - 1]
    }`
  );

  // 第二遍：设置样式、隐藏行、计算年收益
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // 跳过表头

    const isLastRow = rowNumber === totalRows;
    let isYearEndRow = false;
    let year = null;

    // 检查是否是某年的最后一个交易日
    for (const [y, data] of yearEndData.entries()) {
      if (data.rowNumber === rowNumber) {
        isYearEndRow = true;
        year = y;
        yearEndRows.add(rowNumber);
        break;
      }
    }

    // 如果不是最后一行且不是年最后交易日，则隐藏
    if (!isLastRow && !isYearEndRow) {
      row.hidden = true;
    }

    // 设置字体颜色：最后一行或年最后交易日行为红色
    if (isLastRow || isYearEndRow) {
      row.eachCell((cell) => {
        // 保留原有的字体样式，只修改颜色
        const currentFont = cell.font || {};
        cell.font = {
          ...currentFont,
          color: { argb: "FFFF0000" },
          bold: currentFont.bold !== undefined ? currentFont.bold : true,
        };
      });
    }

    // 计算年收益（从2011年开始）
    if (isYearEndRow && year >= 2011) {
      const prevYearData = yearEndData.get(year - 1);

      if (
        prevYearData &&
        prevYearData.closePrice !== null &&
        prevYearData.closePrice !== undefined &&
        prevYearData.closePrice !== 0
      ) {
        const currentClosePriceCell = row.getCell("C");
        let currentClosePrice = currentClosePriceCell.value;

        // 如果是字符串，转换为数字
        if (typeof currentClosePrice === "string") {
          currentClosePrice = parseFloat(currentClosePrice);
        }

        if (
          currentClosePrice !== null &&
          currentClosePrice !== undefined &&
          !isNaN(currentClosePrice)
        ) {
          const yearReturn =
            (currentClosePrice - prevYearData.closePrice) /
            prevYearData.closePrice;

          // 将年收益写入J列
          const yearReturnCell = row.getCell(jCol);
          yearReturnCell.value = yearReturn;

          // 格式化为百分比，保留2位小数
          yearReturnCell.numFmt = "0.00%";

          // 在K列显示计算公式
          const formulaCell = row.getCell(kCol);
          const formula = `=(C${rowNumber}-C${prevYearData.rowNumber})/C${prevYearData.rowNumber}`;
          formulaCell.value = formula;

          console.log(
            `  ${year}年收益: ${(yearReturn * 100).toFixed(
              2
            )}% (公式: ${formula})`
          );
        }
      } else if (year >= 2011) {
        console.warn(
          `  ${year}年收益计算失败: 未找到 ${year - 1} 年的数据或收盘价为0`
        );
      }
    }
  });

  // 处理最后一行（如果它还不是某年的最后一个交易日）
  const lastRow = worksheet.getRow(totalRows);
  if (!yearEndRows.has(totalRows)) {
    // 最后一行如果不是年最后交易日，也需要设置红色
    lastRow.eachCell((cell) => {
      const currentFont = cell.font || {};
      cell.font = {
        ...currentFont,
        color: { argb: "FFFF0000" },
        bold: currentFont.bold !== undefined ? currentFont.bold : true,
      };
    });

    // 尝试计算最后一行的年收益
    const lastRowDateCell = lastRow.getCell("A").value;
    if (lastRowDateCell) {
      let lastDate;

      // 使用与上面相同的日期解析逻辑
      if (lastRowDateCell instanceof Date) {
        lastDate = dayjs(lastRowDateCell);
      } else if (typeof lastRowDateCell === "string") {
        const trimmedDate = lastRowDateCell.trim();
        if (/^\d{4}-\d{2}-\d{2}$/.test(trimmedDate)) {
          lastDate = dayjs(trimmedDate, "YYYY-MM-DD");
        } else {
          lastDate = dayjs(trimmedDate);
        }
      } else if (typeof lastRowDateCell === "number") {
        const excelEpoch = new Date(Date.UTC(1899, 11, 30));
        const utcDate = new Date(
          excelEpoch.getTime() + lastRowDateCell * 86400000
        );
        lastDate = dayjs(utcDate);
      }

      if (lastDate && lastDate.isValid()) {
        const lastRowYear = lastDate.year();
        const prevYearData = yearEndData.get(lastRowYear - 1);

        if (
          prevYearData &&
          prevYearData.closePrice !== null &&
          prevYearData.closePrice !== undefined &&
          prevYearData.closePrice !== 0
        ) {
          const lastClosePriceCell = lastRow.getCell("C");
          let lastClosePrice = lastClosePriceCell.value;

          if (typeof lastClosePrice === "string") {
            lastClosePrice = parseFloat(lastClosePrice);
          }

          if (
            lastClosePrice !== null &&
            lastClosePrice !== undefined &&
            !isNaN(lastClosePrice)
          ) {
            const yearReturn =
              (lastClosePrice - prevYearData.closePrice) /
              prevYearData.closePrice;
            const yearReturnCell = lastRow.getCell(jCol);
            yearReturnCell.value = yearReturn;
            yearReturnCell.numFmt = "0.00%";

            // 在K列显示计算公式
            const formulaCell = lastRow.getCell(kCol);
            const formula = `=(C${totalRows}-C${prevYearData.rowNumber})/C${prevYearData.rowNumber}`;
            formulaCell.value = formula;

            console.log(
              `  最后一行(${lastRowYear}年)收益: ${(yearReturn * 100).toFixed(
                2
              )}% (公式: ${formula})`
            );
          }
        }
      }
    }
  }

  // 调整列宽以适应新内容
  worksheet.getColumn(jCol).width = 12;
  worksheet.getColumn(kCol).width = 30;

  console.log(`  工作表 ${worksheet.name} 处理完成`);
}

// 使用示例
const excelFilePath = "./创业板50-创业板指-创业200-创业板综.xlsx"; // 替换为您的Excel文件路径
processExcel(excelFilePath);
// 检查是否提供了文件路径
// if (process.argv.length > 2) {
//   const userFilePath = process.argv[2];
//   processExcel(userFilePath);
// } else if (excelFilePath !== "./创业板50-创业板指-创业200-创业板综.xlsx") {
//   processExcel(excelFilePath);
// } else {
//   console.log("请提供Excel文件路径:");
//   console.log("1. 在代码中修改excelFilePath变量");
//   console.log("2. 或通过命令行参数传递: node script.js 文件路径");
// }

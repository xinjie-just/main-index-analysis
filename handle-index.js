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
    workbook.eachSheet((worksheet, sheetId) => {
      console.log(`处理工作表: ${worksheet.name}`);

      processWorksheet(worksheet);
    });

    // 保存文件
    const outputPath = path.join(
      path.dirname(filePath),
      `processed_${path.basename(filePath)}`
    );
    await workbook.xlsx.writeFile(outputPath);

    console.log("处理完成！");
    console.log(`输出文件: ${outputPath}`);
  } catch (error) {
    console.error("处理Excel文件时出错:", error);
  }
}

// 处理单个工作表的函数
function processWorksheet(worksheet) {
  // 如果工作表没有数据行，跳过
  if (worksheet.rowCount <= 1) {
    console.log(`工作表 ${worksheet.name} 没有数据行，跳过处理`);
    return;
  }

  // 存储每年最后一个交易日的数据 {year: {rowNumber, closePrice}}
  const yearEndData = new Map();
  const yearEndRows = new Set(); // 存储每年最后一个交易日行的行号

  // 获取表头行（第一行）
  const headerRow = worksheet.getRow(1);

  // 确保Q列存在（添加年收益列）
  const QCol = 17; // Q是第17列
  const jCol = 10; // J是第10列

  // 检查Q列是否已经有标题
  const qCell = headerRow.getCell(QCol);
  if (!qCell.value || qCell.value.toString().trim() === "") {
    qCell.value = "年收益";
    // 可以设置表头样式
    qCell.font = { bold: true };
  }

  // 从第2行开始遍历数据（跳过表头）
  let totalRows = 0;

  console.log(`开始解析工作表 ${worksheet.name} 的数据...`);

  // 第一遍：找出每年最后一个交易日
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // 跳过表头

    totalRows = rowNumber; // 更新总行数

    const dateCell = row.getCell("A");
    const dateValue = dateCell.value;

    if (dateValue === null || dateValue === undefined || dateValue === "") {
      return;
    }

    // 将日期转换为dayjs对象
    let date;

    // 情况1: 直接是数字格式的日期，如20141231
    if (typeof dateValue === "number") {
      // 判断是否是像20141231这样的数字（通常大于19000000）
      if (dateValue > 19000000 && dateValue < 25000000) {
        // 将数字转换为字符串，然后解析
        const dateStr = dateValue.toString();
        date = dayjs(dateStr, "YYYYMMDD");
      } else {
        // 可能是Excel日期序列号（如42005代表2014-12-31）
        // Excel的日期序列号从1899-12-30开始
        const excelEpoch = new Date(Date.UTC(1899, 11, 30));
        const utcDate = new Date(excelEpoch.getTime() + dateValue * 86400000);
        date = dayjs(utcDate);
      }
    }
    // 情况2: 字符串格式的日期
    else if (typeof dateValue === "string") {
      // 移除可能的前后空格
      const trimmedDate = dateValue.trim();

      // 尝试多种日期格式
      if (/^\d{8}$/.test(trimmedDate)) {
        // 8位数字字符串，如"20141231"
        date = dayjs(trimmedDate, "YYYYMMDD");
      } else if (/^\d{4}-\d{2}-\d{2}$/.test(trimmedDate)) {
        // "2014-12-31"
        date = dayjs(trimmedDate, "YYYY-MM-DD");
      } else if (/^\d{4}\/\d{2}\/\d{2}$/.test(trimmedDate)) {
        // "2014/12/31"
        date = dayjs(trimmedDate, "YYYY/MM/DD");
      } else {
        // 尝试dayjs自动解析
        date = dayjs(trimmedDate);
      }
    }
    // 情况3: Date对象
    else if (dateValue instanceof Date) {
      date = dayjs(dateValue);
    } else {
      return;
    }

    if (!date.isValid()) {
      return;
    }

    const year = date.year();

    // 获取收盘价
    const closePriceCell = row.getCell("J");
    let closePrice = closePriceCell.value;

    // 如果是字符串，尝试转换为数字
    if (typeof closePrice === "string") {
      closePrice = parseFloat(closePrice);
    }

    // 检查收盘价是否有效
    if (closePrice === null || closePrice === undefined || isNaN(closePrice)) {
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
    console.log(`工作表 ${worksheet.name} 没有找到有效的日期数据，跳过处理`);
    return;
  }

  console.log(`工作表 ${worksheet.name} 找到 ${yearEndData.size} 年的数据`);

  // 按年份排序
  const sortedYears = Array.from(yearEndData.keys()).sort((a, b) => a - b);

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

    // 计算所有年份的年收益（从第二年开始）
    if (isYearEndRow && year > sortedYears[0]) {
      const prevYearData = yearEndData.get(year - 1);

      if (
        prevYearData &&
        prevYearData.closePrice !== null &&
        prevYearData.closePrice !== undefined &&
        prevYearData.closePrice !== 0
      ) {
        const currentClosePriceCell = row.getCell("J");
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

          // 将年收益写入Q列
          const yearReturnCell = row.getCell(QCol);
          yearReturnCell.value = yearReturn;

          // 格式化为百分比，保留2位小数
          yearReturnCell.numFmt = "0.00%";

          // 为单元格添加注释/批注来显示公式（可选）
          // 这里我们设置单元格的note属性来显示计算公式
          // 注意：exceljs的注释功能可能需要更多配置

          // 或者，我们可以直接在单元格值中显示公式
          // 但是Excel公式和值是不同的概念
          // 这里我们采用另一种方法：在相邻列显示公式字符串

          console.log(
            `工作表 ${worksheet.name}, ${year}年收益: ${(
              yearReturn * 100
            ).toFixed(2)}%`
          );
        }
      }
    }
  });

  // 特别处理最后一行（如果它还不是某年的最后一个交易日）
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

    // 尝试计算最后一行的年收益（如果可能）
    const lastRowDateCell = lastRow.getCell("A").value;
    if (lastRowDateCell) {
      let lastDate;

      // 使用与上面相同的日期解析逻辑
      if (typeof lastRowDateCell === "number") {
        if (lastRowDateCell > 19000000 && lastRowDateCell < 25000000) {
          const dateStr = lastRowDateCell.toString();
          lastDate = dayjs(dateStr, "YYYYMMDD");
        } else {
          const excelEpoch = new Date(Date.UTC(1899, 11, 30));
          const utcDate = new Date(
            excelEpoch.getTime() + lastRowDateCell * 86400000
          );
          lastDate = dayjs(utcDate);
        }
      } else if (typeof lastRowDateCell === "string") {
        const trimmedDate = lastRowDateCell.trim();
        if (/^\d{8}$/.test(trimmedDate)) {
          lastDate = dayjs(trimmedDate, "YYYYMMDD");
        } else {
          lastDate = dayjs(trimmedDate);
        }
      } else if (lastRowDateCell instanceof Date) {
        lastDate = dayjs(lastRowDateCell);
      }

      if (lastDate && lastDate.isValid()) {
        const lastRowYear = lastDate.year();
        const prevYearData = yearEndData.get(lastRowYear - 1);

        // 如果最后一行有收盘价且不是当年的最后一个交易日，使用当年的最后一个交易日数据计算
        const currentYearData = yearEndData.get(lastRowYear);

        if (currentYearData && currentYearData.rowNumber !== totalRows) {
          // 如果最后一行不是当年的最后一个交易日，使用当年的最后一个交易日数据
          if (
            prevYearData &&
            prevYearData.closePrice !== null &&
            prevYearData.closePrice !== undefined &&
            prevYearData.closePrice !== 0
          ) {
            if (currentYearData.closePrice !== 0) {
              const yearReturn =
                (currentYearData.closePrice - prevYearData.closePrice) /
                prevYearData.closePrice;
              const yearReturnCell = lastRow.getCell(QCol);
              yearReturnCell.value = yearReturn;
              yearReturnCell.numFmt = "0.00%";
            }
          }
        } else if (
          prevYearData &&
          prevYearData.closePrice !== null &&
          prevYearData.closePrice !== undefined &&
          prevYearData.closePrice !== 0
        ) {
          // 如果最后一行就是当年的最后一个交易日，直接计算
          const lastClosePriceCell = lastRow.getCell("J");
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
            const yearReturnCell = lastRow.getCell(QCol);
            yearReturnCell.value = yearReturn;
            yearReturnCell.numFmt = "0.00%";
          }
        }
      }
    }
  }

  // 在R列添加公式显示（显示计算公式字符串）
  const rCol = 18; // R是第18列
  const rHeaderCell = headerRow.getCell(rCol);
  rHeaderCell.value = "计算公式";
  rHeaderCell.font = { bold: true };

  // 为每个有年收益的单元格添加公式显示
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const yearReturnCell = row.getCell(QCol);
    if (yearReturnCell.value !== null && yearReturnCell.value !== undefined) {
      // 查找该行对应的年份
      let year = null;
      for (const [y, data] of yearEndData.entries()) {
        if (data.rowNumber === rowNumber) {
          year = y;
          break;
        }
      }

      // 如果是最后一行但不是年最后交易日
      if (!year && rowNumber === totalRows) {
        const dateCell = row.getCell("A").value;
        if (dateCell) {
          let date;
          if (typeof dateCell === "number") {
            if (dateCell > 19000000 && dateCell < 25000000) {
              const dateStr = dateCell.toString();
              date = dayjs(dateStr, "YYYYMMDD");
            }
          }
          if (date && date.isValid()) {
            year = date.year();
          }
        }
      }

      if (year && year > sortedYears[0]) {
        const prevYear = year - 1;
        const currentYearData = yearEndData.get(year);
        const prevYearData = yearEndData.get(prevYear);

        if (currentYearData && prevYearData) {
          // 构建公式字符串
          const formula = `=(J${currentYearData.rowNumber}-J${prevYearData.rowNumber})/J${prevYearData.rowNumber}`;

          // 将公式字符串写入R列
          const formulaCell = row.getCell(rCol);
          formulaCell.value = formula;

          // 同时为Q列单元格添加注释（如果支持）
          try {
            // exceljs的note/comment功能
            yearReturnCell.note = {
              texts: [{ text: `计算公式: ${formula}` }],
            };
          } catch (error) {
            // 如果添加注释失败，忽略
          }
        }
      }
    }
  });

  // 调整列宽以适应新内容
  worksheet.getColumn(QCol).width = 12;
  worksheet.getColumn(rCol).width = 25;

  console.log(`工作表 ${worksheet.name} 处理完成`);
}

// 使用示例
const excelFilePath = "./红利低波-1000红利低波-红利低波100-A500红利低波.xlsx"; // 替换为您的Excel文件路径
processExcel(excelFilePath);

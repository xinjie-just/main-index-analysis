const XLSX = require("xlsx");
const XlsxPopulate = require("xlsx-populate");

/**
 * 计算每年的年化指标并处理Excel
 * @param {string} inputFilePath - 输入文件路径
 * @param {string} outputFilePath - 输出文件路径
 */
async function processStockIndexData(inputFilePath, outputFilePath) {
  try {
    // 1. 使用xlsx读取数据
    const workbook = XLSX.readFile(inputFilePath);
    const sheetNames = workbook.SheetNames;

    // 2. 使用xlsx-populate处理样式和写入
    const outputWorkbook = await XlsxPopulate.fromFileAsync(inputFilePath);

    // 处理每个工作表
    for (const sheetName of sheetNames) {
      console.log(`正在处理工作表: ${sheetName}`);

      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // 获取表头行（第一行）
      const headerRow = data[0];

      // 确定列索引
      const dateColIndex = headerRow.indexOf("日期");
      const closeColIndex = headerRow.indexOf("收盘");
      const changeRateColIndex = headerRow.indexOf("涨跌幅(%)");
      const sampleCountColIndex = headerRow.indexOf("样本数量");

      // 在表头添加新列
      headerRow.push("年收益率(%)", "年波动率(%)", "夏普比率");

      // 按年份分组数据，同时存储详细信息
      const yearGroups = {};
      const yearDetails = {}; // 存储每个年份的详细信息

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[dateColIndex]) {
          const dateStr = String(row[dateColIndex]);
          const year = dateStr.substring(0, 4);

          if (!yearGroups[year]) {
            yearGroups[year] = [];
            yearDetails[year] = {
              firstRowIndex: i + 1, // Excel行号（从1开始）
              lastRowIndex: i + 1,
              data: [],
            };
          }

          const tradeData = {
            rowIndex: i,
            excelRowNum: i + 1, // Excel中的行号
            data: row,
            date: dateStr,
            close: parseFloat(row[closeColIndex]) || 0,
            changeRate: parseFloat(row[changeRateColIndex]) || 0,
          };

          yearGroups[year].push(tradeData);
          yearDetails[year].data.push(tradeData);
          yearDetails[year].lastRowIndex = i + 1; // 更新最后一行
        }
      }

      // 获取所有年份并排序
      const years = Object.keys(yearGroups).sort();

      // 计算每年的指标
      const yearMetrics = {};

      for (let i = 0; i < years.length; i++) {
        const year = years[i];
        const yearData = yearGroups[year];
        if (yearData.length === 0) continue;

        // 获取当前年份最后一个交易日
        const lastDay = yearData[yearData.length - 1];
        const firstDay = yearData[0];

        // 计算年收益率
        let yearReturn = 0;
        let yearReturnFormula = "";

        if (i > 0) {
          // 从第二年开始计算
          const prevYear = years[i - 1];
          const prevYearData = yearGroups[prevYear];
          if (prevYearData && prevYearData.length > 0) {
            const prevLastDay = prevYearData[prevYearData.length - 1];

            if (prevLastDay.close > 0 && lastDay.close > 0) {
              yearReturn = (lastDay.close / prevLastDay.close - 1) * 100;

              // 构建公式字符串
              // 收盘价所在列字母
              const closeColLetter = columnIndexToLetter(closeColIndex);
              yearReturnFormula = `=(${closeColLetter}${lastDay.excelRowNum}-${closeColLetter}${prevLastDay.excelRowNum})/${closeColLetter}${prevLastDay.excelRowNum}`;
            }
          }
        }

        // 计算年波动率
        let annualVolatility = 0;
        let volatilityFormula = "";
        if (yearData.length > 1) {
          // 计算日收益率的标准差
          const returns = yearData.map((d) => d.changeRate);
          const mean = returns.reduce((a, b) => a + b, 0) / returns.length;
          const variance =
            returns.reduce((a, b) => a + Math.pow(b - mean, 2), 0) /
            returns.length;
          const dailyStdDev = Math.sqrt(variance);

          // 年化波动率 = 日波动率 × √252（假设每年252个交易日）
          annualVolatility = dailyStdDev * Math.sqrt(252);

          // 构建波动率公式
          const changeRateColLetter = columnIndexToLetter(changeRateColIndex);
          volatilityFormula = `=STDEV.P(${changeRateColLetter}${firstDay.excelRowNum}:${changeRateColLetter}${lastDay.excelRowNum})*SQRT(252)`;
        }

        // 计算夏普比率（假设无风险利率为3%）
        const riskFreeRate = 3.0; // 3% 无风险利率
        let sharpeRatio = 0;
        let sharpeRatioFormula = "";
        if (annualVolatility > 0) {
          sharpeRatio = (yearReturn - riskFreeRate) / annualVolatility;
          sharpeRatioFormula = `=(年收益率单元格地址-3%)/年波动率单元格地址`;
        }

        yearMetrics[year] = {
          yearReturn: parseFloat(yearReturn.toFixed(2)),
          annualVolatility: parseFloat(annualVolatility.toFixed(2)),
          sharpeRatio: parseFloat(sharpeRatio.toFixed(2)),
          lastRowIndex: lastDay.rowIndex,
          excelRowNum: lastDay.excelRowNum,
          formulas: {
            yearReturn: yearReturnFormula,
            annualVolatility: volatilityFormula,
            sharpeRatio: sharpeRatioFormula,
          },
          yearData: yearData, // 存储年份数据，用于后续构建公式
        };
      }

      // 更新数据
      const outputSheet = outputWorkbook.sheet(sheetName);

      // 添加新的表头列
      const lastColIndex = sampleCountColIndex + 1;

      // 写入表头
      for (let j = 0; j < headerRow.length; j++) {
        outputSheet.cell(1, j + 1).value(headerRow[j]);
      }

      // 处理数据行
      for (let i = 1; i < data.length; i++) {
        const rowData = data[i];
        const dateStr = String(rowData[dateColIndex] || "");
        const year = dateStr.substring(0, 4);

        // 检查是否是每年最后一行
        const isLastDayOfYear =
          yearMetrics[year] && yearMetrics[year].lastRowIndex === i;

        if (!isLastDayOfYear) {
          // 隐藏行
          outputSheet.row(i + 1).hidden(true);
        } else {
          // 标记为最后交易日，只设置红色字体（不添加填充颜色）
          const rowRange = outputSheet.range(i + 1, 1, i + 1, lastColIndex + 3);
          rowRange.style("fontColor", "FF0000"); // 红色字体

          // 添加年化指标
          const metrics = yearMetrics[year];
          if (metrics) {
            // 年收益率 - 带单位%
            const yearReturnCell = outputSheet.cell(i + 1, lastColIndex + 1);
            yearReturnCell.value(metrics.yearReturn + "%");
            yearReturnCell.style("numberFormat", "0.00%");

            // 为单元格添加公式（这样点击时会在公式编辑器中显示）
            if (metrics.formulas.yearReturn) {
              try {
                yearReturnCell.formula(metrics.formulas.yearReturn);
              } catch (e) {
                console.log(`无法为${year}年年收益率设置公式: ${e.message}`);
              }
            }

            // 年波动率 - 带单位%
            const annualVolatilityCell = outputSheet.cell(
              i + 1,
              lastColIndex + 2
            );
            annualVolatilityCell.value(metrics.annualVolatility + "%");
            annualVolatilityCell.style("numberFormat", "0.00%");

            // 添加批注（使用.note()方法）
            if (metrics.formulas.annualVolatility) {
              try {
                annualVolatilityCell.note(
                  `公式: ${metrics.formulas.annualVolatility}`
                );
              } catch (e) {
                console.log(`无法为${year}年年波动率添加批注: ${e.message}`);
              }
            }

            // 夏普比率 - 不带单位
            const sharpeRatioCell = outputSheet.cell(i + 1, lastColIndex + 3);
            sharpeRatioCell.value(metrics.sharpeRatio);
            sharpeRatioCell.style("numberFormat", "0.00");

            // 添加批注（使用.note()方法）
            if (metrics.formulas.sharpeRatio) {
              try {
                // 构建实际的夏普比率公式
                const yearReturnCellRef =
                  columnIndexToLetter(lastColIndex) + metrics.excelRowNum;
                const annualVolatilityCellRef =
                  columnIndexToLetter(lastColIndex + 1) + metrics.excelRowNum;
                const sharpeRatioFormulaActual = `=(${yearReturnCellRef}-3%)/${annualVolatilityCellRef}`;

                sharpeRatioCell.note(`公式: ${sharpeRatioFormulaActual}`);
              } catch (e) {
                console.log(`无法为${year}年夏普比率添加批注: ${e.message}`);
              }
            }
          }
        }
      }

      // 调整列宽以适应内容
      outputSheet.column(lastColIndex + 1).width(12);
      outputSheet.column(lastColIndex + 2).width(12);
      outputSheet.column(lastColIndex + 3).width(12);

      console.log(
        `工作表 ${sheetName} 处理完成，共处理 ${years.length} 年数据`
      );
    }

    // 保存文件
    await outputWorkbook.toFileAsync(outputFilePath);
    console.log(`处理完成，结果已保存到: ${outputFilePath}`);
  } catch (error) {
    console.error("处理过程中发生错误:", error);
    console.error("错误堆栈:", error.stack);
  }
}

// 辅助函数：将列索引转换为Excel列字母
function columnIndexToLetter(columnIndex) {
  // Excel列字母从A开始，对应0
  let letter = "";
  let tempIndex = columnIndex;

  while (tempIndex >= 0) {
    letter = String.fromCharCode(65 + (tempIndex % 26)) + letter;
    tempIndex = Math.floor(tempIndex / 26) - 1;
  }

  return letter;
}

/**
 * 计算指定年份的指标（辅助函数）
 * @param {Array} yearData - 某一年份的所有交易日数据
 * @param {number} prevYearClose - 上一年最后一个交易日的收盘价
 * @param {number} riskFreeRate - 无风险利率（%）
 * @returns {Object} 包含年收益率、年波动率和夏普比率的对象
 */
function calculateYearlyMetrics(yearData, prevYearClose, riskFreeRate = 3.0) {
  if (!yearData || yearData.length === 0) {
    return { yearReturn: 0, annualVolatility: 0, sharpeRatio: 0 };
  }

  // 按日期排序（确保顺序正确）
  const sortedData = [...yearData].sort(
    (a, b) => parseInt(a.date) - parseInt(b.date)
  );

  const lastDay = sortedData[sortedData.length - 1];

  // 计算年收益率
  let yearReturn = 0;
  if (prevYearClose > 0 && lastDay.close > 0) {
    yearReturn = (lastDay.close / prevYearClose - 1) * 100;
  }

  // 计算年波动率
  let annualVolatility = 0;
  if (sortedData.length > 1) {
    // 收集所有日收益率
    const returns = sortedData.map((d) => d.changeRate);

    // 计算均值
    const mean = returns.reduce((a, b) => a + b, 0) / returns.length;

    // 计算方差
    const variance =
      returns.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / returns.length;

    // 计算标准差
    const dailyStdDev = Math.sqrt(variance);

    // 年化波动率（假设252个交易日）
    annualVolatility = dailyStdDev * Math.sqrt(252);
  }

  // 计算夏普比率
  let sharpeRatio = 0;
  if (annualVolatility > 0) {
    sharpeRatio = (yearReturn - riskFreeRate) / annualVolatility;
  }

  return {
    yearReturn: parseFloat(yearReturn.toFixed(2)),
    annualVolatility: parseFloat(annualVolatility.toFixed(2)),
    sharpeRatio: parseFloat(sharpeRatio.toFixed(2)),
  };
}

/**
 * 处理单个工作表的简化版本（如果只需要计算指标）
 * @param {Array} data - 工作表数据（二维数组）
 * @returns {Object} 包含按年份分组的指标结果
 */
function calculateAllYearsMetrics(data) {
  const result = {};

  // 假设第一行是表头
  const headerRow = data[0];
  const dateColIndex = headerRow.indexOf("日期");
  const closeColIndex = headerRow.indexOf("收盘");
  const changeRateColIndex = headerRow.indexOf("涨跌幅(%)");

  if (
    dateColIndex === -1 ||
    closeColIndex === -1 ||
    changeRateColIndex === -1
  ) {
    console.error("未找到必要的列");
    return result;
  }

  // 按年份分组
  const yearGroups = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[dateColIndex]) {
      const dateStr = String(row[dateColIndex]);
      const year = dateStr.substring(0, 4);

      if (!yearGroups[year]) {
        yearGroups[year] = [];
      }

      yearGroups[year].push({
        date: dateStr,
        close: parseFloat(row[closeColIndex]) || 0,
        changeRate: parseFloat(row[changeRateColIndex]) || 0,
      });
    }
  }

  // 按年份排序
  const years = Object.keys(yearGroups).sort();

  // 计算每年的指标
  let prevYearClose = null;
  for (const year of years) {
    let metrics;

    if (prevYearClose !== null) {
      metrics = calculateYearlyMetrics(yearGroups[year], prevYearClose);
    } else {
      // 第一年没有上一年数据，年收益率为0
      metrics = {
        yearReturn: 0,
        annualVolatility: calculateYearlyMetrics(yearGroups[year], 0)
          .annualVolatility,
        sharpeRatio: 0,
      };
    }

    result[year] = metrics;

    // 更新上一年收盘价
    const lastDay = yearGroups[year][yearGroups[year].length - 1];
    prevYearClose = lastDay.close;
  }

  return result;
}

// 使用示例
const inputFile =
  "./800能源-800材料-800工业-800可选-800消费-800医药-800金融-800地产-800信息-800通信-800公用.xlsx";
const outputFile =
  "./800能源-800材料-800工业-800可选-800消费-800医药-800金融-800地产-800信息-800通信-800公用_处理结果.xlsx";

// 主函数调用
async function main() {
  console.log("开始处理Excel文件...");
  await processStockIndexData(inputFile, outputFile);

  // 如果需要查看计算的指标，也可以单独计算
  console.log("\n如果需要查看详细的指标计算结果:");

  const workbook = XLSX.readFile(inputFile);
  const sheetName = workbook.SheetNames[0]; // 取第一个工作表
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  const metrics = calculateAllYearsMetrics(data);
  console.log("年化指标计算结果:");
  console.log(JSON.stringify(metrics, null, 2));

  // 示例：计算单个年份的详细指标
  console.log("\n示例：计算2015年详细指标");

  // 按年份分组
  const yearGroups = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateStr = String(row[data[0].indexOf("日期")] || "");
    const year = dateStr.substring(0, 4);

    if (!yearGroups[year]) {
      yearGroups[year] = [];
    }

    yearGroups[year].push({
      date: dateStr,
      close: parseFloat(row[data[0].indexOf("收盘")]) || 0,
      changeRate: parseFloat(row[data[0].indexOf("涨跌幅(%)")]) || 0,
    });
  }

  // 计算2014年和2015年指标
  if (yearGroups["2014"] && yearGroups["2015"]) {
    const lastDay2014 = yearGroups["2014"][yearGroups["2014"].length - 1];
    const metrics2015 = calculateYearlyMetrics(
      yearGroups["2015"],
      lastDay2014.close
    );

    console.log("2015年指标:");
    console.log(`年收益率: ${metrics2015.yearReturn}%`);
    console.log(`年波动率: ${metrics2015.annualVolatility}%`);
    console.log(`夏普比率: ${metrics2015.sharpeRatio}`);
  }
}

// 运行主函数
main().catch(console.error);

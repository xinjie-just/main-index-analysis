const XLSX = require("xlsx");
const path = require("path");

// 要计算的年跨度（奇数年）
const YEAR_SPANS = [1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21];

// 安全解析日期（支持 number 或 string 格式的 YYYYMMDD）
function parseDateInt(dateVal) {
  let s;
  if (typeof dateVal === "number") {
    s = String(Math.floor(dateVal));
  } else if (typeof dateVal === "string") {
    s = dateVal.trim();
  } else {
    return null;
  }
  if (!/^\d{8}$/.test(s)) return null;
  const year = parseInt(s.substring(0, 4), 10);
  const month = parseInt(s.substring(4, 6), 10) - 1;
  const day = parseInt(s.substring(6, 8), 10);
  const d = new Date(year, month, day);
  if (
    d.getFullYear() !== year ||
    d.getMonth() !== month ||
    d.getDate() !== day
  ) {
    return null;
  }
  return d;
}

function processWorksheet(sheetName, worksheet) {
  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  if (range.e.r < 1) {
    console.warn(`工作表 ${sheetName} 行数不足，跳过`);
    return;
  }

  // 读取标题
  const headers = [];
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const cell = worksheet[XLSX.utils.encode_cell({ r: 0, c: C })];
    headers.push(cell && cell.v ? String(cell.v).trim() : "");
  }

  const sampleColIndex = headers.indexOf("样本数量");
  if (sampleColIndex === -1) {
    console.warn(`工作表 ${sheetName} 未找到“样本数量”列，跳过`);
    return;
  }

  // 提取所有有效数据行
  const data = [];
  for (let R = 1; R <= range.e.r; ++R) {
    const dateCell = worksheet[XLSX.utils.encode_cell({ r: R, c: 0 })];
    if (!dateCell || dateCell.v == null) continue;

    const dateObj = parseDateInt(dateCell.v);
    if (!dateObj) continue;

    const row = { dateObj, rowIdx: R };
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: R, c: C })];
      const key = headers[C] || `col${C}`;
      row[key] = cell ? cell.v : null;
    }
    data.push(row);
  }

  if (data.length === 0) {
    console.warn(`工作表 ${sheetName} 无有效数据，跳过`);
    return;
  }

  // 按日期升序排序
  data.sort((a, b) => a.dateObj - b.dateObj);

  // 构建：年份 -> 该年所有交易日（按时间升序）
  const yearToRows = new Map();
  for (const row of data) {
    const year = row.dateObj.getFullYear();
    if (!yearToRows.has(year)) yearToRows.set(year, []);
    yearToRows.get(year).push(row);
  }

  // 获取每个年份的最后一个交易日（按日期最大）
  const yearLastTrade = new Map();
  for (const [year, rows] of yearToRows) {
    rows.sort((a, b) => a.dateObj - b.dateObj);
    yearLastTrade.set(year, rows[rows.length - 1]);
  }

  // 当前最新年份（最后一个交易日所在年）
  const latestRow = data[data.length - 1];
  const currentYear = latestRow.dateObj.getFullYear();
  const P_end = parseFloat(latestRow["收盘"]);
  if (isNaN(P_end) || P_end <= 0) {
    console.warn(`工作表 ${sheetName} 最新收盘价无效`);
    return;
  }

  // 计算各 N 年年化收益率
  const results = {};
  for (const N of YEAR_SPANS) {
    const startYear = currentYear - N;

    // 检查起始年是否存在最后一个交易日
    if (!yearLastTrade.has(startYear)) {
      results[`近${N}年年化收益率`] = "--";
      continue;
    }

    const startRow = yearLastTrade.get(startYear);
    const P_start = parseFloat(startRow["收盘"]);
    if (isNaN(P_start) || P_start <= 0) {
      results[`近${N}年年化收益率`] = "--";
      continue;
    }

    // 年化公式：(P_end / P_start)^(1/N) - 1
    const annualized = Math.pow(P_end / P_start, 1 / N) - 1;
    results[`近${N}年年化收益率`] = annualized;
  }

  // ====== 写入新列 ======
  const resultColumns = YEAR_SPANS.map((N) => `近${N}年年化收益率`);
  const newHeaders = [
    ...headers.slice(0, sampleColIndex + 1),
    ...resultColumns,
  ];

  const newMaxCol = newHeaders.length - 1;

  // 写入标题行（第0行）
  for (let C = 0; C <= newMaxCol; C++) {
    const ref = XLSX.utils.encode_cell({ r: 0, c: C });
    if (C < newHeaders.length) {
      worksheet[ref] = { t: "s", v: newHeaders[C] };
    } else {
      delete worksheet[ref];
    }
  }

  // 在第二行（第一个数据行，R=1）写入结果
  const writeRow = 1;
  for (let i = 0; i < resultColumns.length; i++) {
    const colName = resultColumns[i];
    const value = results[colName];
    const cellRef = XLSX.utils.encode_cell({
      r: writeRow,
      c: sampleColIndex + 1 + i,
    });

    if (typeof value === "number") {
      worksheet[cellRef] = { t: "n", v: value, z: "0.00%" };
    } else {
      worksheet[cellRef] = { t: "s", v: value }; // '--'
    }
  }

  // 更新工作表范围
  const maxRow = Math.max(range.e.r, writeRow);
  worksheet["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: maxRow, c: newMaxCol },
  });

  console.log(
    `✅ 工作表 ${sheetName}：当前年=${currentYear}，已计算 ${resultColumns.length} 项`
  );
}

// 主函数
function main(inputFilePath) {
  const workbook = XLSX.readFile(inputFilePath);
  const sheetNames = workbook.SheetNames;

  for (const sheetName of sheetNames) {
    try {
      processWorksheet(sheetName, workbook.Sheets[sheetName]);
    } catch (err) {
      console.error(`处理工作表 ${sheetName} 出错:`, err);
    }
  }

  const dir = path.dirname(inputFilePath);
  const baseName = path.basename(inputFilePath, path.extname(inputFilePath));
  const outputFilePath = path.join(dir, `${baseName}_近几年年化收益年.xlsx`);

  XLSX.writeFile(workbook, outputFilePath);
  console.log(`\n🎉 输出文件: ${outputFilePath}`);
}

// 执行
const args = process.argv.slice(2);
if (args.length === 0) {
  console.error('请提供 Excel 文件路径，例如:\nnode script.js "指数数据.xlsx"');
  process.exit(1);
}
main(path.resolve(args[0]));

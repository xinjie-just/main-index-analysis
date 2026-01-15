const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

// ======================
// 工具函数
// ======================

function toPercent(value) {
  if (value == null || value === "") return "";
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed.endsWith("%")) return trimmed;
  }
  const num =
    typeof value === "number" ? value : parseFloat(String(value).trim());
  if (isNaN(num) || num < -10 || num > 100) {
    return String(value).trim();
  }
  return (num * 100).toFixed(2) + "%";
}

function excelDateToDateString(excelDate) {
  if (typeof excelDate !== "number" || isNaN(excelDate)) {
    const str = String(excelDate).trim();
    if (!str) return "";
    if (/^\d{4}[/\-]\d{1,2}[/\-]\d{1,2}/.test(str)) {
      const d = new Date(str);
      if (!isNaN(d.getTime())) {
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, "0");
        const day = String(d.getDate()).padStart(2, "0");
        return `${y}-${m}-${day}`;
      }
    }
    return str;
  }
  const date = new Date((excelDate - 25569) * 86400 * 1000);
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

// 新增：生成竖排表格（两列：标签 | 值）
function createVerticalTable(labels, values) {
  if (labels.length === 0) return "无数据\n\n";
  const formattedValues = values.map((v) => toPercent(v));
  let table = "| 年份/周期 | 数值 |\n|----------|------|\n";
  for (let i = 0; i < labels.length; i++) {
    const label = labels[i] || "";
    const val = formattedValues[i] || "";
    table += `| ${label} | ${val} |\n`;
  }
  return `\n\n${table}\n\n`;
}

// ======================
// 主程序
// ======================

const filePath = path.resolve(__dirname, "主要指数介绍.xlsx");
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

const range = xlsx.utils.decode_range(worksheet["!ref"]);
const rows = [];
for (let R = range.s.r; R <= range.e.r; ++R) {
  const row = [];
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const addr = xlsx.utils.encode_cell({ r: R, c: C });
    const cell = worksheet[addr];
    row.push(cell ? cell.v : "");
  }
  rows.push(row);
}

if (rows.length < 3) {
  console.error("❌ 至少需要2行表头+1行数据");
  process.exit(1);
}

const firstHeader = rows[0].map((v) => String(v || "").trim());
const secondHeader = rows[1].map((v) => String(v || "").trim());

const headers = [];
for (let i = 0; i < secondHeader.length; i++) {
  if (secondHeader[i]) {
    headers.push(secondHeader[i]);
  } else {
    const first = firstHeader[i] || "";
    if (
      first &&
      ![
        "指定年份年收益(%)",
        "指定年份年波动率(%)",
        "基日以来近几年年化收益(%)",
      ].includes(first)
    ) {
      headers.push(first);
    } else {
      headers.push("");
    }
  }
}

const RETURN_YEARS = Array.from({ length: 21 }, (_, i) => String(2005 + i)); // 2005-2025
const RECENT_PERIODS = [
  "近1年",
  "近3年",
  "近5年",
  "近7年",
  "近9年",
  "近11年",
  "近13年",
  "近15年",
  "近17年",
  "近19年",
  "近21年",
];

const returnStart = headers.indexOf("2005");
if (returnStart === -1) {
  console.error('❌ 未找到 "2005" 列');
  process.exit(1);
}

const returnEnd = returnStart + RETURN_YEARS.length - 1;
const volStart = returnEnd + 4;
const volEnd = volStart + RETURN_YEARS.length - 1;
const recentStart = volEnd + 2;
const recentEnd = recentStart + RECENT_PERIODS.length - 1;

console.log(`✅ 收益率: ${returnStart}-${returnEnd}`);
console.log(`✅ 波动率: ${volStart}-${volEnd}`);
console.log(`✅ 近几年年化: ${recentStart}-${recentEnd}`);

const mainFields = [
  "指数简称",
  "指数代码",
  "指数名称",
  "样本数量",
  "选样范围",
  "选样指标",
  "计算方式",
  "权重上限",
  "调样周期",
  "基点",
  "基日",
  "发布日期",
  "基日以来全部年份年平均收益(%)",
];

const fieldToCol = new Map();
for (let i = 0; i < headers.length; i++) {
  const h = headers[i];
  if (mainFields.includes(h)) {
    fieldToCol.set(h, i);
  }
}

const outputDir = path.resolve(__dirname, "认识指数");
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

const dateFields = ["基日", "发布日期"];
const avgReturnField = "基日以来全部年份年平均收益(%)";

// 统一开头说明
const introNote = `> - 基础数据来源：[中证指数](https://www.csindex.com.cn/)。  
> - 基日以来全部年份年平均收益(%)、指定年份年收益(%)、指定年份年波动率(%)、基日以来近几年年化收益(%)，是通过每日收盘数据计算得出的。数据截止时间是 2025 年 12 月 31 日。  
> - 市场有风险，投资需谨慎。本文仅作指数知识普及，不构成任何投资建议。  

`;

for (let rowIndex = 2; rowIndex < rows.length; rowIndex++) {
  const row = rows[rowIndex];
  if (!row || row.every((cell) => cell === "" || cell == null)) continue;

  const shortNameCell = fieldToCol.has("指数简称")
    ? row[fieldToCol.get("指数简称")]
    : row[0];
  const indexShortName = String(shortNameCell || "").trim();
  if (!indexShortName) continue;

  let md = introNote; // ← 添加开头说明

  // 写入主字段
  for (const field of mainFields) {
    const colIdx = fieldToCol.get(field);
    let value = colIdx !== undefined && row[colIdx] != null ? row[colIdx] : "";

    if (dateFields.includes(field)) {
      value = excelDateToDateString(value);
    } else if (field === avgReturnField) {
      value = toPercent(value);
    } else {
      value = String(value).trim();
    }

    md += `## ${field}\n\n${value || "无"}\n\n`;
  }

  // 指定年份年收益（竖排）
  md += `## 指定年份年收益(%)\n\n`;
  const returnValues = [];
  for (let i = 0; i < RETURN_YEARS.length; i++) {
    const col = returnStart + i;
    returnValues.push(col < row.length ? row[col] : "");
  }
  md += createVerticalTable(RETURN_YEARS, returnValues);

  // 指定年份年波动率（竖排）
  md += `## 指定年份年波动率(%)\n\n`;
  const volValues = [];
  for (let i = 0; i < RETURN_YEARS.length; i++) {
    const col = volStart + i;
    volValues.push(col < row.length ? row[col] : "");
  }
  md += createVerticalTable(RETURN_YEARS, volValues);

  // 近几年年化收益（竖排）
  md += `## 基日以来近几年年化收益(%)\n\n`;
  const recentValues = [];
  for (let i = 0; i < RECENT_PERIODS.length; i++) {
    const col = recentStart + i;
    recentValues.push(col < row.length ? row[col] : "");
  }
  md += createVerticalTable(RECENT_PERIODS, recentValues);

  // 新增：市场占比 & 行业分布
  md += `## 市场占比\n\n<small>更新日期: 2026-01-12</small>\n\n`;
  md += `## 行业分布\n\n<small>更新日期: 2026-01-12</small>\n\n`;

  // 保存文件
  const fileName = `认识“${indexShortName}”指数.md`;
  const safeName = fileName.replace(/[<>:"/\\|?*]/g, "_");
  fs.writeFileSync(path.join(outputDir, safeName), md, "utf8");
  console.log(`✅ ${safeName}`);
}

console.log(`\n🎉 共生成 ${rows.length - 2} 个文档，保存至 “认识指数” 文件夹`);

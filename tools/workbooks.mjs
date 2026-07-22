#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const COLORS = {
  navy: "#173F5F",
  blue: "#20639B",
  teal: "#3CAEA3",
  pale: "#EAF2F8",
  warning: "#F6D55C",
  risk: "#ED553B",
  white: "#FFFFFF",
  ink: "#243447",
  grid: "#D9E2F3",
};

const INPUT_SHEETS = [
  ["元数据", ["字段", "值"]],
  ["在读学员", ["学员标识", "校区", "班型", "峰值时段", "续费月份", "状态"]],
  ["续费率", ["校区", "班型", "续费月份", "基准续费率"]],
  ["招生计划", ["校区", "班型", "招生月份", "本期目标人数", "往期目标人数", "往期实际人数", "开课月份", "峰值时段"]],
  ["班型规则", ["班型", "目标班容", "最小班容", "最大班容", "每班教师数", "每班教室数", "教室类型"]],
  ["教师供给", ["教师标识", "校区", "班型", "月份", "峰值时段", "并发班级上限"]],
  ["教室供给", ["教室标识", "校区", "教室类型", "月份", "峰值时段", "是否可用"]],
];

const OUTPUT_SHEETS = [
  ["运行摘要", ["指标", "值"]],
  ["校验问题", ["级别", "代码", "表", "行", "字段", "说明"]],
  ["标准化在读", ["学员标识", "校区", "班型", "峰值时段", "续费月份", "状态", "源行"]],
  ["标准化续费", ["校区", "班型", "续费月份", "基准续费率", "源行"]],
  ["标准化招生", ["校区", "班型", "招生月份", "本期目标", "往期目标", "往期实际", "开课月份", "峰值时段", "源行"]],
  ["标准化班型", ["班型", "目标班容", "最小班容", "最大班容", "每班教师数", "每班教室数", "教室类型", "源行"]],
  ["标准化教师", ["教师标识", "校区", "班型", "月份", "峰值时段", "并发班级上限", "源行"]],
  ["标准化教室", ["教室标识", "校区", "教室类型", "月份", "峰值时段", "是否可用", "源行"]],
  ["预测明细", ["情景", "月份", "校区", "班型", "峰值时段", "续费后在读", "新增学员", "预测学员", "计划班级", "平均班额", "开班状态", "教师需求", "教室类型", "教室需求", "证据键"]],
  ["教师缺口", ["情景", "月份", "校区", "班型", "峰值时段", "需求", "供给", "缺口", "证据键"]],
  ["教室缺口", ["情景", "月份", "校区", "教室类型", "峰值时段", "需求", "供给", "缺口", "证据键"]],
  ["管理建议", ["建议ID", "资源", "校区", "班型/教室类型", "峰值时段", "首次缺口", "峰值缺口", "持续月数", "提前期", "建议启动月", "建议", "证据键"]],
  ["假设与版本", ["路径", "值"]],
];

const SCHEDULE_INPUT_SHEETS = [
  ["暑假明细", ["科目(外)", "上课地点(外)", "班级编码", "班级名称（外）", "当前人数(占名额)", "最大人数", "开课日期", "上课时间(外)", "授课教师", "班级课次数", "是否放班", "年级(外)", "轮次", "备注"]],
  ["秋季明细", ["科目(原)", "上课地点(外)", "班级编码", "班级名称（外）", "当前人数(占名额)", "最大人数", "开课日期", "上课时间(外)", "授课教师", "班级课次数", "是否放班", "年级(外)", "轮次", "备注"]],
];

function columnName(number) {
  let value = number;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function styleHeader(range) {
  range.format = {
    fill: COLORS.blue,
    font: { bold: true, color: COLORS.white, name: "Microsoft YaHei", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: COLORS.grid },
    rowHeight: 30,
  };
}

function styleBody(range) {
  range.format = {
    font: { color: COLORS.ink, name: "Microsoft YaHei", size: 10 },
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: COLORS.grid },
    rowHeight: 24,
  };
}

function setupInputSheet(workbook, name, headers, rows = []) {
  const sheet = workbook.worksheets.add(name);
  const endColumn = columnName(headers.length);
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  sheet.getRange(`A1:${endColumn}1`).values = [headers];
  styleHeader(sheet.getRange(`A1:${endColumn}1`));
  if (rows.length > 0) {
    sheet.getRange(`A2:${endColumn}${rows.length + 1}`).values = rows;
    styleBody(sheet.getRange(`A2:${endColumn}${rows.length + 1}`));
  }
  sheet.getRange(`A1:${endColumn}${Math.max(2, rows.length + 1)}`).format.autofitColumns();
  for (let column = 0; column < headers.length; column += 1) {
    const headerLength = headers[column].length;
    sheet.getRangeByIndexes(0, column, Math.max(2, rows.length + 1), 1).format.columnWidth = Math.min(24, Math.max(12, headerLength + 4));
  }
  return sheet;
}

function setupOutputSheet(workbook, name, headers) {
  const sheet = workbook.worksheets.add(name);
  const endColumn = columnName(headers.length);
  sheet.showGridLines = false;
  sheet.getRange(`A1:${endColumn}1`).merge();
  sheet.getRange("A1").values = [[`${name}模板`]];
  sheet.getRange(`A1:${endColumn}1`).format = {
    fill: COLORS.navy,
    font: { bold: true, color: COLORS.white, name: "Microsoft YaHei", size: 16 },
    verticalAlignment: "center",
    rowHeight: 32,
  };
  sheet.getRange(`A3:${endColumn}3`).values = [headers];
  styleHeader(sheet.getRange(`A3:${endColumn}3`));
  sheet.freezePanes.freezeRows(3);
  sheet.getRange(`A1:${endColumn}3`).format.autofitColumns();
  for (let column = 0; column < headers.length; column += 1) {
    sheet.getRangeByIndexes(0, column, 4, 1).format.columnWidth = Math.min(28, Math.max(12, headers[column].length + 5));
  }
  return sheet;
}

async function saveWorkbook(workbook, target) {
  await fs.mkdir(path.dirname(target), { recursive: true });
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  await xlsx.save(target);
}

function monthSequence() {
  const months = [];
  let year = 2026;
  let month = 8;
  for (let index = 0; index < 12; index += 1) {
    months.push(`${year}-${String(month).padStart(2, "0")}`);
    month += 1;
    if (month === 13) {
      year += 1;
      month = 1;
    }
  }
  return months;
}

function syntheticRows() {
  const enrolled = [];
  for (let index = 1; index <= 12; index += 1) {
    enrolled.push([`SYN-A-${String(index).padStart(3, "0")}`, "观音桥校区", "素养A", "周六上午", "2026-10", "在读"]);
  }
  for (let index = 1; index <= 8; index += 1) {
    enrolled.push([`SYN-B-${String(index).padStart(3, "0")}`, "观音桥校区", "素养B", "周日下午", "2026-09", "在读"]);
  }

  const teachers = [];
  const rooms = [];
  for (const month of monthSequence()) {
    teachers.push([`SYN-TA-1-${month}`, "观音桥校区", "素养A", month, "周六上午", 1]);
    teachers.push([`SYN-TA-2-${month}`, "观音桥校区", "素养A", month, "周六上午", 1]);
    teachers.push([`SYN-TB-1-${month}`, "观音桥校区", "素养B", month, "周日下午", 1]);
    teachers.push([`SYN-TB-2-${month}`, "观音桥校区", "素养B", month, "周日下午", 1]);
    rooms.push([`SYN-RA-1-${month}`, "观音桥校区", "标准教室", month, "周六上午", true]);
    rooms.push([`SYN-RA-2-${month}`, "观音桥校区", "标准教室", month, "周六上午", true]);
    rooms.push([`SYN-RB-1-${month}`, "观音桥校区", "活动教室", month, "周日下午", true]);
  }

  return {
    "元数据": [["data_as_of", "2026-08-01"], ["source_system", "合成评测"], ["template_version", "1.0.0"]],
    "在读学员": enrolled,
    "续费率": [["观音桥校区", "素养A", "2026-10", 0.75], ["观音桥校区", "素养B", "2026-09", 0.8]],
    "招生计划": [
      ["观音桥校区", "素养A", "2026-08", 15, 10, 8, "2026-09", "周六上午"],
      ["观音桥校区", "素养B", "2026-09", 10, 10, 10, "2026-10", "周日下午"],
    ],
    "班型规则": [["素养A", 10, 6, 12, 1, 1, "标准教室"], ["素养B", 8, 5, 10, 2, 1, "活动教室"]],
    "教师供给": teachers,
    "教室供给": rooms,
  };
}

function syntheticScheduleRows() {
  return {
    "暑假明细": [
      ["双语 STEAM", "合成校区", "ST3AS260001", "三年级A班", 11, 20, "2026-07-02", "08:30-10:00", "合成教师甲(1001)", 12, "是", "3年级", "一轮", ""],
      ["科学思维", "合成校区", "SW4BS260002", "四年级B班", 15, 20, "2026-07-21", "10：20-12：20", "合成教师乙", 15, "是", "4年级", "二轮", ""],
      ["博学", "合成校区", "BX2AS260003", "二年级A班", 16, 16, "2026-08-11", "13点-15点", "合成教师丙", 12, "是", "2年级", "三轮", ""],
      ["脑力思维", "合成校区", "NL1CS260004", "一年级基础班", 8, 16, "2026-07-02", "16:00-18:00", "合成教师丁", 12, "是", "1年级", "一轮", ""],
      ["脑力思维", "合成校区", "NL1AS260005", "一年级A班", 7, 16, "2026-07-21", "16:00-18:00", "合成教师丁", 12, "是", "1年级", "二轮", ""],
      ["双语 STEAM", "合成校区", "ST4AS260006", "四年级A班", 9, 20, "2026-07-02", "18:20-20:20", "合成教师庚", 12, "是", "4年级", "一轮", "非标准课次测试"],
      ["双语 STEAM", "合成校区", "ST3BS260007", "三年级B班", 5, 20, "2026-07-21", "13:30-15:30", "合成教师戊", 12, "是", "3年级", "二轮", ""],
      ["双语 STEAM", "合成校区", "ST3CS260008", "三年级C班", 12, 20, "2026-07-21", "13:30-15:30", "合成教师戊", 12, "是", "3年级", "二轮", ""],
      ["科学思维", "合成校区", "SW5AS260009", "五年级A班", 10, 20, "2026-08-11", "16:20-18:20", "合成教师己", 15, "是", "5年级", "三轮", ""],
      ["科学思维", "合成校区", "SW5BS260010", "五年级B班", 10, 20, "2026-08-11", "16:20-18:20", "合成教师己", 15, "是", "5年级", "三轮", ""],
      ["科学思维", "合成校区", "SW4BS260011", "四年级A班", 14, 20, "2026-08-11", "08:30-10:30", "合成教师辛", 15, "是", "4年级", "三轮", "班名难度冲突测试"],
      ["双语 STEAM", "合成校区", "ST3AS260012", "S2双语班", 6, 20, "2026-07-02", "10:20-12:20", "合成教师壬", 12, "是", "S2", "一轮", ""],
      ["博学", "合成校区", "BX3AS260013", "三年级补课班", 8, 20, "2026-07-02", "10:20-12:20", "合成教师癸", 12, "是", "3年级", "一轮", ""],
      ["博学", "合成校区", "BX3BS260014", "三年级B班", 8, 20, "2026-07-02", "10:20-12:20", "合成教师癸", 12, "取消", "3年级", "一轮", ""],
      ["双语 STEAM", "其他合成校区", "ST6AS260015", "六年级A班", 18, 20, "2026-07-02", "08:30-10:30", "合成教师外", 15, "是", "6年级", "一轮", ""],
    ],
    "秋季明细": [
      ["双语 STEAM", "合成校区", "ST3AF260101", "三年级A班", 12, 20, "2026-09-05", "每周三18:20-20:20", "合成教师甲", 16, "是", "3年级", "", "动态周三列测试"],
      ["脑力思维", "合成校区", "NL2CF260102", "二年级基础班", 8, 16, "2026-09-05", "周六08:30-10:30", "合成教师辛", 16, "是", "2年级", "", ""],
      ["脑力思维", "合成校区", "NL2AF260103", "二年级A班", 7, 16, "2026-09-05", "周日10:20-12:20", "合成教师辛", 16, "是", "2年级", "", ""],
      ["科学思维", "合成校区", "SW5BF260104", "五年级B班", 20, 20, "2026-09-05", "星期日10:20-12:20", "合成教师乙", 16, "是", "5年级", "", ""],
      ["双语 STEAM", "合成校区", "ST1AF260105", "一年级A班", 19, 18, "2026-09-05", "周六13:30-15:30", "合成教师丙", 16, "是", "1年级", "", "超售测试"],
      ["博学", "合成校区", "BX3AF260106", "三年级A班", 9, 20, "2026-09-05", "周四18点-20点", "合成教师丁", 16, "是", "3年级", "", ""],
      ["双语 STEAM", "合成校区", "ST4BF260107", "四年级B班", 13, 20, "2026-09-05", "周六16:20-18:20", "", 16, "是", "4年级", "", "缺教师测试"],
      ["博学", "合成校区", "BX4AF260108", "四年级A班", 11, 20, "2026-09-05", "周六18:20-20:20", "合成教师戊", 16, "否", "4年级", "", ""],
      ["双语 STEAM", "合成校区", "ST3BF260109", "S3双语班", 9, 20, "2026-09-05", "周日13:30-15:30", "合成教师己", 16, "是", "S3", "", ""],
    ],
  };
}

function setupManualScheduleSheet(workbook, name, headers, rows) {
  const sheet = setupInputSheet(workbook, name, headers, rows.map((row) => [row[0], row[1], "", ...row.slice(3)]));
  for (let index = 0; index < rows.length; index += 1) {
    const rowNumber = index + 2;
    sheet.getRange(`C${rowNumber}`).formulas = [[`=COUNTA(D${rowNumber}:${columnName(headers.length)}${rowNumber})`]];
  }
  return sheet;
}

async function build(root) {
  const skillAssets = path.join(root, ".agents/skills/xdf-plan-campus-capacity/assets");
  const fixtures = path.join(root, "evals/fixtures");

  const input = Workbook.create();
  for (const [name, headers] of INPUT_SHEETS) setupInputSheet(input, name, headers);
  const inputMetadata = input.worksheets.getItem("元数据");
  inputMetadata.getRange("A2:B4").values = [["data_as_of", "YYYY-MM-DD"], ["source_system", "来源系统名称"], ["template_version", "1.0.0"]];
  styleBody(inputMetadata.getRange("A2:B4"));
  await saveWorkbook(input, path.join(skillAssets, "capacity-input-template.xlsx"));

  const output = Workbook.create();
  for (const [name, headers] of OUTPUT_SHEETS) setupOutputSheet(output, name, headers);
  await saveWorkbook(output, path.join(skillAssets, "capacity-output-template.xlsx"));

  const synthetic = Workbook.create();
  const rows = syntheticRows();
  for (const [name, headers] of INPUT_SHEETS) setupInputSheet(synthetic, name, headers, rows[name]);
  synthetic.worksheets.getItem("续费率").getRange("D2:D3").setNumberFormat("0.0%");
  await saveWorkbook(synthetic, path.join(fixtures, "synthetic-capacity-input.xlsx"));

  await buildScheduleFixtures(root);
}

async function buildScheduleFixtures(root) {
  const fixtures = path.join(root, "evals/fixtures");
  const scheduleInput = Workbook.create();
  const scheduleRows = syntheticScheduleRows();
  for (const [name, headers] of SCHEDULE_INPUT_SHEETS) setupInputSheet(scheduleInput, name, headers, scheduleRows[name]);
  await saveWorkbook(scheduleInput, path.join(fixtures, "synthetic-schedule-input.xlsx"));

  const manual = Workbook.create();
  setupManualScheduleSheet(
    manual,
    "暑假课表",
    ["科目", "教师", "班量", "一轮08点", "二轮10点", "三轮13点"],
    [
      ["双语", "合成计划教师甲", "", "新班ST3A", "", ""],
      ["思维", "合成计划教师乙", "", "新班NL1", "新班NL2A", "新班NL1B"],
      ["双语", "合成计划教师丙", "", "3人\nST6AS260099", "", ""],
      ["博学", "", "", "", "新班BX3A", ""],
      ["博学", "合成计划教师丁", "", "", "", "新班ST3A"],
    ],
  );
  setupManualScheduleSheet(
    manual,
    "秋季课表",
    ["科目", "教师", "班量", "周三18点", "周六08点", "周日10点"],
    [["思维", "合成计划教师戊", "", "新班SW3B", "", ""]],
  );
  await saveWorkbook(manual, path.join(fixtures, "synthetic-manual-schedule.xlsx"));
}

async function verify(workbookPath, renderDir) {
  const blob = await FileBlob.load(workbookPath);
  const workbook = await SpreadsheetFile.importXlsx(blob);
  const sheetInfo = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 12000 });
  const formulas = await workbook.inspect({ kind: "formula", maxChars: 12000, options: { maxResults: 200 } });
  const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 200 }, maxChars: 12000 });
  await fs.mkdir(renderDir, { recursive: true });
  const names = workbook.worksheets.items.map((sheet) => sheet.name);
  for (const name of names) {
    const image = await workbook.render({ sheetName: name, autoCrop: "all", scale: 1, format: "png" });
    const bytes = new Uint8Array(await image.arrayBuffer());
    const safeName = name.replaceAll(/[\\/:*?"<>|]/g, "_");
    await fs.writeFile(path.join(renderDir, `${safeName}.png`), bytes);
  }
  console.log(JSON.stringify({ workbookPath, sheets: names, sheetInfo: sheetInfo.ndjson, formulas: formulas.ndjson, errors: errors.ndjson }, null, 2));
}

const [command, first, second] = process.argv.slice(2);
if (command === "build" && first) {
  await build(path.resolve(first));
} else if (command === "build-schedule" && first) {
  await buildScheduleFixtures(path.resolve(first));
} else if (command === "verify" && first && second) {
  await verify(path.resolve(first), path.resolve(second));
} else {
  console.error("用法: workbooks.mjs build <repo-root> | build-schedule <repo-root> | verify <workbook.xlsx> <render-dir>");
  process.exit(2);
}

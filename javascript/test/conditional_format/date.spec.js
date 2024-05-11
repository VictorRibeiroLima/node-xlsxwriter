// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const {
  Workbook,
  ConditionalFormatDate,
  Color,
  Format,
} = require('../../src/index');
const fs = require('fs');
const findRootDir = require('../util');
const rootPath = findRootDir(__dirname);
const path = rootPath + '/temp/conditional_format';

test('save to file with conditional format ("ConditionalFormatDate")', async (t) => {
  const thisMoth = new Date();
  const nextMonth = new Date();
  const lastMonth = new Date();
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  lastMonth.setMonth(lastMonth.getMonth() - 1);
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const red = new Color({
    red: 255,
    green: 102,
    blue: 102,
  });

  const green = new Color({
    red: 102,
    green: 255,
    blue: 102,
  });
  const blue = new Color({
    red: 102,
    green: 102,
    blue: 255,
  });

  const redFormat = new Format({
    backgroundColor: red,
    numFmt: 'dd/mm/yyyy',
  });

  const greenFormat = new Format({
    backgroundColor: green,
    numFmt: 'dd/mm/yyyy',
  });

  const blueFormat = new Format({
    backgroundColor: blue,
    numFmt: 'dd/mm/yyyy',
  });

  const lastMonthF = new ConditionalFormatDate({
    rule: 'lastMonth',
    format: redFormat,
  });

  const thisMonthF = new ConditionalFormatDate({
    rule: 'thisMonth',
    format: greenFormat,
  });

  const nextMonthF = new ConditionalFormatDate({
    rule: 'nextMonth',
    format: blueFormat,
  });

  // Write conditional format over the same range.
  sheet.addConditionalFormat({
    firstRow: 0,
    lastRow: 10,
    firstColumn: 0,
    lastColumn: 0,
    format: lastMonthF,
  });

  sheet.addConditionalFormat({
    firstRow: 0,
    lastRow: 10,
    firstColumn: 0,
    lastColumn: 0,
    format: thisMonthF,
  });

  sheet.addConditionalFormat({
    firstRow: 0,
    lastRow: 10,
    firstColumn: 0,
    lastColumn: 0,
    format: nextMonthF,
  });

  for (let row = 0; row <= 10; row++) {
    if (row % 2 === 0) {
      sheet.writeDate(row, 0, lastMonth);
      continue;
    }
    if (row % 3 === 0) {
      sheet.writeDate(row, 0, thisMoth);
      continue;
    }

    sheet.writeDate(row, 0, nextMonth);
  }

  await workbook.saveToFile(`${path}/save-to-file-with-format-date.xlsx`);

  assert.ok(fs.existsSync(`${path}/save-to-file-with-format-date.xlsx`));
});

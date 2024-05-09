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

test('save to file with conditional format ("ConditionalFormatDate")', async (t) => {
  const thisMoth = new Date();
  const nextMonth = new Date();
  const lastMonth = new Date();
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  lastMonth.setMonth(lastMonth.getMonth() - 1);
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const red = new Color(255, 102, 102);
  const green = new Color(102, 255, 102);
  const blue = new Color(102, 102, 255);

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

  for (let i = 0; i <= 10; i++) {
    if (i % 2 === 0) {
      sheet.writeDate(0, i, lastMonth);
      continue;
    }
    if (i % 3 === 0) {
      sheet.writeDate(0, i, thisMoth);
      continue;
    }

    sheet.writeDate(0, i, nextMonth);
  }

  await workbook.saveToFile(
    './temp/conditional_format/save-to-file-with-format-date.xlsx',
  );

  assert.ok(
    fs.existsSync(
      './temp/conditional_format/save-to-file-with-format-date.xlsx',
    ),
  );
});

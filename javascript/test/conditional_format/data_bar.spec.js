// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const {
  Workbook,
  ConditionalFormatDataBar,
  Color,
} = require('../../src/index');
const fs = require('fs');

test('save to file with conditional format ("ConditionalFormatDataBar")', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const lightGreen = new Color({
    red: 204,
    green: 255,
    blue: 204,
  });

  // Write a standard Excel data bar. The conditional format is applied over
  // the full range of values from minimum to maximum.
  const conditionalFormat = new ConditionalFormatDataBar();

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 1,
    format: conditionalFormat,
  });

  // Write a data bar a user defined range. Values <= 3 are shown with zero
  // bar width while values >= 7 are shown with the maximum bar width.
  // with a light green fill color.
  const minMaxConditionalFormat = new ConditionalFormatDataBar({
    minRule: {
      type: 'number',
      value: 3,
    },
    maxRule: {
      type: 'number',
      value: 7,
    },
    fillColor: lightGreen,
  });

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 3,
    lastColumn: 3,
    format: minMaxConditionalFormat,
  });

  for (let i = 2; i <= 11; i++) {
    sheet.writeNumber(1, i, i);
    sheet.writeNumber(3, i, i);
  }

  await workbook.saveToFile(
    './temp/conditional_format/save-to-file-with-format-data-bar.xlsx',
  );

  assert.ok(
    fs.existsSync(
      './temp/conditional_format/save-to-file-with-format-data-bar.xlsx',
    ),
  );
});

// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const findRootDir = require('../util');
const rootPath = findRootDir(__dirname);
const path = rootPath + '/temp/conditional_format';
const { Workbook, ConditionalFormatTop, Format } = require('../../src/index');
const fs = require('fs');

test('save to file with conditional format ("ConditionalFormatTop")', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const lightGreen = new Format({
    backgroundColor: { blue: 102, green: 255, red: 102 },
  });

  const lightRed = new Format({
    backgroundColor: { blue: 102, green: 102, red: 255 },
  });

  const top = new ConditionalFormatTop({
    rule: { type: 'top', value: 10 },
    format: lightGreen,
  });

  const bottom = new ConditionalFormatTop({
    rule: { type: 'bottom', value: 10 },
    format: lightRed,
  });

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 2,
    format: top,
  });

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 2,
    format: bottom,
  });

  for (let row = 2; row <= 11; row++) {
    for (let col = 1; col <= 2; col++) {
      sheet.writeNumber(row, col, Math.random() * 100);
    }
  }

  const filePath = `${path}/conditional_format_top.xlsx`;
  await workbook.saveToFile(filePath);
  assert.ok(fs.existsSync(filePath), 'file has been saved');
});

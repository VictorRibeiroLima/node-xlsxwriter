// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const findRootDir = require('../util');
const rootPath = findRootDir(__dirname);
const path = rootPath + '/temp/conditional_format';
const { Workbook, ConditionalFormatText, Format } = require('../../src/index');
const fs = require('fs');

test('save to file with conditional format ("ConditionalFormatText")', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const lightGreen = new Format({
    backgroundColor: { blue: 102, green: 255, red: 102 },
  });

  const textFormat = new ConditionalFormatText({
    rule: { type: 'beginsWith', value: 'Hello' },
    format: lightGreen,
  });

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 1,
    format: textFormat,
  });

  for (let row = 2; row <= 11; row++) {
    if ((row & 1) === 0) {
      sheet.writeString(row, 1, 'Hello World');
    } else {
      sheet.writeString(row, 1, 'Something else');
    }
  }

  const filePath = `${path}/conditional_format_text.xlsx`;
  await workbook.saveToFile(filePath);
  assert.ok(fs.existsSync(filePath), 'file has been saved');
});

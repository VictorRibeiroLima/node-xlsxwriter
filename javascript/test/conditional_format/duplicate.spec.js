// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const {
  Workbook,
  ConditionalFormatDuplicate,
  Color,
  Format,
} = require('../../src/index');
const fs = require('fs');

test('save to file with conditional format ("ConditionalFormatDuplicate")', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const lightRed = new Color({
    red: 255,
    green: 102,
    blue: 102,
  });

  const lightGreen = new Color({
    red: 102,
    green: 255,
    blue: 102,
  });

  const duplicated = new Format({
    backgroundColor: lightRed,
  });

  const unique = new Format({
    backgroundColor: lightGreen,
  });

  const duplicateFormat = new ConditionalFormatDuplicate({
    format: duplicated,
  });

  const uniqueFormat = new ConditionalFormatDuplicate({
    format: unique,
    invert: true,
  });

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 10,
    format: duplicateFormat,
  });

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 10,
    format: uniqueFormat,
  });

  for (let i = 2; i <= 11; i++) {
    for (let j = 1; j <= 10; j++) {
      if (i % 2 === 0) {
        sheet.writeString(j, i, `${i}-${j}`);
      } else {
        sheet.writeString(j, i, 'duplicated');
      }
    }
  }

  const fileName =
    './temp/conditional_format/save-to-file-with-format-duplicate.xlsx';

  await workbook.saveToFile(fileName);

  assert.ok(fs.existsSync(fileName), 'File has been saved to the file system');
});

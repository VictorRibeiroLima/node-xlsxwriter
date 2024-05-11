// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const {
  Workbook,
  ConditionalFormatFormula,
  Color,
  Format,
  Formula,
} = require('../../src/index');
const fs = require('fs');
const findRootDir = require('../util');
const rootPath = findRootDir(__dirname);
const path = rootPath + '/temp/conditional_format';

test('save to file with conditional format ("ConditionalFormatFormula")', async (t) => {
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

  const redFormat = new Format({
    backgroundColor: lightRed,
  });

  const greenFormat = new Format({
    backgroundColor: lightGreen,
  });

  const isOdd = new Formula({
    formula: '=ISODD(B3)',
  });

  const isEven = new Formula({
    formula: '=ISEVEN(B3)',
  });

  const oddFormat = new ConditionalFormatFormula({
    formula: isOdd,
    format: redFormat,
  });

  const evenFormat = new ConditionalFormatFormula({
    formula: isEven,
    format: greenFormat,
  });

  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 10,
    format: oddFormat,
  });

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 10,
    format: evenFormat,
  });

  for (let row = 2; row <= 11; row++) {
    for (let col = 1; col <= 10; col++) {
      sheet.writeNumber(row, col, row);
    }
  }

  await workbook.saveToFile(`${path}/save-to-file-with-format-formula.xlsx`);

  assert.ok(fs.existsSync(`${path}/save-to-file-with-format-formula.xlsx`));
});

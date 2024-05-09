// @ts-check
const { Workbook, Format, Formula, ConditionalFormatCell } = require('../src');
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');

test('save to file with formula', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  //Write a formula
  const formula = new Formula({
    formula: '=10*B1 + C1',
  });

  //The result of the formula will be 10*5 + 2 = 52
  sheet.writeFormula(0, 0, formula);
  sheet.writeNumber(1, 0, 5);
  sheet.writeNumber(2, 0, 2);

  const fileName = './temp/save-to-file-with-formula.xlsx';

  await workbook.saveToFile(fileName);
  assert.ok(fs.existsSync(fileName));
});

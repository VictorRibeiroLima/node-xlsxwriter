// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const {
  Workbook,
  ConditionalFormatDuplicate,
  Color,
  Format,
} = require('../../src/index');

const workbook = new Workbook();
const sheet = workbook.addSheet();

const lightRed = new Color(255, 102, 102);
const lightGreen = new Color(102, 255, 102);

const duplicated = new Format({
  backgroundColor: lightRed,
});

const unique = new Format({
  backgroundColor: lightGreen,
});

//This will paint red the duplicated rows
const duplicateFormat = new ConditionalFormatDuplicate({
  format: duplicated,
});

//This will paint green the unique rows
const uniqueFormat = new ConditionalFormatDuplicate({
  format: unique,
  invert: true,
});

//Add the conditional formats to the sheet on the same range
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

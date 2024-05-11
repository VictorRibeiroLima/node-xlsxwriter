// @ts-check

const {
  Workbook,
  ConditionalFormatFormula,
  Color,
  Format,
  Formula,
} = require('../../src/index');
const fs = require('fs');

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

// Dynamic true will be ignored by the conditional format
const isOdd = new Formula({
  formula: '=ISODD(B3)',
  dynamic: true,
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

workbook.saveToBuffer();

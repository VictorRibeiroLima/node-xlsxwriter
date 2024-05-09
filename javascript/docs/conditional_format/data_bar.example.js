// @ts-check

const {
  Workbook,
  ConditionalFormatDataBar,
  Color,
} = require('../../src/index');

const workbook = new Workbook();
const sheet = workbook.addSheet();

const lightGreen = new Color(102, 255, 102);
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

workbook.saveToBase64Sync();

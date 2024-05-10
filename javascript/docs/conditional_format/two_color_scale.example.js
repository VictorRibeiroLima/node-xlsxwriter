// @ts-check
const { ConditionalFormatTwoColorScale, Sheet } = require('../../src/index');

const sheet = new Sheet('example');

// Write a 2 color scale formats with standard Excel colors. The conditional
// format is applied from the lowest to the highest value.
const twoColorScale = new ConditionalFormatTwoColorScale();
sheet.addConditionalFormat({
  firstRow: 2,
  lastRow: 11,
  firstColumn: 1,
  lastColumn: 1,
  format: twoColorScale,
});

// Write a 2 color scale formats with standard Excel colors but a user
// defined range. Values <= 3 will be shown with the minimum color while
// values >= 7 will be shown with the maximum color.
const twoColorScaleCustomRange = new ConditionalFormatTwoColorScale({
  minRule: { type: 'number', value: 3 },
  maxRule: { type: 'number', value: 7 },
});

sheet.addConditionalFormat({
  firstRow: 2,
  lastRow: 11,
  firstColumn: 3,
  lastColumn: 3,
  format: twoColorScaleCustomRange,
});

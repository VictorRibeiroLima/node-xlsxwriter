// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const {
  Workbook,
  ConditionalFormatThreeColorScale,
} = require('../../src/index');
const fs = require('fs');

test('save to file with conditional format ("ConditionalFormatThreeColorScale")', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  // Write a 3 color scale formats with standard Excel colors. The conditional
  // format is applied from the lowest to the highest value.

  const twoColorScale = new ConditionalFormatThreeColorScale();
  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 1,
    format: twoColorScale,
  });

  // Write a 3 color scale formats with standard Excel colors but a user
  // defined range. Values <= 3 will be shown with the minimum color while
  // values >= 7 will be shown with the maximum color.
  const twoColorScaleCustomRange = new ConditionalFormatThreeColorScale({
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

  const twoColorScaleCustomRangeWithColor =
    new ConditionalFormatThreeColorScale({
      minRule: { type: 'number', value: 1 },
      midRule: { type: 'number', value: 4 },
      maxRule: { type: 'number', value: 7 },
      minColor: { red: 0, green: 0, blue: 255 },
      midColor: { red: 0, green: 255, blue: 0 },
      maxColor: { red: 255, green: 0, blue: 0 },
    });

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 5,
    lastColumn: 5,
    format: twoColorScaleCustomRangeWithColor,
  });

  for (let i = 2; i <= 11; i++) {
    sheet.writeNumber(1, i, i);
    sheet.writeNumber(3, i, i);
    sheet.writeNumber(5, i, i);
  }

  await workbook.saveToFile(
    './temp/conditional_format/save-to-file-with-format-three-color-scale.xlsx',
  );
  assert.ok(
    fs.existsSync(
      './temp/conditional_format/save-to-file-with-format-three-color-scale.xlsx',
    ),
  );
});

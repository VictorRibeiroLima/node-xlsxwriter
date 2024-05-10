// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const { Workbook, ConditionalFormatBlank, Format } = require('../../src/index');
const fs = require('fs');

test('save to file with conditional format ("ConditionalFormatBlank")', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const greenFormat = new Format({
    backgroundColor: { red: 102, green: 255, blue: 102 },
  });

  const redFormat = new Format({
    backgroundColor: { red: 255, green: 102, blue: 102 },
  });

  // Write a conditional format over a range.
  const blankFormat = new ConditionalFormatBlank({
    format: redFormat,
  });

  // Invert the blank conditional format to show non-blank values.
  const nonBlankFormat = new ConditionalFormatBlank({
    format: greenFormat,
    invert: true,
  });

  sheet.addConditionalFormat({
    firstRow: 0,
    lastRow: 12,
    firstColumn: 0,
    lastColumn: 0,
    format: blankFormat,
  });

  sheet.addConditionalFormat({
    firstRow: 0,
    lastRow: 12,
    firstColumn: 0,
    lastColumn: 0,
    format: nonBlankFormat,
  });

  for (let i = 0; i <= 12; i++) {
    if ((i & 1) == 0) {
      sheet.writeNumber(0, i, i);
    }
  }

  await workbook.saveToFile(
    './temp/conditional_format/save-to-file-with-format-blank.xlsx',
  );
  assert.ok(
    fs.existsSync(
      './temp/conditional_format/save-to-file-with-format-blank.xlsx',
    ),
  );
});

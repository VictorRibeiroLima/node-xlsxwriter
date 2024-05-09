const { test } = require('node:test');
const assert = require('node:assert');
const {
  Workbook,
  ConditionalFormatAverage,
  Format,
} = require('../../src/index');
const fs = require('fs');

test('save to file with format ("ConditionalFormatAverage")', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  let redFormat = new Format({
    backgroundColor: { red: 255, green: 102, blue: 102 },
  });

  let greenFormat = new Format({
    backgroundColor: { red: 102, green: 255, blue: 102 },
  });

  //Write a conditional format. The default criteria is Above Average.
  let averageFormat = new ConditionalFormatAverage({
    format: greenFormat,
  });

  //  worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 10,
    format: averageFormat,
  });

  // Write another conditional format over the same range.
  let averageFormat2 = new ConditionalFormatAverage({
    format: redFormat,
    rule: 'belowAverage',
  });

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 10,
    format: averageFormat2,
  });

  for (let i = 2; i <= 11; i++) {
    for (let j = 1; j <= 10; j++) {
      sheet.writeNumber(j, i, i);
    }
  }

  await workbook.saveToFile(
    './temp/conditional_format/save-to-file-with-format-average.xlsx',
  );
  assert.ok(
    fs.existsSync(
      './temp/conditional_format/save-to-file-with-format-average.xlsx',
    ),
  );
});

test('save to file with no format ("ConditionalFormatAverage")', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  // a Average format without a format it's useless, but it should not crash
  const averageFormat2 = new ConditionalFormatAverage({
    rule: 'belowAverage',
  });

  sheet.addConditionalFormat({
    firstRow: 2,
    lastRow: 11,
    firstColumn: 1,
    lastColumn: 10,
    format: averageFormat2,
  });

  for (let i = 2; i <= 11; i++) {
    for (let j = 1; j <= 10; j++) {
      sheet.writeNumber(j, i, i);
    }
  }

  await workbook.saveToFile(
    './temp/conditional_format/save-to-file-with-no-format-average.xlsx',
  );
  assert.ok(
    fs.existsSync(
      './temp/conditional_format/save-to-file-with-no-format-average.xlsx',
    ),
  );
});

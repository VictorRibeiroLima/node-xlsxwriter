// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const { Workbook, ConditionalFormatCell, Format } = require('../../src/index');
const fs = require('fs');

test('save to file with conditional format ("ConditionalFormatCell") greaterThanOrEqualTo', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const greenFormat = new Format({
    backgroundColor: { red: 102, green: 255, blue: 102 },
  });

  const conditionalFormat = new ConditionalFormatCell({
    format: greenFormat,
    rule: {
      type: 'greaterThanOrEqualTo',
      value: 5,
    },
  });

  sheet.addConditionalFormat({
    firstRow: 0,
    lastRow: 12,
    firstColumn: 0,
    lastColumn: 0,
    format: conditionalFormat,
  });

  for (let i = 0; i <= 12; i++) {
    sheet.writeNumber(0, i, i);
  }

  await workbook.saveToFile(
    './temp/conditional_format/save-to-file-with-format-cell-greater-than-or-equal-to.xlsx',
  );
  assert.ok(
    fs.existsSync(
      './temp/conditional_format/save-to-file-with-format-cell-greater-than-or-equal-to.xlsx',
    ),
  );
});

test('save to file with conditional format ("ConditionalFormatCell") between', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const greenFormat = new Format({
    backgroundColor: { red: 102, green: 255, blue: 102 },
  });

  const conditionalFormat = new ConditionalFormatCell({
    format: greenFormat,
    rule: {
      type: 'between',
      value: 3,
      optionalValue: 7,
    },
  });

  sheet.addConditionalFormat({
    firstRow: 0,
    lastRow: 12,
    firstColumn: 0,
    lastColumn: 0,
    format: conditionalFormat,
  });

  for (let i = 0; i <= 12; i++) {
    sheet.writeNumber(0, i, i);
  }

  await workbook.saveToFile(
    './temp/conditional_format/save-to-file-with-format-cell-between.xlsx',
  );
  assert.ok(
    fs.existsSync(
      './temp/conditional_format/save-to-file-with-format-cell-between.xlsx',
    ),
  );

  const badConditionalFormat = new ConditionalFormatCell({
    format: greenFormat,
    rule: {
      type: 'between',
      value: 3,
    },
  });

  try {
    sheet.addConditionalFormat({
      firstRow: 12,
      lastRow: 24,
      firstColumn: 0,
      lastColumn: 0,
      format: badConditionalFormat,
    });

    await workbook.saveToFile(
      './temp/conditional_format/save-to-file-with-format-cell-between.xlsx',
    );
  } catch (error) {
    assert.strictEqual(
      error.message,
      `Missing second value for 'between' rule`,
    );
  }
});

test('save to file with conditional format ("ConditionalFormatCell") notBetween', async (t) => {
  const today = new Date();
  const tomorrow = addDays(today, 1);

  const dateFormat = new Format({
    numFmt: 'dd/mm/yyyy',
  });

  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const greenFormat = new Format({
    backgroundColor: { red: 102, green: 255, blue: 102 },
  });

  const conditionalFormat = new ConditionalFormatCell({
    format: greenFormat,
    rule: {
      type: 'notBetween',
      value: today,
      optionalValue: tomorrow,
    },
  });

  sheet.addConditionalFormat({
    firstRow: 0,
    lastRow: 12,
    firstColumn: 0,
    lastColumn: 0,
    format: conditionalFormat,
  });

  for (let i = 0; i <= 12; i++) {
    sheet.writeDate(0, i, addDays(today, i), dateFormat);
  }

  await workbook.saveToFile(
    './temp/conditional_format/save-to-file-with-format-cell-not-between.xlsx',
  );

  assert.ok(
    fs.existsSync(
      './temp/conditional_format/save-to-file-with-format-cell-not-between.xlsx',
    ),
  );

  const badConditionalFormat = new ConditionalFormatCell({
    format: greenFormat,
    rule: {
      type: 'notBetween',
      value: today,
      optionalValue: 'aaa',
    },
  });

  try {
    sheet.addConditionalFormat({
      firstRow: 12,
      lastRow: 24,
      firstColumn: 0,
      lastColumn: 0,
      format: badConditionalFormat,
    });

    await workbook.saveToFile(
      './temp/conditional_format/save-to-file-with-format-cell-not-between.xlsx',
    );
  } catch (error) {
    assert.strictEqual(error.message, `failed to downcast any to object`);
  }
});

function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

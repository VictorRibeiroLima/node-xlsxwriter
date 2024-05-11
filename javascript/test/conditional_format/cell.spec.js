// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const { Workbook, ConditionalFormatCell, Format } = require('../../src/index');
const fs = require('fs');
const findRootDir = require('../util');
const rootPath = findRootDir(__dirname);
const path = rootPath + '/temp/conditional_format';

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

  for (let row = 0; row <= 12; row++) {
    sheet.writeNumber(row, 0, row);
  }

  await workbook.saveToFile(
    `${path}/save-to-file-with-format-cell-greater-than-or-equal-to.xlsx`,
  );
  assert.ok(
    fs.existsSync(
      `${path}/save-to-file-with-format-cell-greater-than-or-equal-to.xlsx`,
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

  for (let row = 0; row <= 12; row++) {
    sheet.writeNumber(row, 0, row);
  }

  await workbook.saveToFile(
    `${path}/save-to-file-with-format-cell-between.xlsx`,
  );
  assert.ok(
    fs.existsSync(`${path}/save-to-file-with-format-cell-between.xlsx`),
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
      `${path}/save-to-file-with-format-cell-between.xlsx`,
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

  for (let row = 0; row <= 12; row++) {
    sheet.writeDate(row, 0, addDays(today, row), dateFormat);
  }

  await workbook.saveToFile(
    `${path}/save-to-file-with-format-cell-not-between.xlsx`,
  );

  assert.ok(
    fs.existsSync(`${path}/save-to-file-with-format-cell-not-between.xlsx`),
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
      `${path}/save-to-file-with-format-cell-not-between.xlsx`,
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

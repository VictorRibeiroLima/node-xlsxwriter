// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const { Workbook, ConditionalFormatCell, Format } = require('../../src/index');

function greaterThanOrEqualTo() {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const greenFormat = new Format({
    backgroundColor: { red: 102, green: 255, blue: 102 },
  });

  //Paints every cell that is greater than or equal to 5 with a green background
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

  workbook.saveToBufferSync();
}

function between() {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  const greenFormat = new Format({
    backgroundColor: { red: 102, green: 255, blue: 102 },
  });

  //Paints every cell that is between 3 and 7 with a green background
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

  workbook.saveToBufferSync();

  //This will throw an error when saved because the optionalValue is missing
  const badConditionalFormat = new ConditionalFormatCell({
    format: greenFormat,
    rule: {
      type: 'between',
      value: 3,
    },
  });
}

function notBetween() {
  const addDays = (date, days) => {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  };

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

  //Paints every cell that is not between today and tomorrow with a green background
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

  workbook.saveToBase64Sync();

  //This will throw an error when saved because the optionalValue is different from the of value
  const badConditionalFormat = new ConditionalFormatCell({
    format: greenFormat,
    rule: {
      type: 'notBetween',
      value: today,
      optionalValue: 'aaa',
    },
  });
}

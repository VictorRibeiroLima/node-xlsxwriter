// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const { Workbook, Link, Format, Color, Border } = require('../src/index');
const fs = require('fs');

test('save to base64 basic', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const link = new Link('http://example.com', 'Example', 'tooltip');
  sheet.writeString(1, 1, 'Hello, World!');
  sheet.writeNumber(2, 1, 42);
  sheet.writeLink(3, 1, link);
  const base64 = await workbook.saveToBase64();
  assert.ok(typeof base64 === 'string');
  fs.writeFileSync(
    './temp/save_to_base64_basic.xlsx',
    Buffer.from(base64, 'base64'),
  );
});

test('save to base64 with format', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const format = new Format({
    align: 'center',
    bold: true,
    backgroundColor: new Color({
      red: 255,
      green: 0,
      blue: 0,
    }),
    fontSize: 16,
    underline: 'double',
    fontScheme: 'minor',
    fontName: 'Arial',
  });

  format.setBorder(new Border('thin', new Color()));
  sheet.writeString(1, 1, 'Hello, World!', format);
  const base64 = await workbook.saveToBase64();
  assert.ok(typeof base64 === 'string');
  fs.writeFileSync(
    './temp/save_to_base64_with_format.xlsx',
    Buffer.from(base64, 'base64'),
  );
});

test('save to base64 sync', (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const link = new Link('http://example.com', 'Example', 'tooltip');
  sheet.writeString(1, 1, 'Hello, World!');
  sheet.writeNumber(2, 1, 42);
  sheet.writeLink(3, 1, link);
  const base64 = workbook.saveToBase64Sync();
  assert.ok(typeof base64 === 'string');
  fs.writeFileSync(
    './temp/save_to_base64_sync.xlsx',
    Buffer.from(base64, 'base64'),
  );
});

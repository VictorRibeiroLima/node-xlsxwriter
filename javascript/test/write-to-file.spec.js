// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const { Workbook, Link, Format, Color, Border } = require('../src/index');
const fs = require('fs');

test('save to file basic', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const link = new Link('http://example.com', 'Example', 'tooltip');
  sheet.writeString(1, 1, 'Hello, World!');
  sheet.writeNumber(2, 1, 42);
  sheet.writeLink(3, 1, link);
  await workbook.saveToFile('./temp/save-to-file-basic.xlsx');
  assert.ok(fs.existsSync('./temp/save-to-file-basic.xlsx'));
});

test('save to file with format', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const format = new Format({
    align: 'center',
    bold: true,
    backgroundColor: new Color(255, 0, 0),
    fontSize: 16,
    underline: 'double',
    fontScheme: 'minor',
    fontName: 'Arial',
  });

  format.setBorder(new Border('thin', new Color(0, 0, 0)));
  sheet.writeString(1, 1, 'Hello, World!', format);
  await workbook.saveToFile('./temp/save-to-file-with-format.xlsx');
  assert.ok(fs.existsSync('./temp/save-to-file-with-format.xlsx'));
});

test('save to file basic sync', (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const link = new Link('http://example.com', 'Example', 'tooltip');
  sheet.writeString(1, 1, 'Hello, World!');
  sheet.writeNumber(2, 1, 42);
  sheet.writeLink(3, 1, link);
  workbook.saveToFileSync('./temp/save-to-file-basic-sync.xlsx');
  assert.ok(fs.existsSync('./temp/save-to-file-basic-sync.xlsx'));
});

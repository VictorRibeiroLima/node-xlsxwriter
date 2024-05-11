// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const { Workbook, Link, Format, Color, Border } = require('../src/index');
const fs = require('fs');
const findRootDir = require('./util');
const rootPath = findRootDir(__dirname);
const path = rootPath + '/temp';

test('save to file basic', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const link = new Link('http://example.com', 'Example', 'tooltip');
  sheet.writeString(1, 1, 'Hello, World!');
  sheet.writeNumber(2, 1, 42);
  sheet.writeLink(3, 1, link);
  await workbook.saveToFile(`${path}/save_to_file_basic.xlsx`);
  assert.ok(fs.existsSync(`${path}/save_to_file_basic.xlsx`));
});

test('save to file with format', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const format = new Format({
    align: 'center',
    bold: true,
    backgroundColor: new Color({ red: 255 }),
    fontSize: 16,
    underline: 'double',
    fontScheme: 'minor',
    fontName: 'Arial',
  });

  format.setBorder(new Border('thin', new Color()));
  sheet.writeString(1, 1, 'Hello, World!', format);
  await workbook.saveToFile(`${path}/save_to_file_with_format.xlsx`);
  assert.ok(fs.existsSync(`${path}/save_to_file_with_format.xlsx`));
});

test('save to file basic sync', (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const link = new Link('http://example.com', 'Example', 'tooltip');
  sheet.writeString(1, 1, 'Hello, World!');
  sheet.writeNumber(2, 1, 42);
  sheet.writeLink(3, 1, link);
  workbook.saveToFileSync(`${path}/save_to_file_basic_sync.xlsx`);
  assert.ok(fs.existsSync(`${path}/save_to_file_basic_sync.xlsx`));
});

// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const { Workbook, Link, Format, Color, Border } = require('../src/index');

const fs = require('fs');

test('save to buffer basic', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const link = new Link('http://example.com', 'Example', 'tooltip');
  sheet.writeString(1, 1, 'Hello, World!');
  sheet.writeNumber(2, 1, 42);
  sheet.writeLink(3, 1, link);
  const buffer = await workbook.saveToBuffer();
  assert.ok(buffer instanceof Buffer);
  fs.writeFileSync('./temp/save_to_buffer_basic.xlsx', buffer);
});

test('save to buffer with format', async (t) => {
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
  const buffer = await workbook.saveToBuffer();
  assert.ok(buffer instanceof Buffer);
  fs.writeFileSync('./temp/save_to_buffer_with_format.xlsx', buffer);
});

test('save to buffer with random values', async (t) => {
  const promise = new Promise((resolve) => {
    resolve('Hello, World!');
  });
  const obj = {};
  const date = new Date();
  const num = 42;
  const arr = [1, 2, 3];

  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  sheet.writeCell(1, 1, promise);
  sheet.writeCell(2, 1, obj);
  sheet.writeCell(3, 1, date);
  sheet.writeCell(4, 1, num);
  sheet.writeCell(5, 1, arr);
  const buffer = await workbook.saveToBuffer();
  assert.ok(buffer instanceof Buffer);
  fs.writeFileSync('./temp/save_to_buffer_with_random_values.xlsx', buffer);
});

test('save to buffer with multiple sheets', async (t) => {
  const workbook = new Workbook();
  const sheet1 = workbook.addSheet();
  const sheet2 = workbook.addSheet();
  sheet1.writeString(1, 1, 'Sheet 1');
  sheet2.writeString(1, 1, 'Sheet 2');
  const buffer = await workbook.saveToBuffer();
  assert.ok(buffer instanceof Buffer);
  fs.writeFileSync('./temp/save_to_buffer_with_multiple_sheets.xlsx', buffer);
});

test('buffer from json', async (t) => {
  const objects = [
    {
      name: 'John',
      age: 30,
      website: new Link('http://example.com', 'Example', 'tooltip'),
      date: new Date(),
    },
    {
      name: 'Jane',
      age: 25,
      website: new Link('http://example.com', 'Example', 'tooltip'),
      date: new Date(),
    },
  ];

  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  sheet.writeFromJson(objects);

  const buffer = await workbook.saveToBuffer();
  assert.ok(buffer instanceof Buffer);
  fs.writeFileSync('./temp/buffer_from_json.xlsx', buffer);
});

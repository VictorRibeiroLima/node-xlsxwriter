// @ts-check
import { test } from 'node:test';
import assert from 'node:assert';
import * as writer from '../src/index.js';
const { Workbook, Link, Format, Color, Border } = writer;

test('(mjs) save to buffer basic', async (t) => {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const link = new Link('http://example.com', 'Example', 'tooltip');
  sheet.writeString(1, 1, 'Hello, World!');
  sheet.writeNumber(2, 1, 42);
  sheet.writeLink(3, 1, link);
  const buffer = await workbook.saveToBuffer();
  assert.ok(buffer instanceof Buffer);
});

test('(mjs) save to buffer with format', async (t) => {
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
});

test('(mjs) save to buffer with random values', async (t) => {
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
});

test('(mjs) save to buffer with multiple sheets', async (t) => {
  const workbook = new Workbook();
  const sheet1 = workbook.addSheet();
  const sheet2 = workbook.addSheet();
  sheet1.writeString(1, 1, 'Sheet 1');
  sheet2.writeString(1, 1, 'Sheet 2');
  const buffer = await workbook.saveToBuffer();
  assert.ok(buffer instanceof Buffer);
});

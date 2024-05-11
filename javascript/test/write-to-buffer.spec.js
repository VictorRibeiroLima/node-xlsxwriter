// @ts-check
const { test } = require('node:test');
const assert = require('node:assert');
const { Workbook, Link, Format, Color, Border } = require('../src/index');
const findRootDir = require('./util');
const rootPath = findRootDir(__dirname);
const path = rootPath + '/temp';
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
  fs.writeFileSync(`${path}/save_to_buffer_basic.xlsx`, buffer);
});

test('save to buffer with format', async (t) => {
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
  const buffer = await workbook.saveToBuffer();
  assert.ok(buffer instanceof Buffer);
  fs.writeFileSync(`${path}/save_to_buffer_with_format.xlsx`, buffer);
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
  fs.writeFileSync(`${path}/save_to_buffer_with_random_values.xlsx`, buffer);
});

test('save to buffer with multiple sheets', async (t) => {
  const workbook = new Workbook();
  const sheet1 = workbook.addSheet();
  const sheet2 = workbook.addSheet();
  sheet1.writeString(1, 1, 'Sheet 1');
  sheet2.writeString(1, 1, 'Sheet 2');
  const buffer = await workbook.saveToBuffer();
  assert.ok(buffer instanceof Buffer);
  fs.writeFileSync(`${path}/save_to_buffer_with_multiple_sheets.xlsx`, buffer);
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
  fs.writeFileSync(`${path}/buffer_from_json.xlsx`, buffer);
});

test('specific buffer', async (t) => {
  const format = new Format({
    numFmt: 'dd/mm/yy',
  });

  const objects = [
    {
      nr_contrato: '10887816',
      cpf: '18811851505',
      data_vencto_divida: new Date('2024-05-07T23:03:56.117Z'),
      data_termino_divida: new Date('2024-05-07T23:03:56.121Z'),
      praca: 'AA',
      valor: '2,00',
      data_compromisso: null,
      vr_contrato: '3500,00',
      nome_principal: 'Manuel Christiansen',
      endereco: 'Earnest Brekke 234',
      complemento_endereco: '',
      bairro: '',
      municipio: 'Leah Wintheiser',
      cep: '22222-220',
      estado: 'AA',
      ddd_telefone: '5535958549689',
    },
    {
      nr_contrato: '10887816',
      cpf: '06224356394',
      data_vencto_divida: new Date('2024-05-07T23:03:56.126Z'),
      data_termino_divida: new Date('2024-05-07T23:03:56.130Z'),
      praca: 'AA',
      valor: '2,00',
      data_compromisso: null,
      vr_contrato: '3500,00',
      nome_principal: 'Pablo Stracke MD',
      endereco: 'Beth Weissnat 234',
      complemento_endereco: '',
      bairro: '',
      municipio: 'James Kling',
      cep: '22222-220',
      estado: 'AA',
      ddd_telefone: '5581937887608',
    },
    {
      nr_contrato: '10887816',
      cpf: '43632826730',
      data_vencto_divida: new Date('2024-05-07T23:03:56.134Z'),
      data_termino_divida: new Date('2024-05-07T23:03:56.137Z'),
      praca: 'AA',
      valor: '2,00',
      data_compromisso: null,
      vr_contrato: '3500,00',
      nome_principal: 'Ada Ryan',
      endereco: 'Mark Grady 234',
      complemento_endereco: '',
      bairro: '',
      municipio: 'Dora Ward I',
      cep: '22222-220',
      estado: 'AA',
      ddd_telefone: '5592975561790',
    },
    {
      nr_contrato: '10887816',
      cpf: '43632826730',
      data_vencto_divida: new Date('2024-05-07T23:03:56.134Z'),
      data_termino_divida: new Date('2024-05-07T23:03:56.137Z'),
      praca: 'AA',
      valor: '2,00',
      data_compromisso: new Date('2024-05-07T23:03:56.137Z'),
      vr_contrato: '3500,00',
      nome_principal: 'Ada Ryan',
      endereco: 'Mark Grady 234',
      complemento_endereco: '',
      bairro: '',
      municipio: 'Dora Ward I',
      cep: '22222-220',
      estado: 'AA',
      ddd_telefone: '5592975561790',
    },
  ];

  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  sheet.writeFromJson(objects, {
    columnFormats: {
      data_vencto_divida: format,
      data_termino_divida: format,
      data_compromisso: format,
    },
  });

  const buffer = await workbook.saveToBuffer();
  assert.ok(buffer instanceof Buffer);
  fs.writeFileSync(`${path}/specific_buffer.xlsx`, buffer);
});

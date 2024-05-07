# Node-xlsxwriter
node-xlsxwriter is a Node.js wrapper for the Rust library [rust_xlsxwriter](
  https://docs.rs/rust_xlsxwriter/0.64.2/rust_xlsxwriter/index.html).

It allows you to create Excel files in the `.xlsx` format,with a high level of performance,in a simple way.

it's an ongoing project, so it's not ready for production yet.

## Installation
```bash
npm install node-xlsxwriter
```

## Usage
```javascript
const { Workbook, Sheet, Format, Color, Border } = require('node-xlsxwriter');
/*
import * as xlsx from 'node-xlsxwriter';
const { Workbook, Sheet, Format, Color, Border } = xlsx;
for typescript/Es6
*/

const workbook = new Workbook();
const sheet = new Sheet('SomeSheet');
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
// const sheet = workbook.addSheet(); if you don't care about the name
sheet.writeString(0, 0, 'Hello');
sheet.writeNumber(0, 1, 123);
sheet.writeLink(0, 2, 'https://github.com');
sheet.writeCell(0, 3, 'World', 'string');
sheet.writeString(1, 0, 'Hello', format);
workbook.pushSheet(sheet); // not necessary if you use addSheet

const buffer = workbook.saveToBufferSync();
const asyncBuffer = await workbook.saveToBuffer();
const path = 'path/to/file.xlsx';
const base64 = workbook.saveToBase64Sync();
workbook.saveToFileSync(path);
```

You also can use the `writeFromJson` method to create a sheet from a JSON object.

```javascript
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
```
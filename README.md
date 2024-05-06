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
const { Workbook, Sheet } = require('node-xlsxwriter');

const workbook = new Workbook();
const sheet = new Sheet('SomeSheet');
// const sheet = workbook.addSheet(); if you don't care about the name
sheet.writeString(0, 0, 'Hello');
sheet.writeNumber(0, 1, 123);
sheet.writeLink(0, 2, 'https://github.com');
sheet.writeCell(0, 3, 'World', 'string');
workbook.pushSheet(sheet); // not necessary if you use addSheet

const buffer = workbook.saveToBufferSync();
const asyncBuffer = await workbook.saveToBuffer();
const path = 'path/to/file.xlsx';
workbook.saveToFileSync(path);
```
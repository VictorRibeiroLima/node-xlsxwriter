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

Obs: The `writeFromJson` method trades performance for convenience,
so if will only generate the JSON for the sheet and not use anywhere, it's better to use the `writeCell` method instead.

# More complex examples
More complex examples can be found in the `javascript/docs` folder.

# Warnings ⚠️

## Binary
The installation process will automatically download the `./native/node-xlsxwriter.node` binary for your system, but if you have any problems, you can download it manually from the [releases page](
  https://github.com/VictorRibeiroLima/node-xlsxwriter/releases
).

## Work in progress 
Remember that this project is still in development, so some features may not work as expected.

Also, the API may change in the future.

Any help is welcome.

## Format
Formats you can and should reuse them, because they are cached internally, so you can save memory and improve performance.

Example:

`don't do this`
```javascript
for (let i = 0; i < 100; i++) {
  const format = new Format();
  sheet.writeString(0, i, 'Hello', format);
}
```

`do this`
```javascript
const format = new Format();
for (let i = 0; i < 100; i++) {
  sheet.writeString(0, i, 'Hello', format);
}
```

The impact is not noticeable in small files, but it can be significant in large files.

# Building from source
If you want to build the project from source, you need to have Rust [installed](https://www.rust-lang.org) on your machine.

```bash
git clone git@github.com:VictorRibeiroLima/node-xlsxwriter.git
cd node-xlsxwriter
npm install
npm run release:native
```

This will generate the `./native/node-xlsxwriter.node` binary for your system.
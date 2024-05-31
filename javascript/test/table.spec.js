// @ts-check

const {
  Table,
  Workbook,
  TableColumn,
  Formula,
  Format,
  Color,
} = require('../src/index');
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const findRootDir = require('./util');
const rootPath = findRootDir(__dirname);
const path = rootPath + '/temp';

function setup() {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();
  const items = ['Apples', 'Pears', 'Bananas', 'Oranges'];
  const data = [
    [10000, 5000, 8000, 6000],
    [2000, 3000, 4000, 5000],
    [6000, 6000, 6500, 6000],
    [500, 300, 200, 700],
  ];

  //Write data
  for (let i = 0; i < items.length; i++) {
    sheet.writeString(i + 3, 1, items[i]);
  }

  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
      sheet.writeNumber(i + 3, j + 2, data[i][j]);
    }
  }

  //Increase column width to fit data
  for (let i = 1; i < 10; i++) {
    sheet.addColumnConfig({
      index: i,
      size: {
        value: 16,
      },
    });
  }
  return workbook;
}

test('table basic', async (t) => {
  const workbook = setup();
  const sheet = workbook.sheets[0];
  //Create table
  const columns = [];
  columns.push(
    new TableColumn({
      header: 'Product',
      totalLabel: 'Totals',
    }),
  );
  columns.push(
    new TableColumn({
      header: 'Quarter 1',
      totalFunction: {
        type: 'sum',
      },
    }),
  );
  columns.push(
    new TableColumn({
      header: 'Quarter 2',
      totalFunction: {
        type: 'sum',
      },
    }),
  );
  columns.push(
    new TableColumn({
      header: 'Quarter 3',
      totalFunction: {
        type: 'sum',
      },
    }),
  );
  columns.push(
    new TableColumn({
      header: 'Quarter 4',
      totalFunction: {
        type: 'sum',
      },
    }),
  );
  columns.push(
    new TableColumn({
      header: 'Year',
      totalFunction: {
        type: 'sum',
      },
      formula: new Formula({
        formula: 'SUM(Table1[@[Quarter 1]:[Quarter 4]])',
      }),
    }),
  );

  const table = new Table({
    columns,
    totalRow: true,
  });

  sheet.addTable({
    firstColumn: 1,
    firstRow: 2,
    lastColumn: 6,
    lastRow: 7,
    table,
  });

  await workbook.saveToFile(`${path}/table_basic.xlsx`);
  assert(fs.existsSync(`${path}/table_basic.xlsx`), 'file exists');
});

test('table with format', async (t) => {
  const workbook = setup();
  const sheet = workbook.sheets[0];
  const format = new Format({
    numFmt: '$#,##0.00',
  });
  const blue = new Format({
    fontColor: new Color({ blue: 255 }),
  });
  const red = new Format({
    fontColor: new Color({ red: 255 }),
  });
  const green = new Format({
    fontColor: new Color({ green: 255 }),
  });
  const yellow = new Format({
    fontColor: new Color({ red: 255, green: 255 }),
  });

  const columns = [];
  columns.push(
    new TableColumn({
      header: 'Product',
    }),
  );

  columns.push(
    new TableColumn({
      header: 'Quarter 1',
      format,
      headerFormat: blue,
    }),
  );

  columns.push(
    new TableColumn({
      header: 'Quarter 2',
      format,
      headerFormat: red,
    }),
  );

  columns.push(
    new TableColumn({
      header: 'Quarter 3',
      format,
      headerFormat: green,
    }),
  );

  columns.push(
    new TableColumn({
      header: 'Quarter 4',
      format,
      headerFormat: yellow,
    }),
  );

  const table = new Table({
    columns,
  });

  sheet.addTable({
    firstColumn: 1,
    firstRow: 2,
    lastColumn: 5,
    lastRow: 6,
    table,
  });

  await workbook.saveToFile(`${path}/table_with_format.xlsx`);

  assert(fs.existsSync(`${path}/table_with_format.xlsx`), 'file exists');
});

test('table with table function', async (t) => {
  const workbook = setup();
  const sheet = workbook.sheets[0];
  const columns = [];

  columns.push(
    new TableColumn({
      totalLabel: 'Totals',
    }),
  );
  columns.push(
    new TableColumn({
      totalFunction: {
        type: 'sum',
      },
    }),
  );
  columns.push(
    new TableColumn({
      totalFunction: {
        type: 'sum',
      },
    }),
  );
  columns.push(
    new TableColumn({
      totalFunction: {
        type: 'sum',
      },
    }),
  );
  columns.push(
    new TableColumn({
      totalFunction: {
        type: 'custom',
        formula: new Formula({
          formula: 'SUM([Column5])',
        }),
      },
    }),
  );

  const table = new Table({
    columns,
    totalRow: true,
  });

  sheet.addTable({
    firstColumn: 1,
    firstRow: 2,
    lastColumn: 5,
    lastRow: 7,
    table,
  });

  await workbook.saveToFile(`${path}/table_with_table_function.xlsx`);
  assert(
    fs.existsSync(`${path}/table_with_table_function.xlsx`),
    'file exists',
  );
});

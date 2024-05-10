// @ts-check
const { Workbook, Format, Formula, Sheet, colorUtils } = require('../src');

/*
  //------------------------------------------------------------------------
  //                            WANING: 
  //    I have tested the output file of some of those formulas in various xlsx readers.
  //    (Excel, LibreOffice, Google Sheets)
  //    And only Excel was able to read the formulas correctly,but it seems to be a problem with the xlsx reader,not the xlsx file.
  //------------------------------------------------------------------------
*/

const workbook = new Workbook();

// Create some header formats to use in the worksheets.
const header1 = new Format({
  foregroundColor: colorUtils.hexToColor('74AC4C'),
  fontColor: colorUtils.hexToColor('FFFFFF'),
});

const header2 = new Format({
  foregroundColor: colorUtils.hexToColor('528FD3'),
  fontColor: colorUtils.hexToColor('FFFFFF'),
});

// -----------------------------------------------------------------------
// Example of using the FILTER() function.
// -----------------------------------------------------------------------
const filterSheet = new Sheet('Filter');
//I have found i problem in every xlsx reader that i have tested besides Excel for this formula
const filterFormula = new Formula({
  formula: '=FILTER(A1:D17,C1:C17=K2)',
  dynamic: true,
});
filterSheet.writeFormula(1, 5, filterFormula);

// Write the data the function will work on.
filterSheet.writeString(0, 10, 'Product', header2);
filterSheet.writeString(1, 10, 'Apple');
filterSheet.writeString(0, 5, 'Region', header2);
filterSheet.writeString(0, 6, 'Sales Rep', header2);
filterSheet.writeString(0, 7, 'Product', header2);
filterSheet.writeString(0, 8, 'Units', header2);

//Add sample worksheet data to work on.
writeWorkSheetData(filterSheet, header1);

// -----------------------------------------------------------------------
// Example of using the UNIQUE() function.
// -----------------------------------------------------------------------
const uniqueSheet = new Sheet('Unique');

const uniqueFormula = new Formula({
  formula: '=UNIQUE(B2:B17)',
  dynamic: true,
});

const uniqueSortFormula = new Formula({
  formula: '=SORT(UNIQUE(B2:B17))',
  dynamic: true,
});

uniqueSheet.writeFormula(1, 5, uniqueFormula);
uniqueSheet.writeFormula(1, 7, uniqueSortFormula);

// Write the data the function will work on.
uniqueSheet.writeString(0, 5, 'Unique Sales Rep', header2);
uniqueSheet.writeString(0, 7, 'Sorted Unique Sales Rep', header2);
writeWorkSheetData(uniqueSheet, header1);

// -----------------------------------------------------------------------
// Example of using the SORT() function.
// -----------------------------------------------------------------------
const sortSheet = new Sheet('Sort');

const sortFormula = new Formula({
  formula: '=SORT(B2:B17)',
  dynamic: true,
});

const sortAndFilterFormula = new Formula({
  formula: '=SORT(FILTER(C2:D17,D2:D17>5000,""),2,1)',
  dynamic: true,
});

sortSheet.writeFormula(1, 5, sortFormula);

sortSheet.writeFormula(1, 7, sortAndFilterFormula);

// Write the data the function will work on.
sortSheet.writeString(0, 5, 'Sales Rep', header2);
sortSheet.writeString(0, 7, 'Product"', header2);
sortSheet.writeString(0, 8, 'Units', header2);

writeWorkSheetData(sortSheet, header1);

// -----------------------------------------------------------------------
// Example of using the SORTBY() function.
// -----------------------------------------------------------------------
const sortbySheet = new Sheet('SortBy');

const sortbyFormula = new Formula({
  formula: '=SORTBY(A2:B9,B2:B9)',
  dynamic: true,
});

sortbySheet.writeFormula(1, 3, sortbyFormula);

// Write the data the function will work on.
sortbySheet.writeString(0, 0, 'Name', header2);
sortbySheet.writeString(0, 1, 'Age', header2);

sortbySheet.writeString(1, 0, 'Tom');
sortbySheet.writeString(2, 0, 'Fred');
sortbySheet.writeString(3, 0, 'Amy');
sortbySheet.writeString(4, 0, 'Sal');
sortbySheet.writeString(5, 0, 'Fritz');
sortbySheet.writeString(6, 0, 'Srivan');
sortbySheet.writeString(7, 0, 'Xi');
sortbySheet.writeString(8, 0, 'Hector');

sortbySheet.writeNumber(1, 1, 52);
sortbySheet.writeNumber(2, 1, 65);
sortbySheet.writeNumber(3, 1, 22);
sortbySheet.writeNumber(4, 1, 73);
sortbySheet.writeNumber(5, 1, 19);
sortbySheet.writeNumber(6, 1, 39);
sortbySheet.writeNumber(7, 1, 19);
sortbySheet.writeNumber(8, 1, 66);

sortbySheet.writeString(0, 3, 'Name', header2);
sortbySheet.writeString(0, 4, 'Age', header2);

// -----------------------------------------------------------------------
// Example of using dynamic ranges with older Excel functions.
// -----------------------------------------------------------------------

const dynamicRangeSheet = new Sheet('Dynamic Range');

const dynamicRangeFormula = new Formula({
  formula: '=LEN(A1:A3)',
  dynamic: true,
});

dynamicRangeSheet.addArrayFormula({
  firstColumn: 1,
  firstRow: 0,
  lastColumn: 1,
  lastRow: 2,
  formula: dynamicRangeFormula,
});

dynamicRangeSheet.writeString(0, 0, 'Foo');
dynamicRangeSheet.writeString(1, 0, 'Food');
dynamicRangeSheet.writeString(2, 0, 'Frood');

// -----------------------------------------------------------------------
workbook.pushSheet(sortSheet);
workbook.pushSheet(filterSheet);
workbook.pushSheet(uniqueSheet);
workbook.pushSheet(sortbySheet);
workbook.pushSheet(dynamicRangeSheet);

workbook.saveToFile('dynamic_formula.xlsx');

//--------------------------------------------END OF EXAMPLES---------------------------------------------------
/**
 * A simple function and data structure to populate some of the worksheets.
 * @param {Sheet} worksheet - The worksheet to populate.
 * @param {Format} header - The header format to use.
 
 */

function writeWorkSheetData(worksheet, header) {
  /**
   * @type {Array} worksheetData - The data for the worksheet.
   * @property {string} 0 - The region where the sale was made.
   * @property {string} 1 - The name of the salesperson who made the sale.
   * @property {string} 2 - The type of fruit that was sold.
   * @property {number} 3 - The quantity of the fruit that was sold.
   */
  const worksheetData = [
    ['East', 'Tom', 'Apple', 6380],
    ['West', 'Fred', 'Grape', 5619],
    ['North', 'Amy', 'Pear', 4565],
    ['South', 'Sal', 'Banana', 5323],
    ['East', 'Fritz', 'Apple', 4394],
    ['West', 'Sravan', 'Grape', 7195],
    ['North', 'Xi', 'Pear', 5231],
    ['South', 'Hector', 'Banana', 2427],
    ['East', 'Tom', 'Banana', 4213],
    ['West', 'Fred', 'Pear', 3239],
    ['North', 'Amy', 'Grape', 6520],
    ['South', 'Sal', 'Apple', 1310],
    ['East', 'Fritz', 'Banana', 6274],
    ['West', 'Sravan', 'Pear', 4894],
    ['North', 'Xi', 'Grape', 7580],
    ['South', 'Hector', 'Apple', 9814],
  ];

  worksheet.writeString(0, 0, 'Region', header);
  worksheet.writeString(0, 1, 'Sales Rep', header);
  worksheet.writeString(0, 2, 'Product', header);
  worksheet.writeString(0, 3, 'Units', header);

  let row = 1;
  for (const data of worksheetData) {
    worksheet.writeString(row, 0, data[0]);
    worksheet.writeString(row, 1, data[1]);
    worksheet.writeString(row, 2, data[2]);
    worksheet.writeNumber(row, 3, data[3]);
    row += 1;
  }
}

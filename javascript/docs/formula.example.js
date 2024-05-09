// @ts-check
const { Workbook, Format, Formula, ConditionalFormatCell } = require('../src');

/*
Excel stores formulas in the format of the US English version, regardless of the language or locale of the end-user’s version of Excel.
Therefore all formula function names written using node-xlsxwriter must be in English. 
In addition, formulas must be written with the US style separator/range operator which is a comma (not semi-colon).

OK.
const formula = new Formula({
  formula: '=SUM(1, 2, 3)',
});

Semi-colon separator. Causes Excel error on file opening.
const formula = new Formula({
  formula: '=SUM(1; 2; 3)',
});

French function name. Causes Excel error on file opening.
const formula = new Formula({
  formula: '=SOMME(1, 2, 3)',
});
*/
function simpleExample() {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  //Write a formula
  const formula = new Formula({
    formula: '=10*B1 + C1',
  });

  //The result of the formula will be 10*5 + 2 = 52
  sheet.writeFormula(0, 0, formula);
  sheet.writeNumber(0, 1, 5);
  sheet.writeNumber(0, 2, 2);

  workbook.saveToBufferSync();
}

function exampleWithResult() {
  /*
  The rust_xlsxwriter (and node-xlsxwriter by extension) library doesn’t calculate the result of a formula and instead stores the value “0” as the formula result.
  It then sets a global flag in the XLSX file to say that all formulas and functions should be recalculated when the file is opened.
  This works fine with Excel and the majority of spreadsheet applications.
  However, applications that don’t have a facility to calculate formulas will only display the “0” result.
  Examples of such applications are Excel viewers, PDF converters, and some mobile device applications.
*/

  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  //Write a formula
  const formula = new Formula({
    formula: '=10*B1 + C1',
    result: '52',
  });

  //The result of the formula will be 10*5 + 2 = 52
  sheet.writeFormula(0, 0, formula);
  sheet.writeNumber(0, 1, 5);
  sheet.writeNumber(0, 2, 2);

  workbook.saveToBufferSync();
}

function exampleWithFutureFunctions() {
  const workbook = new Workbook();
  const sheet = workbook.addSheet();

  //Write a formula
  const formula = new Formula({
    formula: '=10*B1 + C1',
  });

  //The result of the formula will be 10*5 + 2 = 52
  sheet.writeFormula(0, 0, formula);
  sheet.writeNumber(0, 1, 5);
  sheet.writeNumber(0, 2, 2);

  workbook.saveToBufferSync();
}

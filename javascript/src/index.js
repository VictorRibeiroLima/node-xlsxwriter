// @ts-check

const Workbook = require('./models/workbook');
const { Sheet, ArrayFormulaSheetValue } = require('./models/sheet');
const Color = require('./models/color');
const Format = require('./models/format');
const Formula = require('./models/formula');
const Link = require('./models/link');
const { Border, DiagonalBorder } = require('./models/border');
const {
  ConditionalFormatTwoColorScale,
  ConditionalFormatThreeColorScale,
  ConditionalFormatAverage,
  ConditionalFormatBlank,
  ConditionalFormatCell,
  ConditionalFormatDataBar,
  ConditionalFormatDate,
  ConditionalFormatDuplicate,
  ConditionalFormatError,
  ConditionalFormatFormula,
  ConditionalFormatCustomIcon,
  ConditionalFormatIconSet,
  ConditionalFormatText,
  ConditionalFormatTop,
} = require('./models/conditional_format');

const { colorUtils } = require('./utils');

module.exports = {
  colorUtils,
  Workbook,
  Sheet,
  Color,
  Format,
  Formula,
  Link,
  Border,
  DiagonalBorder,
  ArrayFormulaSheetValue,
  ConditionalFormatTwoColorScale,
  ConditionalFormatThreeColorScale,
  ConditionalFormatAverage,
  ConditionalFormatBlank,
  ConditionalFormatCell,
  ConditionalFormatDataBar,
  ConditionalFormatDate,
  ConditionalFormatDuplicate,
  ConditionalFormatError,
  ConditionalFormatFormula,
  ConditionalFormatCustomIcon,
  ConditionalFormatIconSet,
  ConditionalFormatText,
  ConditionalFormatTop,
};

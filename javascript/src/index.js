// @ts-check

const Workbook = require('./models/workbook');
const Sheet = require('./models/sheet');
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
} = require('./models/conditional_format');

module.exports = {
  Workbook,
  Sheet,
  Color,
  Format,
  Formula,
  Link,
  Border,
  DiagonalBorder,
  ConditionalFormatTwoColorScale,
  ConditionalFormatThreeColorScale,
  ConditionalFormatAverage,
  ConditionalFormatBlank,
  ConditionalFormatCell,
  ConditionalFormatDataBar,
  ConditionalFormatDate,
  ConditionalFormatDuplicate,
};

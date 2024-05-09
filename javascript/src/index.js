// @ts-check

const Workbook = require('./models/workbook');
const Sheet = require('./models/sheet');
const Color = require('./models/color');
const Format = require('./models/format');
const Link = require('./models/link');
const { Border, DiagonalBorder } = require('./models/border');
const {
  ConditionalFormatTwoColorScale,
  ConditionalFormatThreeColorScale,
  ConditionalFormatAverage,
  ConditionalFormatBlank,
  ConditionalFormatCell,
  ConditionalFormatDataBar,
} = require('./models/conditional_format');

module.exports = {
  Workbook,
  Sheet,
  Color,
  Format,
  Link,
  Border,
  DiagonalBorder,
  ConditionalFormatTwoColorScale,
  ConditionalFormatThreeColorScale,
  ConditionalFormatAverage,
  ConditionalFormatBlank,
  ConditionalFormatCell,
  ConditionalFormatDataBar,
};

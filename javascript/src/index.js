// @ts-check

const Workbook = require('./models/workbook');
const Sheet = require('./models/sheet');
const Color = require('./models/color');
const Format = require('./models/format');
const Link = require('./models/link');
const { Border, DiagonalBorder } = require('./models/border');

module.exports = {
  Workbook,
  Sheet,
  Color,
  Format,
  Link,
  Border,
  DiagonalBorder,
};

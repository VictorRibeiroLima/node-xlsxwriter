// @ts-check
const {
  Workbook,
  ConditionalFormatIconSet,
  ConditionalFormatCustomIcon,
  Format,
} = require('../../src/index');

const workbook = new Workbook();
const sheet = workbook.addSheet();
const centerFormat = new Format({
  align: 'center',
});

const boldFormat = new Format({
  bold: true,
});

const defaultRulesForThree = [
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 0,
    },
  }),
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 33,
    },
  }),
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 67,
    },
  }),
];

const defaultRulesForFour = [
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 0,
    },
  }),
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 25,
    },
  }),
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 50,
    },
  }),
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 75,
    },
  }),
];

const defaultRulesForFive = [
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 0,
    },
  }),
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 20,
    },
  }),
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 40,
    },
  }),
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 60,
    },
  }),
  new ConditionalFormatCustomIcon({
    iconRule: {
      type: 'percent',
      value: 80,
    },
  }),
];

sheet.writeString(0, 0, 'Icon style conditional format', boldFormat);
// Three Traffic lights - Green is highest.
const threeTrafficLights = new ConditionalFormatIconSet({
  iconType: 'threeTrafficLights',
  icons: defaultRulesForThree,
});
sheet.addConditionalFormat({
  firstRow: 1,
  lastRow: 1,
  firstColumn: 1,
  lastColumn: 3,
  format: threeTrafficLights,
});
sheet.writeString(
  1,
  0,
  'Three Traffic lights - Green is highest',
  centerFormat,
);
sheet.writeNumber(1, 1, 1);
sheet.writeNumber(1, 2, 2);
sheet.writeNumber(1, 3, 3);

// Reversed - Red is highest.
const threeTrafficLightsReversed = new ConditionalFormatIconSet({
  iconType: 'threeTrafficLights',
  icons: defaultRulesForThree,
  reverse: true,
});

sheet.addConditionalFormat({
  firstRow: 2,
  lastRow: 2,
  firstColumn: 1,
  lastColumn: 3,
  format: threeTrafficLightsReversed,
});

sheet.writeString(2, 0, 'Three Traffic lights - Red is highest', centerFormat);
sheet.writeNumber(2, 1, 1);
sheet.writeNumber(2, 2, 2);
sheet.writeNumber(2, 3, 3);

// Icons only - The number data is hidden.
const threeTriangles = new ConditionalFormatIconSet({
  iconType: 'threeTrafficLights',
  icons: defaultRulesForThree,
  showIconsOnly: true,
});

sheet.addConditionalFormat({
  firstRow: 3,
  lastRow: 3,
  firstColumn: 1,
  lastColumn: 3,
  format: threeTriangles,
});

sheet.writeString(3, 0, 'Three Triangles - Icons only', centerFormat);
sheet.writeNumber(3, 1, 1);
sheet.writeNumber(3, 2, 2);
sheet.writeNumber(3, 3, 3);

//------------------------------------------------------------------------------------------------

sheet.writeString(4, 0, 'Other 3-5 icon examples', boldFormat);

// Three arrows.
const threeArrows = new ConditionalFormatIconSet({
  iconType: 'threeArrows',
  icons: defaultRulesForThree,
});

sheet.addConditionalFormat({
  firstRow: 5,
  lastRow: 5,
  firstColumn: 1,
  lastColumn: 3,
  format: threeArrows,
});

sheet.writeString(5, 0, 'Three Arrows', centerFormat);
sheet.writeNumber(5, 1, 1);
sheet.writeNumber(5, 2, 2);
sheet.writeNumber(5, 3, 3);

// Three symbols.
const threeSymbols = new ConditionalFormatIconSet({
  iconType: 'threeSymbols',
  icons: defaultRulesForThree,
});

sheet.addConditionalFormat({
  firstRow: 6,
  lastRow: 6,
  firstColumn: 1,
  lastColumn: 3,
  format: threeSymbols,
});

sheet.writeString(6, 0, 'Three Symbols', centerFormat);
sheet.writeNumber(6, 1, 1);
sheet.writeNumber(6, 2, 2);
sheet.writeNumber(6, 3, 3);

// Four Arrows.
const fourArrows = new ConditionalFormatIconSet({
  iconType: 'fourArrows',
  icons: defaultRulesForFour,
});

sheet.addConditionalFormat({
  firstRow: 7,
  lastRow: 7,
  firstColumn: 1,
  lastColumn: 4,
  format: fourArrows,
});

sheet.writeString(7, 0, 'Four Arrows', centerFormat);
sheet.writeNumber(7, 1, 1);
sheet.writeNumber(7, 2, 2);
sheet.writeNumber(7, 3, 3);
sheet.writeNumber(7, 4, 4);

// Four circles - Red (highest) to Black (lowest).
const fourRating = new ConditionalFormatIconSet({
  iconType: 'fourRedToBlack',
  icons: defaultRulesForFour,
});

sheet.addConditionalFormat({
  firstRow: 8,
  lastRow: 8,
  firstColumn: 1,
  lastColumn: 4,
  format: fourRating,
});

sheet.writeString(8, 0, 'Four Circles - Red to Black', centerFormat);
sheet.writeNumber(8, 1, 1);
sheet.writeNumber(8, 2, 2);
sheet.writeNumber(8, 3, 3);
sheet.writeNumber(8, 4, 4);

// Four rating histograms.
const fourRatingHistogram = new ConditionalFormatIconSet({
  iconType: 'fourHistograms',
  icons: defaultRulesForFour,
});

sheet.addConditionalFormat({
  firstRow: 9,
  lastRow: 9,
  firstColumn: 1,
  lastColumn: 4,
  format: fourRatingHistogram,
});

sheet.writeString(9, 0, 'Four Rating Histograms', centerFormat);
sheet.writeNumber(9, 1, 1);
sheet.writeNumber(9, 2, 2);
sheet.writeNumber(9, 3, 3);
sheet.writeNumber(9, 4, 4);

// Five Arrows.
const fiveArrows = new ConditionalFormatIconSet({
  iconType: 'fiveArrows',
  icons: defaultRulesForFive,
});

sheet.addConditionalFormat({
  firstRow: 10,
  lastRow: 10,
  firstColumn: 1,
  lastColumn: 5,
  format: fiveArrows,
});

sheet.writeString(10, 0, 'Five Arrows', centerFormat);
sheet.writeNumber(10, 1, 1);
sheet.writeNumber(10, 2, 2);
sheet.writeNumber(10, 3, 3);
sheet.writeNumber(10, 4, 4);
sheet.writeNumber(10, 5, 5);

// Five rating histograms.
const fiveRatingHistogram = new ConditionalFormatIconSet({
  iconType: 'fiveHistograms',
  icons: defaultRulesForFive,
});

sheet.addConditionalFormat({
  firstRow: 11,
  lastRow: 11,
  firstColumn: 1,
  lastColumn: 5,
  format: fiveRatingHistogram,
});

sheet.writeString(11, 0, 'Five Rating Histograms', centerFormat);
sheet.writeNumber(11, 1, 1);
sheet.writeNumber(11, 2, 2);
sheet.writeNumber(11, 3, 3);
sheet.writeNumber(11, 4, 4);
sheet.writeNumber(11, 5, 5);

// Five rating quadrants.
const fiveRatingQuadrants = new ConditionalFormatIconSet({
  iconType: 'fiveQuadrants',
  icons: defaultRulesForFive,
});

sheet.addConditionalFormat({
  firstRow: 12,
  lastRow: 12,
  firstColumn: 1,
  lastColumn: 5,
  format: fiveRatingQuadrants,
});

sheet.writeString(12, 0, 'Five Rating Quadrants', centerFormat);
sheet.writeNumber(12, 1, 1);
sheet.writeNumber(12, 2, 2);
sheet.writeNumber(12, 3, 3);
sheet.writeNumber(12, 4, 4);
sheet.writeNumber(12, 5, 5);

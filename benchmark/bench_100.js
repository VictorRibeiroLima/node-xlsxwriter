// @ts-check

const Benchmarkify = require('benchmarkify');
const XLSX = require('xlsx');

const setup = require('./setup');
const t = setup(100);

const benchmark = new Benchmarkify('Test 100', {
  description: 'Test 100 users',
  chartImage: true,
}).printHeader();

benchmark
  .createSuite('Data to xlsx buffer', {
    time: 100_000,
    description: 'Write xlsx to a buffer using libs',
  })

  .add('XLSX lib (write type "buffer")', () => {
    XLSX.write(t.xlsx, { type: 'buffer' });
  })

  .add('node-xlsxwritter (sync)', () => {
    t.nodeXlsxwritter.saveToBufferSync();
  })

  .add('node-xlsxwritter (async)', async () => {
    await t.nodeXlsxwritter.saveToBuffer();
  });

// @ts-ignore
benchmark.run();

const XLSX = require('xlsx');

const setup = require('./setup');
const t = setup(100_000);

async function main() {
  console.time('xlsx lib');
  XLSX.write(t.xlsx, { type: 'buffer' });
  console.timeEnd('xlsx lib');

  console.time('exel4node');
  await t.excel4Node.writeToBuffer();
  console.timeEnd('exel4node');

  console.time('node-xlsxwritter (sync)');
  t.nodeXlsxwritter.saveToBufferSync();
  console.timeEnd('node-xlsxwritter (sync)');

  console.time('node-xlsxwritter (async)');
  await t.nodeXlsxwritter.saveToBuffer();
  console.timeEnd('node-xlsxwritter (async)');
}

main().then().catch(console.error);

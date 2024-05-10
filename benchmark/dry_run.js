const XLSX = require('xlsx');

const setup = require('./setup');
const t = setup(100_000);

let timeSince = new Date().getTime();
function sayHello() {
  const timeNow = new Date().getTime();
  const timeDiff = timeNow - timeSince;
  if (timeDiff > 2) {
    console.log('time with event loop blocked:', timeDiff);
  }

  timeSince = timeNow;
}

async function main() {
  const time = setInterval(sayHello, 1);
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
  clearInterval(time);
}

main().then().catch(console.error);

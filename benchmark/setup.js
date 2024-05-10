// @ts-check
const { faker } = require('@faker-js/faker');

const xl = require('excel4node');
const XLSX = require('xlsx');
const { Workbook } = require('../javascript/src/index');

function testData(num) {
  const data = [];
  for (let i = 0; i < num; i++) {
    data.push(createRandomUser());
  }
  return data;
}
function createRandomUser() {
  return {
    userId: faker.string.uuid(),
    username: faker.internet.userName(),
    email: faker.internet.email(),
    avatar: faker.image.avatar(),
    password: faker.internet.password(),
    birthdate: faker.date.birthdate(),
    registeredAt: faker.date.past(),
  };
}

function setupXlsx(testData) {
  const worksheet = XLSX.utils.json_to_sheet(testData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet);
  return workbook;
}

function setupExcel4Node(testData) {
  const wb = new xl.Workbook();
  //write headers
  const ws = wb.addWorksheet('Sheet 1');

  ws.cell(1, 1).string('userId');
  ws.cell(1, 2).string('username');
  ws.cell(1, 3).string('email');
  ws.cell(1, 4).string('avatar');
  ws.cell(1, 5).string('password');
  ws.cell(1, 6).string('birthdate');
  ws.cell(1, 7).string('registeredAt');

  //write data
  for (let i = 0; i < testData.length; i++) {
    const user = testData[i];
    ws.cell(i + 2, 1).string(user.userId);
    ws.cell(i + 2, 2).string(user.username);
    ws.cell(i + 2, 3).string(user.email);
    ws.cell(i + 2, 4).string(user.avatar);
    ws.cell(i + 2, 5).string(user.password);
    ws.cell(i + 2, 6).date(user.birthdate);
    ws.cell(i + 2, 7).date(user.registeredAt);
  }

  return wb;
}

function setupNodeXlsxwritter(testData) {
  const wb = new Workbook();
  const sheet = wb.addSheet();

  sheet.writeString(0, 0, 'userId');
  sheet.writeString(0, 1, 'username');
  sheet.writeString(0, 2, 'email');
  sheet.writeString(0, 3, 'avatar');
  sheet.writeString(0, 4, 'password');
  sheet.writeString(0, 5, 'birthdate');
  sheet.writeString(0, 6, 'registeredAt');

  //write data
  for (let i = 0; i < testData.length; i++) {
    const user = testData[i];
    sheet.writeString(i + 1, 0, user.userId);
    sheet.writeString(i + 1, 1, user.username);
    sheet.writeString(i + 1, 2, user.email);
    sheet.writeString(i + 1, 3, user.avatar);
    sheet.writeString(i + 1, 4, user.password);
    sheet.writeDate(i + 1, 5, user.birthdate);
    sheet.writeDate(i + 1, 6, user.registeredAt);
  }
  return wb;
}

function setup(num) {
  const td = testData(num);
  const xlsx = setupXlsx(td);
  const excel4Node = setupExcel4Node(td);
  const nodeXlsxwritter = setupNodeXlsxwritter(td);
  return { xlsx, excel4Node, nodeXlsxwritter };
}
module.exports = setup;

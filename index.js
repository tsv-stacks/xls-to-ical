const XLSX = require('xlsx');
const ical = require('ical-generator');
const fs = require('fs')

// import file
const workbook = XLSX.readFile('./excel-files/one-month-template.xlsx')

// const worksheet = workbook.Sheets[workbook.SheetNames[0]]

// console.log(sheets);
// console.log(workbook.SheetNames)

const sheet = workbook.Sheets[workbook.SheetNames[0]]

const data = XLSX.utils.sheet_to_json(sheet);

// Accessing a cell
const firstRowFirstCell = sheet['B3'];
console.log(firstRowFirstCell);

// Accessing a row
const thirdRow = data[3];
console.log(thirdRow);

// date in number
const dayNum = data[2];
console.log(dayNum);
// date in day
const dayDate = data[1];
console.log(dayDate);

// separate file into sheets

// seperate file into users

// collate users

// convert into format for ical

// ical for each user

// use writeFile to make .ics

// test

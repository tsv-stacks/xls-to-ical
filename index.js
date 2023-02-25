const XLSX = require('xlsx');
const ical = require('ical-generator');
const fs = require('fs')
const moment = require('moment-timezone');


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

const month = sheet['C2']
console.log(month);

// Accessing a row
const thirdRow = data[3];
// console.log(thirdRow);

// date in number
const dayNum = data[2];
// console.log(dayNum);
// date in day
const dayDate = data[1]; // not needed
// console.log(dayDate);


// find way to combine dayNum and dayDate in a way that can be read by ical
function generateTaskArray(obj1, obj2) {
    const result = [];
    for (const key in obj1) {
        if (obj2.hasOwnProperty(key)) {
            result.push([obj2[key], obj1[key]]);
        }
    }
    return result;
}

const taskArray = generateTaskArray(thirdRow, dayNum);
console.log(taskArray);

function convertToICAL(month, tasks) {
    const ical = require('ical-generator')({
        domain: 'example.com',
        prodId: '//SuperCool App//iCal Generator//EN',
        timezone: 'Europe/London'
    });

    const monthNum = new Date(`1 ${month}`).getMonth() + 1;
    console.log(monthNum)
    tasks.forEach(task => {
        console.log('task log: ' + tasks);
        const event = ical.createEvent({
            start: new Date(`2022-${monthNum}-${task[0]}T08:30:00`),
            end: new Date(`2022-${monthNum}-${task[0]}T16:30:00`),
            timezone: 'Europe/London',
            summary: task[1],
            description: task[1],
            location: 'London'
        });
    });

    return ical.toString();
}

// Generate the iCal file contents and write to disk
const iCalData = convertToICAL('April', taskArray)
// fs.writeFileSync('my-calendar.ics', iCalData);

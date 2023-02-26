const XLSX = require('xlsx');
// const ical = require('ical-generator');
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

    let monthNum = new Date(`1 ${month}`).getMonth() + 1;
    if (isNaN(monthNum)) {
        throw new Error(`Invalid month: ${month}`);
    }

    const events = [];
    // console.log(monthNum, tasks[0])
    tasks.forEach(task => {
        if (isNaN(task[0])) {
            throw new Error(`Invalid task date: ${task[0]}`);
        }
        monthNum = fDateValue(monthNum)
        task[0] = fDateValue(task[0])
        // console.log(monthNum, task[0])
        if (shiftCheck(task[1])) {
            let newDate = fDateValue(Number(task[0]) + 1)
            // console.log(newDate);
            // way to check SS rolls over into next month
            const event = ical.createEvent({
                start: new Date(`2023-${monthNum}-${task[0]}T16:30:00`),
                end: new Date(`2023-${monthNum}-${newDate}T08:30:00`),
                timezone: 'Europe/London',
                summary: getSummary(task[1]),
                description: getSummary(task[1]),
                location: 'STC'
            });
            events.push(event);
        } else {
            const event = ical.createEvent({
                start: new Date(`2023-${monthNum}-${task[0]}T08:30:00`),
                end: new Date(`2023-${monthNum}-${task[0]}T16:30:00`),
                timezone: 'Europe/London',
                summary: getSummary(task[1]),
                description: getSummary(task[1]),
                location: 'STC'
            });
            events.push(event);
        }

    });

    return events;
}

// Generate the iCal file contents and write to disk
const events = convertToICAL('december', taskArray);
events.forEach(event => {
    // console.log(event.toString());
    // fs.writeFileSync('my-calendar.ics', event.toString());
});

function fDateValue(value) {
    if (value >= 10) {
        return value.toString()
    }
    return value.toString().padStart(2, '0');
}

function getSummary(taskType) {
    if (taskType === 'SS') {
        return 'Standby Support';
    } else if (taskType === 'DM') {
        return 'Duty Manager';
    } else if (taskType === 'D') {
        return 'On Duty';
    } else if (taskType === 'S') {
        return 'On Standby';
    } else if (taskType === 'W') {
        return 'Weekend';
    } else if (taskType === 'F') {
        return 'Flex';
    } else if (taskType === 'BHS') {
        return 'Bank Holiday Standby';
    } else if (taskType === 'BHSS') {
        return 'Bank Holiday Standby Support';
    } else if (taskType === 'BHD') {
        return 'Bank Holiday Duty';
    } else if (taskType === 'BHW') {
        return 'Bank Holiday Weekend';
    } else {
        return taskType
    }
}


function shiftCheck(tasktype) {
    if (tasktype === 'SS' || tasktype === 'S') {
        return true
    } else if (tasktype === 'BHS' || tasktype === 'BHSS') {
        return true
    }
    return false
}

// end of month check

var xlsx = require('node-xlsx').default;

const workSheetsFromFile = xlsx.parse('DataFiles/testFile.xlsx');


function excelDateToJSDate(excelDate) {
    return new Date((excelDate - 25568)*86400*1000).toDateString();
}

let testDate = workSheetsFromFile[0]['data'][1][3];
console.log(testDate + "; " + excelDateToJSDate(testDate));

testDate = workSheetsFromFile[0]['data'][1][8];
console.log(testDate + "; " + excelDateToJSDate(testDate));

testDate = workSheetsFromFile[0]['data'][14][8];
console.log(testDate + "; " + excelDateToJSDate(testDate));

testDate = workSheetsFromFile[0]['data'][770][3];
console.log(testDate + "; " + excelDateToJSDate(testDate));


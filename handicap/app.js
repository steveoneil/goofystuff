var xlsx = require('node-xlsx').default;

const workSheetsFromFile = xlsx.parse('DataFiles/testFile.xlsx');

console.log(Object.keys(workSheetsFromFile[0]['name']));
console.log(Object.keys(workSheetsFromFile[0]['data']));
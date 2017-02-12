var XLSX = require('xlsx');

var workbook = XLSX.readFile('DataFiles/testFile.xlsx');

let scoreSheet = workbook.Sheets['Sheet1'];

let outputWB = {SheetNames: [], Sheets: {}};

outputWB.SheetNames.push('Output Sheet');
outputWB.Sheets['Output Sheet'] = scoreSheet;

XLSX.writeFile(outputWB, 'DataFiles/testFileOut2.xlsx')

// console.log(workbook.Sheets);

// const workSheetsFromFile = xlsx.parse('DataFiles/testFile.xlsx');

// function Golfer(id, name) {
//     this.id = id;
//     this.name = name;
//     this.teeSheetAppearances = 0;
//     this.scoresPosted = 0;
// }

// const scoreColumns = workSheetsFromFile[0]['data'][0];
// const scoreData = workSheetsFromFile[0]['data'].slice(1);

// const HOMECOURSE = 'St. George\'s Golf & Country Club';

// var roster = [];

// for (i = 0; i < scoreData.length; i++) {
//     let scoreEntry = scoreData[i];
//     if ((scoreEntry[9] === HOMECOURSE) && (!scoreEntry[7].includes('C'))) {
//         let existingGolfer = false;
//         for (j = 0; ((j < roster.length) && (!existingGolfer)); j++) {
//             if (scoreEntry[0] === roster[j].id) {
//                 roster[j].scoresPosted++;
//                 existingGolfer = true;
//             }
//         }
//         if (!existingGolfer) {
//             roster.push(new Golfer(scoreEntry[0], scoreEntry[2]));
//             roster[roster.length - 1].scoresPosted++;
//         }
//     }
// }

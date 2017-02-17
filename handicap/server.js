
const express = require ('express');
const app = express();

app.set('views', 'pages/views');
app.set('view engine', 'ejs');

app.use(express.static(__dirname + '/'));

app.get('/', function(req, res){
    let scoreData = readScores(inputScoreFile, inputScoreFileSheetName);
    let spr = addScores(scoreData, HOMECOURSE);
    res.render('index', {spr: spr});
});

app.listen(8080, () => {
	console.log('Server Started on http://localhost:8080');
	console.log('Press CTRL + C to stop server');
});

// --------------- app.js code ----------------------

var XLSX = require('xlsx');

const HOMECOURSE = 'St. George\'s Golf & Country Club';
const inputScoreFile = 'DataFiles/SourceFiles/scores_test.xlsx';
const inputScoreFileSheetName = 'NGN_General_Scores_Posted_By_Da';
const outputSPRFile = 'DataFiles/OutputFiles/testFileOut3.xlsx';
const outputSPRFileSheetName = 'Score Posting Report';

// Golfer object to be used to accumulate data for the Score Posting Report
function Golfer(id, name) {
    this.id = id;
    this.name = name;
    this.teeSheetAppearances = 0;
    this.scoresPosted = 0;
    this.adjustments = 0;
    this.onSPR = true;
}

// Translates cell addresses (eg. from {c:0, r:0} to 'A1')
function mapCell (C,R) {
    return XLSX.utils.encode_cell({c: C, r: R})
}

// Read in score data workbook. Returns scoreData spreadsheet object
function readScores (scoreFile, sheetName) {
    const scoreWorkbook = XLSX.readFile(scoreFile, {cellFormula: false, cellHTML: false});
    let scoreData = scoreWorkbook.Sheets[sheetName];
    return scoreData;
}

// Function that takes in scores-entered data and returns consolidated data by golfer
function addScores (scoreData, homeCourse) {

    let spr = [];
    let lastRow = XLSX.utils.decode_range(scoreData['!ref']).e.r;

    for (let R = 1; R <= lastRow; R++) {
        if ((scoreData[mapCell(9,R)].v === homeCourse) && (!scoreData[mapCell(7,R)].v.includes('C'))) {
            let existingGolfer = false;
            for (let i = 0; ((i < spr.length) && (!existingGolfer)); i++) {
                if (scoreData[mapCell(0,R)].v === spr[i].id) {
                    spr[i].scoresPosted++;
                    existingGolfer = true;
                }
            }
            if (!existingGolfer) {
                spr.push(new Golfer(scoreData[mapCell(0,R)].v, scoreData[mapCell(2,R)].v));
                spr[spr.length - 1].scoresPosted++;
            }
        }
    }
    return spr;
}

// Function that takes in consolidated golfer data and writes out report in cell-delimited format
function writeSPRReport (spr) {

    var scoreReport = { 'A1': {t: 's', v: 'Individual Id'},
                        'B1': {t: 's', v: 'Name'},
                        'C1': {t: 's', v: 'Tee Sheet Appearances'},
                        'D1': {t: 's', v: 'Scores Posted'},
                        'E1': {t: 's', v: 'Adjustments'}
    };

    for (let i = 0; i < spr.length; i++) {
        let R = i + 1;
        scoreReport[mapCell(0,R)] = {t: 'n', v: spr[i].id};
        scoreReport[mapCell(1,R)] = {t: 's', v: spr[i].name};
        scoreReport[mapCell(2,R)] = {t: 'n', v: spr[i].teeSheetAppearances};
        scoreReport[mapCell(3,R)] = {t: 'n', v: spr[i].scoresPosted};
        scoreReport[mapCell(4,R)] = {t: 'n', v: spr[i].adjustments};
    }

    scoreReport['!ref'] = XLSX.utils.encode_range({s: {c: 0, r: 0}, e: {c: 4, r: spr.length + 1}});
    return scoreReport;
}

// Prepare spreadsheet for output and write Excel file
function writeExcelFile (scoreReport) {
    let outputWB = {SheetNames: [], Sheets: {}};
    outputWB.SheetNames.push(outputSPRFileSheetName);
    outputWB.Sheets[outputSPRFileSheetName] = scoreReport;
    XLSX.writeFile(outputWB, outputSPRFile)
}

// let scoreData = readScores(inputScoreFile, inputScoreFileSheetName);
// let spr = addScores(scoreData, HOMECOURSE);
// let scoreReport = writeSPRReport(spr);
// writeExcelFile(scoreReport);



const Excel = require('exceljs');
var fs = require('fs');
const chalk = require('chalk');
const log = console.log;

var headRow;
var data;

/**
 * Function to read txt file and convert it as JSON
 * @returns txt file as json Obj
 */
function readTxt() {
    try {  
        data = fs.readFileSync('test.txt', 'utf8');
    
        let obj = {};
        let splitted = data.toString().split("\n");
       
        for (let i = 0; i<splitted.length; i++) {
            let splitLine = splitted[i].split(":");
            obj[splitLine[0].trim()] = splitLine[1].trim();
        }
        return obj
    
    } catch(e) {
        console.log('Error:', e.stack);
    }
}

// txt file converted to JSON obj
var result = readTxt();

// EXCEL SETUP
var workbook = new Excel.Workbook(); 
// confirm EXCEL WORKBOOK name
workbook.xlsx.readFile("./ProjectOne.xlsx")
    .then(function() {
        //confirm sheet name of EXCEL WORKBOOK
        var worksheet = workbook.getWorksheet("Sheet1");

        worksheet.columns = [
            { header: 'Court Case No.', key: 'Court Case No.'},
            { header: 'State Case No.', key: 'State Case No.' },
            { header: 'Name', key: 'Name' },
            { header: 'Date of Birth', key: 'Date of Birth' },
            { header: 'Date Filed', key: 'Date Filed' },
            { header: 'Date Closed', key: 'Date Closed' },
            { header: 'Warrant Type', key: 'Warrant Type' },
            { header: 'Warrant Amount', key: 'Warrant Amount' },
            { header: 'Previous Case', key: 'Previous Case' },
            { header: 'Next Case', key: 'Next Case' },
            { header: 'Assessment Amount', key: 'Assessment Amount' },
            { header: 'Balance Due', key: 'Balance Due' },
            { header: 'Stay Due Date', key: 'Stay Due Date' },
            { header: 'Judge', key: 'Judge' },
            { header: 'Defense Attorney', key: 'Defense Attorney' },
            { header: 'File Section', key: 'File Section' },
            { header: 'File Location', key: 'File Location' },
            { header: 'Box No', key: 'Box No' },
            { header: 'Probation Start Date', key: 'Probation Start Date' },
            { header: 'Probation End Date', key: 'Probation End Date' },
            { header: 'Probation Length', key: 'Probation Length' },
            { header: 'Probation Type', key: 'Probation Type' },
            { header: 'Defendant in Jail', key: 'Defendant in Jail' },
            { header: 'Defendant Release to', key: 'Defendant Release to' },
            { header: 'Bond Amount', key: 'Bond Amount' },
            { header: 'Bond Status', key: 'Bond Status' },
            { header: 'Bond Type', key: 'Bond Type' },
            { header: 'Bond Issue Date', key: 'Bond Issue Date' },
            { header: 'Seq No.', key: 'Seq No.' },
            { header: 'Charge', key: 'Charge' },
            { header: 'Charge Type', key: 'Charge Type' },
            { header: 'Disposition', key: 'Disposition' },
            { header: 'Sentence', key: 'Sentence' },
          ];

        // to add missing header values
        headRow = worksheet.getRow(1).values;
        for(i=1;i<headRow.length;i++){
            if(!result.hasOwnProperty(headRow[i])){
                result[headRow[i]] = "";
            }
        }

        const newRow = worksheet.insertRow(2,result)
        if(newRow.cellCount>0){
            log( chalk.green("\nSUCCESSFULLY ROW INSERTED..."));
            workbook.xlsx.writeFile('ProjectOne.xlsx')
        }
});









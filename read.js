const ExcelJS = require('exceljs');
const workbook = new ExcelJS.Workbook(); 
//var filename = './xls/template.xlsx';
var filename = './xls/template.xlsx';
var sheet = 'Sheet1';
var label_1 = 'Box No. :';
var tagLABEL;

try{
    workbook.xlsx.readFile(filename).then(function() {
        //let workSheet = workbook.getWorksheet(sheet);
        let obj = new Label(workbook,sheet);
        obj.setLabel(8,4);
        obj.getLabel(10);
    });
}catch{
    console.log('error');
}

class Label{
    ROWS = new Array();

    constructor(workbook,sheetName){
        this.workbook = workbook;
        this.sheetName = sheetName;
    }

    setLabel(rowNum,colNum){
        rowNum++;
        colNum++;

        let workSheet = this.workbook.getWorksheet(this.sheetName);
        let ROWS = new Array();

        for(let y=1;y < rowNum;y++){
            let COLUMS = new Array();
            let row = workSheet.getRow(y);

            for(let x = 1;x < colNum;x++){                
                let CELL = row.getCell(x);
                COLUMS.push(CELL);
            }
            ROWS.push(COLUMS);
        }
        //console.log(ROWS);
        this.tagLABEL = ROWS;
    }

    getLabel(posStart){
        let ROWS = this.tagLABEL;
        let workSheet = this.workbook.getWorksheet(this.sheetName);

        for(let y=0;y < ROWS.length;y++){
            let COLUMS = ROWS[y];
            let saveRow = workSheet.getRow(y+posStart);

            for(let x = 0;x < COLUMS.length;x++){                
                process.stdout.write(COLUMS[x].value +"|");
                //workSheet.getRow(y+posStart).getCell(x).value = COLUMS[x].value;
                //let saveCell = saveRow.getCell(x);
                //saveCell.value = COLUMS[x].value;
            }
            saveRow.commit();
            console.log("------------------");
        }
        //Finally creating XLSX file
        //this.workbook.xlsx.writeFile('./xls/new.xlsx');
    }
}
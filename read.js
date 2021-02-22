const ExcelJS = require('exceljs');

const workbook = new ExcelJS.Workbook(); 
//var filename = './xls/template.xlsx';
var filename = './xls/template.xlsx';
var sheet = 'Sheet1';

var border_1 = {
    top: {style:'medium', color: {argb:'00000000'}},
    left: {style:'medium', color: {argb:'00000000'}},
    bottom: {style:'medium', color: {argb:'00000000'}},
    right: {style:'medium', color: {argb:'00000000'}}
  };

try{
    workbook.xlsx.readFile(filename).then(function() {
        //let workSheet = workbook.getWorksheet(sheet);
        let obj = new Label(workbook,sheet);
        obj.setLabel(8,4);                  // original Tempplate(RowNumber,columnNumber)
        obj.getLabel(10,300);                 // (positionStart, number of copy)
    });
}catch{
    console.log('error');
}
//----------------------------------------------------------
//
//----------------------------------------------------------
class Label{
    ROWS = new Array();
    label_1 = 'L3 / Box No. : ';
    label_2 = 13336;
    label_3 = 13385;
    label_4 = 3080321001;

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

    getLabel(posStart,copyNumber){
        let ROWS = this.tagLABEL;
        let workSheet = this.workbook.getWorksheet(this.sheetName);

        for(let z = 1; z <= copyNumber; z++){
            let nextPos = posStart * z ;
            for(let y=0;y < ROWS.length;y++){
                let COLUMS = ROWS[y];
                let saveRow = workSheet.getRow(y+nextPos);
                //console.log(workSheet.getRow(y+posStart).height);

                for(let x = 0;x < COLUMS.length;x++){           
                    //process.stdout.write(COLUMS[x].value +"|");
                    let saveCell    = saveRow.getCell(x+1);
                    saveCell.value  = COLUMS[x].value;
                    saveCell.style  = COLUMS[x].style;
                    saveCell.height = COLUMS[x].height;
                    if(y===0 && x ===1){
                        saveCell.value  = this.label_1+(z+1);
                    }else if(y===4 && x ===1){
                        saveCell.value  = this.label_2+(50*z);
                    }else if(y===4 && x ===2){
                        saveCell.value  = this.label_3+(50*z);
                    }else if(y===6 && x ===1){
                        let label = "ID : LISH"+ (parseInt(this.label_4)+(z));
                        saveCell.value  = label;
                    }
                }
                if(y===2){
                    let a = y+nextPos;
                    workSheet.mergeCells('B'+a+':C'+a);
                    workSheet.getCell('B'+a+':C'+a).border  = border_1;
                }
                saveRow.height = workSheet.getRow(y+1).height;
                saveRow.commit();                
            }
            console.log("\n------------------ copy"+z);
        }
        //Finally creating XLSX file
        this.workbook.xlsx.writeFile('./xls/new.xlsx');
    }
}

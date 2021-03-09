const ExcelJS = require('exceljs');
const Label = require('./Label.js');

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
console.log('test');

import { Injectable } from '@angular/core';
import {saveAs} from 'file-saver'; 
import {  WorkBook, read, utils, write, readFile } from 'xlsx';
@Injectable()
export class FileExcelService {

  wbout = [];
  table = [];
   ws: any;
  constructor() { 
    this.setExcelProperties('');
  }

  s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i !== s.length; ++i) {
      view[i] = s.charCodeAt(i) & 0xFF;
    };
    return buf;
  }


  SaveToExcel(tableData, fileName: string= 'QuestionSheet') {
    this.setTableData(tableData,fileName);
    saveAs(new Blob([this.s2ab(this.wbout)], { type: 'application/octet-stream' }), fileName+'.xlsx');
   }

   getTableData()
   {
     return this.table;
   }
   
   setTableData(tableData, fileName: string)
   {
     this.table = tableData;
     this.setExcelProperties(fileName);
   }


   // excel Detail
   setExcelProperties(fileName: string)
   {
    const ws_name = fileName.substr(0, 25); //'QuestionSheet'
  //  const ws_name = ''; // worksheet name cannot exceed 31 chracters length
    const wb: WorkBook = { SheetNames: [], Sheets: {} };
    this.ws = utils.json_to_sheet(this.getTableData());
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = this.ws;
     this.wbout = write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });
   }
   
   convertExcelToJson(file)
   {
    let reader = new FileReader();
    let workbookkk;
    let XL_row_object;
    let json_object;
    reader.readAsBinaryString(file);
    return new Promise((resolve, reject) => {
      reader.onload = function(){
        //  alert(reader.result);
        let data = reader.result;
         workbookkk=read(data,{type: 'binary'});
         console.log(workbookkk);
         workbookkk.SheetNames.forEach(function(sheetName) {
          // Here is your object
           XL_row_object = utils.sheet_to_json(workbookkk.Sheets[sheetName]);
           json_object = JSON.stringify(XL_row_object);
        //  console.log(json_object);
        //  console.log(XL_row_object);
            resolve(XL_row_object);
        });
        };
    });
    }
}

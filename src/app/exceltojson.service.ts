import { Injectable } from '@angular/core';
import {DomSanitizer } from '@angular/platform-browser'
import * as XLSX from 'xlsx'

@Injectable({
  providedIn: 'root'
})
export class ExceltojsonService {

  get_json_data = "";
  supported_formats = [".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xls", ".xlt", ".xls"]

  constructor(private sanitizer: DomSanitizer) { }

  convertExcelToJson(file){
    if (file){
      let isValid = file.name.endsWith(this.supported_formats.map(ext =>{
        return ext
      }));
      if (isValid === true){
        let work_book = null;
        const reader = new FileReader();
  
        reader.onload = (e) => {
          let data = reader.result;
          let rowObj = {};
          let key = null;
    
  
          work_book = XLSX.read(data, { type: 'binary'});
          let get_strings = work_book.Strings;
          get_strings.forEach(element => {
            key = element.t.toLowerCase();
            key = key.split(" ");
            key = key.join("_");
            rowObj[key] = element.t;
          });
          this.get_json_data = JSON.stringify(rowObj);
        }
        reader.readAsBinaryString(file);
      }
    }
  }

  downloadJsonData(){
    return this.sanitizer.bypassSecurityTrustUrl(`data:text/json;charset=utf-8,${encodeURIComponent(this.get_json_data)}`)
  }
}

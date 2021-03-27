import { Injectable } from '@angular/core';
import {DomSanitizer } from '@angular/platform-browser'
import * as XLSX from 'xlsx'

@Injectable({
  providedIn: 'root'
})
export class ExceltojsonService {

  get_json_data = "";
  supported_formats = ["xlsx", "xlsm", "xlsb", "xltx", "xltm", "xls", "xlt", "xls"]
  alert_message = null;

  constructor(private sanitizer: DomSanitizer) { }

  convertExcelToJson(file){
    if (file){
      let extension = file.name.split(".")[1];
      let isValid = this.supported_formats.includes(extension);
      if (isValid === true){
        this.alert_message = null;
        const reader = new FileReader();
  
        reader.onload = (e) => {
          let data = reader.result;
          let work_book = null;
  
          work_book = XLSX.read(data, { type: 'binary'});
          let excelInputs = this.getExcelInputs(work_book);
          this.get_json_data = JSON.stringify(excelInputs, undefined, 4);
        }
        reader.readAsBinaryString(file);
      }
      else this.alert_message = "Your file format is not supported. Only excel formats are supported!"
    }
  }

  downloadJsonData(){
    return this.sanitizer.bypassSecurityTrustUrl(`data:text/json;charset=utf-8,${encodeURIComponent(this.get_json_data)}`)
  }

  getExcelInputs(work_book){
    console.log("work_book", work_book);
    let sheetname = work_book.SheetNames;
    let Obj = {};
    let a19 = work_book.Sheets[sheetname].A19.v;
    let b3 = work_book.Sheets[sheetname].B3.v;
    let b4 = work_book.Sheets[sheetname].B4.v;
    let b5 = work_book.Sheets[sheetname].B5.v
    let b6 = work_book.Sheets[sheetname].B6.v;
    let b7 = work_book.Sheets[sheetname].B7.v;
    let b8 = work_book.Sheets[sheetname].B8.v;
    let b10 = work_book.Sheets[sheetname].B10.v;
    let b12 = work_book.Sheets[sheetname].B12.v;
    let b13 = work_book.Sheets[sheetname].B13.v;
    let b14 = work_book.Sheets[sheetname].B14.v;
    let b15 = work_book.Sheets[sheetname].B15.v;
    let b16 = work_book.Sheets[sheetname].B16.v;
    let b19 = work_book.Sheets[sheetname].B19.v;
    let d19 = work_book.Sheets[sheetname].D19.v;
    let e19 = work_book.Sheets[sheetname].E19.v;
    let f19 = work_book.Sheets[sheetname].F19.v;
    let d10 = work_book.Sheets[sheetname].D10.v;
    let d12 = work_book.Sheets[sheetname].D12.v;
    let d13 = work_book.Sheets[sheetname].D13.v;
    let d14 = work_book.Sheets[sheetname].D14.v;
    let d15 = work_book.Sheets[sheetname].D15.v;
    let e9 = work_book.Sheets[sheetname].E9.v;
    let f4 = work_book.Sheets[sheetname].F4.v;
    let f5 = work_book.Sheets[sheetname].F5.w;
    let f6 = work_book.Sheets[sheetname].F6.v;
    let f7 = work_book.Sheets[sheetname].F7.v;
    // let b24 = work_book.Sheets[sheetname].B24.v;
    let b25 = work_book.Sheets[sheetname].B25.v;
    let e24 = work_book.Sheets[sheetname].E24.v;
    let e25 = work_book.Sheets[sheetname].E25.v;
    let e26 = work_book.Sheets[sheetname].E26.v;
    let e27 = work_book.Sheets[sheetname].E27.v;
    let e28 = work_book.Sheets[sheetname].E28.v;
    let e29 = work_book.Sheets[sheetname].E29.v;
    let e30 = work_book.Sheets[sheetname].E30.v;
    let f24 = work_book.Sheets[sheetname].F24.v;
    let f25 = work_book.Sheets[sheetname].F25.v;
    let f26 = work_book.Sheets[sheetname].F26.v;
    let f27 = work_book.Sheets[sheetname].F27.v;
    let f28 = work_book.Sheets[sheetname].F28.v;
    let f29 = work_book.Sheets[sheetname].F29.v;
    let f30 = work_book.Sheets[sheetname].F30.v;


    let item_number_1 = work_book.Sheets[sheetname].D20.v
    let descr_row_1 =  work_book.Sheets[sheetname].B20.v;
    let quantity_row_1 = work_book.Sheets[sheetname].D20.v;
    let unit_price_row_1 = work_book.Sheets[sheetname].E20.v;
    let total_row_1 = work_book.Sheets[sheetname].F20.v;

    let item_number_2 = work_book.Sheets[sheetname].D21.v
    let descr_row_2 = work_book.Sheets[sheetname].B21.v;
    let quantity_row_2 = work_book.Sheets[sheetname].D21.v;
    let unit_price_row_2 = work_book.Sheets[sheetname].E21.v;
    let total_row_2 = work_book.Sheets[sheetname].F21.v;

    let item_number_3 = work_book.Sheets[sheetname].D22.v 
    let descr_row_3 = work_book.Sheets[sheetname].B22.v;
    let quantity_row_3 = work_book.Sheets[sheetname].D22.v;
    let unit_price_row_3 = work_book.Sheets[sheetname].E22.v;
    let total_row_3 = work_book.Sheets[sheetname].F22.v;

    let item_number_4 = work_book.Sheets[sheetname].D23.v
    let descr_row_4 = work_book.Sheets[sheetname].B23.v;
    let quantity_row_4 = work_book.Sheets[sheetname].D23.v;
    let unit_price_row_4 = work_book.Sheets[sheetname].E23.v;
    let total_row_4 = work_book.Sheets[sheetname].F23.v;




    // let row5_qantity = work_book.Sheets[sheetname].D24.v;
    // let row6_qantity = work_book.Sheets[sheetname].D25.v;
    // let row7_qantity = work_book.Sheets[sheetname].D26.v;
    // let row8_qantity = work_book.Sheets[sheetname].D27.v;
    // let row9_qantity = work_book.Sheets[sheetname].D28.v;
    // let row10_qantity = work_book.Sheets[sheetname].D29.v;
    // let row11_qantity = work_book.Sheets[sheetname].D30.v;

    Obj["invoice_title"] =  b3;
    Obj["date"] =  {
      "field_name": f4,
      "value": f5
    };
    Obj["invoice_number"] =  {
      "field_name": f6,
      "value": f7
    };

    // Obj[f4] =  f5;
    // Obj[f6] =  f7;
    Obj["comment"] =  e9;

    // Company information
    Obj["company_information"] = {
      "company_name": b4,
      "address":b5,
      "city": b6,
      "phone_number": b7,
      "email_address": b8
    }

  // Customer information (BILL TO)
    Obj["bill_to"] = {
      "header_title": b10,
      "name": b12,
      "company_name": b13,
      "address": b14,
      "phone_number": b15,
      "email_address": b16
    }
    // SHIP TO  
    Obj["ship_to"] = {
      "header_title": d10,
      "name": d12,
      "company_name": d13,
      "address": d14,
      "phone_number": d15
    }

    Obj["headers"] = {
      "item_number": a19,
      "description": b19,
      "qantity": d19,
      "unit_price": e19,
      "total": f19
    };

    // Obj["description"] = b19;
    // Obj["qantity"] = d19;
    // Obj["unit_price"] = e19;
    // Obj["total"] = f19;

    Obj["row_1"] = {
      "item_number": item_number_1,
      "description": descr_row_1,
      "quantity": quantity_row_1,
      "unit_price": unit_price_row_1,
      "total": total_row_1
    }

    Obj["row_2"] = {
      "item_number": item_number_2,
      "description": descr_row_2,
      "quantity": quantity_row_2,
      "unit_price": unit_price_row_2,
      "total": total_row_2
    }

    Obj["row_3"] = {
      "item_number": item_number_3,
      "description": descr_row_3,
      "quantity": quantity_row_3,
      "unit_price": unit_price_row_3,
      "total": total_row_3
    }

    Obj["row_4"] = {
      "item_number": item_number_4,
      "description": descr_row_4,
      "quantity": quantity_row_4,
      "unit_price": unit_price_row_4,
      "total": total_row_4
    }
    Obj["remarks_payment_instructions"] = b25;
    //  {
    //   "1": b24,
    //   "2": b25
    // };
    // Obj["payment_instructions"] =  b24;
    // Obj["last_comment"] =  b25;

    Obj["sub_total"] = {
      "field_name": e24,
      "value": f24
    }
    Obj["discount"] = {
      "field_name": e25,
      "value": f25
    }
    Obj["sub_total_discount"] = {
      "field_name": e26,
      "value": f26
    }
    Obj["tax_rate"] = {
      "field_name": e27 + '(%)',
      "value": f27
    }
    Obj["total_tax"] = {
      "field_name": e28,
      "value": f28
    }
    Obj["shipping_handling"] = {
      "field_name": e29,
      "value": f29
    }
    Obj["balance_due"] = {
      "field_name": e30,
      "value": f30
    }
    // Obj[e24] =  f24;
    // Obj[e25] =  f25;
    // Obj[e26] =  f26;
    // Obj[e27 + '(%)'] =  f27;
    // Obj[e28] =  f28;
    // Obj[e29] =  f29;
    // Obj[e30] =  f30;

    return Obj
  }
}

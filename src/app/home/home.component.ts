import { Component, OnInit } from '@angular/core';
import { ExceltojsonService } from '../exceltojson.service';



@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css']
})
export class HomeComponent implements OnInit {
  
  file = "";
  file_name = null;

  constructor( public exceltojson:ExceltojsonService) { }

  ngOnInit(): void {
  }
  
  getFile(event){
    this.file = event.target.files[0];
    this.file_name = event.target.files[0].name
  }

  generateJsonFile(){
    this.exceltojson.convertExcelToJson(this.file);
  }

  downloadNow(){
    this.exceltojson.downloadJsonData();
  }
}
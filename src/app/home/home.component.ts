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
  
  // This func gets the file uploaded
  getFile(event){
    this.file = event.target.files[0];
    this.file_name = event.target.files[0].name;
    this.file_name = this.file_name.split(".")[0];
  }

  // This func calls the convert func from the service
  generateJsonFile(){
    this.exceltojson.convertExcelToJson(this.file);
  }

// This func calls the download func from the service
  downloadNow(){
    this.exceltojson.downloadJsonData();
  }
}

import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterOutlet } from '@angular/router';
import * as XLSX from 'xlsx';
// ติดตั้ง npm install xlsx --save
@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, RouterOutlet],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'Export to Excel';

  employees = [
    { FirstName: 'สมชาย', LastName: 'ใจดี', Age: 29, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'สมหญิง', LastName: 'สวยงาม', Age: 30, gender: 'หญิง', prefix: 'นาง' },
    { FirstName: 'สิริรัตน์', LastName: 'สุขใจ', Age: 28, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'สุรชัย', LastName: 'ใจกว้าง', Age: 35, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'วรรณศรี', LastName: 'มีความสุข', Age: 26, gender: 'หญิง', prefix: 'นาง' },
    { FirstName: 'สมปรารถนา', LastName: 'ใจสะอาด', Age: 32, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'สมหญิง', LastName: 'หน้าใส', Age: 29, gender: 'หญิง', prefix: 'นาง' },
    { FirstName: 'สุขสวัสดิ์', LastName: 'มีความสุข', Age: 33, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'สุชาติ', LastName: 'รอบรู้', Age: 27, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'วรรณรท', LastName: 'เตรียมตัว', Age: 31, gender: 'หญิง', prefix: 'นาง' },
    { FirstName: 'สมปราชญ์', LastName: 'ใจเย็น', Age: 28, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'สมหญิง', LastName: 'มีรอยยิ้ม', Age: 29, gender: 'หญิง', prefix: 'นาง' },
    { FirstName: 'วิไลวรรณ', LastName: 'น่ารัก', Age: 34, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'สมชาย', LastName: 'ใจดี', Age: 25, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'สมหญิง', LastName: 'สวยงาม', Age: 30, gender: 'หญิง', prefix: 'นาง' },
    { FirstName: 'สิริรัตน์', LastName: 'สุขใจ', Age: 28, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'สุรชัย', LastName: 'ใจกว้าง', Age: 35, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'วรรณศรี', LastName: 'มีความสุข', Age: 26, gender: 'หญิง', prefix: 'นาง' },
    { FirstName: 'สมปรารถนา', LastName: 'ใจสะอาด', Age: 32, gender: 'ชาย', prefix: 'นาย' },
    { FirstName: 'สมชาย', LastName: 'ใจดี', Age: 29, gender: 'ชาย', prefix: 'นาย' },
  ];

  /* Default name for Excel file when downloaded - ชื่อไฟล์ที่จะโหลด */
  fileName = 'ExcelSheet.xlsx';

  exportexcel() {
    /** passing table id **/
    let data = document.getElementById('table-data');
    const ws: XLSX.WorkSheet =XLSX.utils.table_to_sheet(data);

    /* generate workbook and add the worksheet - สร้าง workbook และเพิ่ม worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1'); // sheet name - ชื่อ sheet

    /* save to file - บันทึกเป็นไฟล์ */
    XLSX.writeFile(wb, this.fileName);

    // how to Credit : https://www.youtube.com/watch?v=j2gQArYvgw0
  }
}

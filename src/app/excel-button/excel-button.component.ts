import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ExcelExportService } from '../excel-export.service';

@Component({
  selector: 'app-excel-button',
  standalone: true,
  imports: [CommonModule],
  template: `
    <button 
      class="px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 disabled:bg-gray-400"
      [disabled]="isExporting"
      (click)="exportToExcel()">
      <span *ngIf="!isExporting">Export to Excel</span>
      <span *ngIf="isExporting">Exporting...</span>
    </button>
  `
})
export class ExcelButtonComponent {
  isExporting = false;

  constructor(private excelService: ExcelExportService) {}

  async exportToExcel() {
    this.isExporting = true;
    try {
      const data = [
        { name: 'John Doe', amount: 1234.56, date: new Date('2024-01-15') },
        { name: 'Jane Smith', amount: 7890.12, date: new Date('2024-02-20') }
      ];

      await this.excelService.exportToExcel(data, 'data-export');
    } finally {
      this.isExporting = false;
    }
  }
}

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, ExcelButtonComponent],
  template: `
    <div class="container mx-auto p-4">
      <h1 class="text-2xl font-bold mb-4">XLS Export Demo</h1>
      <app-excel-button></app-excel-button>
    </div>
  `
})
export class AppComponent {
  title = 'pdf-export-demo';
}

import { CommonModule } from '@angular/common';
// pdf-export.service.ts
import { Component, Injectable } from '@angular/core';
import { ExcelButtonComponent } from './excel-button/excel-button.component';


@Injectable({
  providedIn: 'root'
})
export class PdfExportService {
  constructor() { }
}
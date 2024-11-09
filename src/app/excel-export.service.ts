import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

@Injectable({
  providedIn: 'root'
})
export class ExcelExportService {
  constructor() {}

  async exportToExcel(data: any[], fileName: string) {
    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');

    // Get headers from the first data item
    const headers = Object.keys(data[0] || {});

    // Define columns with proper formatting
    worksheet.columns = headers.map(header => ({
      header: this.capitalizeHeader(header),
      key: header,
      width: 15, // Default width
      style: {
        font: { name: 'Arial', size: 11 },
        alignment: { vertical: 'middle' }
      }
    }));

    // Style the header row
    worksheet.getRow(1).font = { bold: true, size: 12 };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4F81BD' }
    };
    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

    // Add data rows with proper formatting
    data.forEach(item => {
      const row = worksheet.addRow(item);
      
      // Apply formatting to specific column types
      headers.forEach((header, index) => {
        const cell = row.getCell(index + 1);
        const value = item[header];

        if (this.isDateValue(value)) {
          // Format dates
          cell.value = new Date(value);
          cell.numFmt = 'dd/mm/yyyy';
        } else if (this.isNumberValue(value)) {
          // Format numbers
          cell.value = Number(value);
          cell.numFmt = '#,##0.00';
          cell.alignment = { horizontal: 'right' };
        }
      });

      // Add alternating row colors
      if (row.number % 2 === 0) {
        row.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF2F2F2' }
        };
      }
    });

    // Add totals row if there are numeric columns
    this.addTotalsRow(worksheet, headers, data);

    // Add borders to all cells
    worksheet.eachRow(row => {
      row.eachCell(cell => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    // Auto-filter for all columns
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: headers.length }
    };

    // Freeze the header row
    worksheet.views = [
      { state: 'frozen', xSplit: 0, ySplit: 1 }
    ];

    // Generate and save the file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    saveAs(blob, `${fileName}.xlsx`);
  }

  private capitalizeHeader(header: string): string {
    return header
      .split(/(?=[A-Z])|_/)
      .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
      .join(' ');
  }

  private isDateValue(value: any): boolean {
    if (value instanceof Date) return true;
    if (typeof value === 'string') {
      const date = new Date(value);
      return date instanceof Date && !isNaN(date.getTime());
    }
    return false;
  }

  private isNumberValue(value: any): boolean {
    if (typeof value === 'number') return true;
    if (typeof value === 'string') {
      return !isNaN(parseFloat(value)) && isFinite(Number(value));
    }
    return false;
  }

  private addTotalsRow(worksheet: ExcelJS.Worksheet, headers: string[], data: any[]) {
    const numericColumns = headers.filter(header => 
      data.some(row => this.isNumberValue(row[header]))
    );

    if (numericColumns.length > 0) {
      const totalRow = worksheet.addRow({});
      
      // Style the totals row
      totalRow.font = { bold: true };
      totalRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6E6' }
      };

      // Add "Totals" label
      totalRow.getCell(1).value = 'Totals';

      // Add sum formulas for numeric columns
      numericColumns.forEach(header => {
        const colIndex = headers.indexOf(header) + 1;
        const lastDataRow = worksheet.rowCount - 1;
        const cell = totalRow.getCell(colIndex);
        
        // Excel column letter
        const colLetter = worksheet.getColumn(colIndex).letter;
        
        // SUM formula excluding the header and totals row
        cell.value = {
          formula: `SUM(${colLetter}2:${colLetter}${lastDataRow})`
        };
        cell.numFmt = '#,##0.00';
        cell.alignment = { horizontal: 'right' };
      });
    }
  }
}
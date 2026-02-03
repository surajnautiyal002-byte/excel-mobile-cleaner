import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.html',
  styleUrls: ['./app.css']
})
export class AppComponent {

  fileName = '';
  workbook: XLSX.WorkBook | null = null;

  sheetNames: string[] = [];
  selectedSheet = '';

  rawData: any[][] = [];
  previewData: any[][] = [];

  showPreview = false;

  headerRowIndex: number | null = null;
  headers: string[] = [];
  selectedColumns: number[] = [];

  /* UX + Performance */
  isDragging = false;
  isProcessing = false;
  progress = 0;

  /* ================= FILE UPLOAD ================= */

  onFileChange(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    this.fileName = file.name.replace(/\.[^/.]+$/, '');
    this.workbook = null;
    this.showPreview = false;
    this.isProcessing = false;
    this.progress = 0;

    const reader = new FileReader();
    reader.onload = (e: any) => {
      const data = new Uint8Array(e.target.result);
      this.workbook = XLSX.read(data, { type: 'array' });
      this.sheetNames = this.workbook.SheetNames;
      this.selectedSheet = this.sheetNames[0];
    };
    reader.readAsArrayBuffer(file);
  }

  /* ================= DRAG & DROP ================= */

  onDragOver(event: DragEvent) {
    event.preventDefault();
    this.isDragging = true;
  }

  onDragLeave(event: DragEvent) {
    event.preventDefault();
    this.isDragging = false;
  }

  onDrop(event: DragEvent) {
    event.preventDefault();
    this.isDragging = false;

    if (!event.dataTransfer || !event.dataTransfer.files.length) return;

    const file = event.dataTransfer.files[0];

    if (!file.name.match(/\.(xls|xlsx)$/i)) {
      alert('Please drop a valid Excel file (.xls or .xlsx)');
      return;
    }

    this.onFileChange({ target: { files: [file] } });
  }

  /* ================= PREVIEW ================= */

  previewSheet() {
    if (!this.workbook) {
      alert('File is still loading. Please click Preview again.');
      return;
    }

    const sheet = this.workbook.Sheets[this.selectedSheet];
    this.rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    this.previewData = this.rawData.slice(0, 20);

    this.headerRowIndex = null;
    this.headers = [];
    this.selectedColumns = [];
    this.showPreview = true;
  }

  /* ================= HEADER ROW ================= */

  setHeaderRow(index: number) {
    this.headerRowIndex = index;
    this.headers = this.rawData[index].map(h => String(h || '').trim());
    this.selectedColumns = [];

    for (let col = 0; col < this.headers.length; col++) {
      let validCount = 0;

      for (let row = index + 1; row < Math.min(index + 10, this.rawData.length); row++) {
        if (this.cleanMobile(this.rawData[row][col])) validCount++;
      }

      if (validCount >= 3) this.selectedColumns.push(col);
    }
  }

  toggleColumn(index: number) {
    this.selectedColumns.includes(index)
      ? this.selectedColumns = this.selectedColumns.filter(i => i !== index)
      : this.selectedColumns.push(index);
  }

  /* ================= CLEAN & DOWNLOAD (CHUNKED) ================= */

  async cleanAndDownload() {
    if (this.headerRowIndex === null || this.selectedColumns.length === 0) {
      alert('Select header row and at least one mobile column');
      return;
    }

    if (this.isProcessing) return;

    this.isProcessing = true;
    this.progress = 0;

    const cleaned: any[][] = [];
    cleaned.push(this.headers);

    const totalRows = this.rawData.length - (this.headerRowIndex + 1);
    const chunkSize = 500;
    let processed = 0;

    for (let start = this.headerRowIndex + 1; start < this.rawData.length; start += chunkSize) {
      const end = Math.min(start + chunkSize, this.rawData.length);

      for (let i = start; i < end; i++) {
        const row = [...this.rawData[i]];
        const validMobiles: string[] = [];
        const perColumn: (string | null)[] = [];

        for (const col of this.selectedColumns) {
          const cleanedMobile = this.cleanMobile(row[col]);
          perColumn.push(cleanedMobile);
          if (cleanedMobile) validMobiles.push(cleanedMobile);
        }

        if (validMobiles.length === 0) {
          processed++;
          continue;
        }

        const fallback = validMobiles[0];

        this.selectedColumns.forEach((col, idx) => {
          row[col] = perColumn[idx] ?? fallback;
        });

        cleaned.push(row);
        processed++;
      }

      this.progress = Math.round((processed / totalRows) * 100);
      await this.yieldToBrowser();
    }

    const ws = XLSX.utils.aoa_to_sheet(cleaned);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Cleaned');

    const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([buffer]), `C_${this.fileName}.xlsx`);

    this.isProcessing = false;
    this.progress = 100;
  }

  private yieldToBrowser(): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, 0));
  }

  /* ================= MOBILE CLEANER ================= */

  cleanMobile(value: any): string | null {
    if (!value) return null;

    const matches = String(value).match(/\d{10,12}/g);
    if (!matches) return null;

    for (let num of matches) {
      if (num.length === 12 && num.startsWith('91')) num = num.slice(2);

      if (
        num.length === 10 &&
        /^[6-9]\d{9}$/.test(num) &&
        !/^(\d)\1{9}$/.test(num)
      ) {
        return `+91${num}`;
      }
    }

    return null;
  }
}

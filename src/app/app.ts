import { Component, NgZone, ChangeDetectorRef, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import type { WorkBook, WorkSheet } from 'xlsx';
import type JSZip from 'jszip';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.html',
  styleUrls: ['./app.css']
})
export class AppComponent implements OnInit, OnDestroy {

  fileName = '';
  workbook: WorkBook | null = null;

  sheetNames: string[] = [];
  selectedSheet = '';

  rawData: any[][] = [];
  previewData: any[][] = [];

  showPreview = false;

  headerRowIndex: number | null = null;
  headers: string[] = [];
  selectedColumns: number[] = [];
  csvStreamingFile: File | null = null;
  csvDelimiter = ',';

  /* UX + Performance */
  isDragging = false;
  isProcessing = false;
  progress = 0;
  errorMessage = '';
  successMessage = '';
  isUploading = false;
  uploadProgress = 0;

  /* Statistics */
  stats = {
    total: 0,
    valid: 0,
    duplicates: 0,
    invalidPattern: 0,
    invalidLength: 0
  };
  showStats = false;
  statDownloads = {
    valid: [] as Array<{ row: number; column: string; original: string; cleaned: string }>,
    duplicates: [] as Array<{ row: number; mobile: string }>,
    invalidPattern: [] as Array<{ row: number; column: string; value: string }>,
    invalidLength: [] as Array<{ row: number; column: string; value: string }>
  };

  /* Export Options */
  exportMode: 'full' | 'unique' | 'mobile-name' | 'keep-all' = 'full';
  selectedNameColumn: number | null = null;

  private readonly MAX_FILE_SIZE = 500 * 1024 * 1024; // 500MB
  private readonly MAX_XLSX_BROWSER_SAFE_SIZE = 80 * 1024 * 1024; // 80MB
  private readonly CSV_STREAM_CHUNK_SIZE = 4 * 1024 * 1024; // 4MB
  private readonly MAX_STAT_DOWNLOAD_ROWS = 10000;
  private readonly MAX_SHEET_ROWS = 2000000;
  private readonly MAX_SHEET_COLUMNS = 500;
  private readonly MAX_SHEET_CELLS = 50000000;
  private readonly LARGE_SHEET_WARNING_ROWS = 250000;
  private readonly FAST_MODE_ROWS = 100000;
  private readonly MAX_PREVIEW_ROWS = 50;
  private readonly MIN_VALID_MOBILES = 3;
  private readonly INVALID_PATTERNS = [
    /^(\d)\1{9}$/,           // All same digits
    /^0123456789$/,          // Sequential
    /^1234567890$/,
    /^9876543210$/
  ];

  private boundHandlePaste = this.handlePaste.bind(this);
  private errorTimer: ReturnType<typeof setTimeout> | null = null;
  private successTimer: ReturnType<typeof setTimeout> | null = null;
  private xlsxModule: typeof import('xlsx') | null = null;
  private jsZipModule: typeof JSZip | null = null;
  private DEBUG = false;

  constructor(private ngZone: NgZone, private cdr: ChangeDetectorRef) {}

  ngOnInit() {
    // Listen for paste events globally
    document.addEventListener('paste', this.boundHandlePaste);
  }

  ngOnDestroy() {
    document.removeEventListener('paste', this.boundHandlePaste);
    if (this.errorTimer) clearTimeout(this.errorTimer);
    if (this.successTimer) clearTimeout(this.successTimer);
  }

  private handlePaste(event: ClipboardEvent) {
    const items = event.clipboardData?.items;
    if (!items) return;

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.kind === 'file') {
        const file = item.getAsFile();
        if (file) {
          event.preventDefault();
          this.ngZone.run(() => {
            this.onFileChange({ target: { files: [file] } });
            this.cdr.detectChanges();
          });
          return;
        }
      }
    }
  }

  /* ================= FILE UPLOAD ================= */

  onFileChange(event: Event | { target: { files: File[] } }) {
    const target = (event.target || (event as any).target) as HTMLInputElement;
    this.clearMessages();
    const file = target.files?.[0] || ((event as any).target?.files?.[0]);
    if (!file) return;

    if (file.size > this.MAX_FILE_SIZE) {
      this.showError(`File too large. Maximum size is ${this.MAX_FILE_SIZE / 1024 / 1024}MB`);
      target.value = '';
      return;
    }

    if (!this.isValidExcelFile(file)) {
      this.showError('Please select a supported spreadsheet file (.xls, .xlsx, .xlsm, .xlsb, .csv, .ods)');
      target.value = '';
      return;
    }

    this.fileName = file.name.replace(/\.[^/.]+$/, '').replace(/[<>:"/\\|?*]/g, '_');
    this.resetState();
    this.isUploading = true;
    this.uploadProgress = 0;

    // Simulate progress for better UX
    const progressInterval = setInterval(() => {
      if (this.uploadProgress < 90) {
        this.uploadProgress += 10;
        this.cdr.detectChanges();
      }
    }, 100);

    const extMatch = (file.name.match(/\.[^/.]+$/) || [''])[0].toLowerCase();
    const ext = extMatch;

    if (ext === '.xlsx' && file.size > this.MAX_XLSX_BROWSER_SAFE_SIZE) {
      this.showError(
        `This XLSX is very large (${(file.size / 1024 / 1024).toFixed(1)}MB) and may freeze the browser. ` +
        'Please upload CSV for this dataset, or split the Excel file into smaller parts.'
      );
      if ('value' in target) target.value = '';
      this.isUploading = false;
      return;
    }

    if (ext === '.csv') {
      this.ngZone.run(async () => {
        try {
          const preview = await this.loadCsvPreview(file, (p) => {
            this.uploadProgress = Math.max(5, Math.min(100, p));
            this.cdr.detectChanges();
          });

          clearInterval(progressInterval);
          this.uploadProgress = 100;
          this.workbook = null;
          this.sheetNames = ['CSV'];
          this.selectedSheet = 'CSV';
          this.csvStreamingFile = file;
          this.csvDelimiter = preview.delimiter;
          this.rawData = preview.rows;
          this.previewData = preview.rows.slice(0, this.MAX_PREVIEW_ROWS);
          this.headerRowIndex = null;
          this.headers = [];
          this.selectedColumns = [];
          this.showPreview = this.previewData.length > 0;

          if (this.previewData.length === 0) {
            this.showError('CSV file appears to be empty.');
            this.isUploading = false;
            return;
          }

          this.autoDetectHeader();
          this.isUploading = false;
          this.showSuccess(`CSV loaded. Previewing first ${this.previewData.length} rows.`);
          this.cdr.detectChanges();
        } catch (error) {
          clearInterval(progressInterval);
          this.isUploading = false;
          this.showError('Failed to read CSV file. Please ensure it is valid.');
          console.error('CSV parsing error:', error);
        }
      });
      return;
    }

    const reader = new FileReader();

    reader.onload = (e: any) => {
      clearInterval(progressInterval);
      this.uploadProgress = 100;
      this.cdr.detectChanges();

      this.ngZone.run(async () => {
        try {
          const XLSX = await this.loadXlsx();
          const result = e.target.result as any;

          // Text-based formats: CSV, TSV, TXT, XML
          if (ext === '.tsv' || ext === '.txt') {
            const text = result as string;
            const lines = text.split(/\r\n|\n/).filter(l => l.length > 0);
            // Detect delimiter for .txt (prefer tab for .tsv)
            const delim = ext === '.tsv' ? '\t' : this.detectDelimiter(lines[0] || '');
            const rows = lines.map(line => this.parseLine(line, delim));
            const ws = XLSX.utils.aoa_to_sheet(rows);
            this.workbook = { SheetNames: ['Sheet1'], Sheets: { Sheet1: ws } } as WorkBook;
          } else if (ext === '.xml') {
            const text = result as string;
            this.workbook = XLSX.read(text, { type: 'string', cellDates: true, dense: true });
          } else {
            // Binary formats (.xls, .xlsx, .xlsm, .xlsb, .ods)
            const data = new Uint8Array(result);
            if (ext === '.xlsx') {
              const preflight = await this.preflightXlsxComplexity(data);
              if (preflight.error) {
                this.showError(preflight.error);
                this.isUploading = false;
                return;
              }
              if (preflight.warning) {
                this.showSuccess(preflight.warning);
              }
            }
            // XLSX library does not execute macros - reading macro-enabled files is safe.
            try {
              this.workbook = XLSX.read(data, { type: 'array', cellDates: true, dense: true });
            } catch (binaryReadError: any) {
              const message = String(binaryReadError?.message || '');
              const canRepairZip =
                ext === '.xlsx' &&
                /bad compressed size|invalid zip|end of data|corrupt|crc/i.test(message);
              if (!canRepairZip) throw binaryReadError;

              if (this.DEBUG) console.warn('Primary XLSX parse failed. Attempting ZIP repair...', message);
              const repairedData = await this.repairZipContainer(data);
              this.workbook = XLSX.read(repairedData, { type: 'array', cellDates: true, dense: true });
              this.showSuccess('File repaired and loaded successfully');
            }
          }
          
          if (this.DEBUG) console.log('Workbook parsed, sheets:', this.workbook?.SheetNames);
          if (!this.workbook?.SheetNames?.length) {
            this.showError('Excel file contains no sheets');
            this.isUploading = false;
            return;
          }
          
          this.sheetNames = this.workbook.SheetNames;
          this.selectedSheet = this.sheetNames[0];
          
          setTimeout(async () => {
            this.isUploading = false;
            this.showSuccess('File loaded successfully');
            await this.previewSheet();
            this.cdr.detectChanges();
          }, 300);
        } catch (error: any) {
          const message = String(error?.message || '').toLowerCase();
          if (message.includes('password') || message.includes('encrypted') || message.includes('decrypt')) {
            this.showError('This file appears to be password-protected or encrypted. Please remove protection and try again.');
          } else {
            this.showError('Failed to read Excel file. Please ensure it is valid.');
          }
          console.error('Excel parsing error:', error);
          this.workbook = null;
          this.isUploading = false;
        }
      });
    };
    reader.onerror = () => {
      clearInterval(progressInterval);
      this.ngZone.run(() => {
        this.showError('Failed to load file. Please try again.');
        console.error('FileReader error:', reader.error);
        this.isUploading = false;
      });
    };
    // Read text formats as text, binary formats as array buffer
    if (ext === '.csv' || ext === '.tsv' || ext === '.txt' || ext === '.xml') {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  }

  /* ================= DRAG & DROP ================= */

  onDragOver(event: DragEvent) {
    event.preventDefault();
    this.isDragging = true;
  }

  onDragLeave(event: DragEvent) {
    event.preventDefault();
    const target = event.currentTarget as HTMLElement;
    const related = event.relatedTarget as HTMLElement | null;
    if (!related || !target.contains(related)) {
      this.isDragging = false;
    }
  }

  onDrop(event: DragEvent) {
    event.preventDefault();
    this.isDragging = false;

    if (!event.dataTransfer?.files.length) return;

    const file = event.dataTransfer.files[0];

    if (!this.isValidExcelFile(file)) {
      this.showError('Please drop a supported spreadsheet file (.xls, .xlsx, .xlsm, .xlsb, .csv, .tsv, .txt, .xml, .ods)');
      return;
    }

    this.onFileChange({ target: { files: [file] } });
  }

  /* ================= PREVIEW ================= */

  async previewSheet() {
    this.clearMessages();
    
    if (!this.workbook) {
      this.showError('File is still loading. Please wait and try again.');
      return;
    }

    if (!this.selectedSheet || !this.workbook.Sheets[this.selectedSheet]) {
      this.showError('Selected sheet not found');
      return;
    }

    try {
      const XLSX = await this.loadXlsx();
      if (this.DEBUG) console.log('Previewing sheet:', this.selectedSheet, 'Available sheets:', this.workbook?.SheetNames);
      const sheet = this.workbook.Sheets[this.selectedSheet];
      const sheetAssessment = this.assessSheetSize(sheet, XLSX);
      if (sheetAssessment.error) {
        this.showError(sheetAssessment.error);
        return;
      }

      // Use raw values so numeric mobile cells are not converted to scientific-notation strings.
      this.rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
      if (this.DEBUG) console.log('Raw data rows:', this.rawData.length);
      
      if (this.rawData.length === 0) {
        this.showError('Selected sheet is empty');
        return;
      }

      this.previewData = this.rawData.slice(0, this.MAX_PREVIEW_ROWS);
      this.headerRowIndex = null;
      this.headers = [];
      this.selectedColumns = [];
      this.showPreview = true;
      
      // Auto-detect and mark header row
      this.autoDetectHeader();
      
      const sizeNote = sheetAssessment.warning ? ` ${sheetAssessment.warning}` : '';
      this.showSuccess(`Loaded ${this.rawData.length} rows (showing first ${this.previewData.length}).${sizeNote}`);
    } catch (error) {
      this.showError('Failed to preview sheet');
      console.error('Preview error:', error);
    } finally {
      this.cdr.detectChanges();
    }
  }

  private autoDetectHeader() {
    for (let i = 0; i < Math.min(5, this.rawData.length); i++) {
      const row = this.rawData[i];
      if (!row || row.length === 0) continue;
      
      const hasText = row.some(cell => {
        const str = String(cell || '').trim();
        return str && /[a-zA-Z]/.test(str);
      });
      
      if (hasText) {
        this.setHeaderRow(i);
        return;
      }
    }
  }

  /* ================= HEADER ROW ================= */

  setHeaderRow(index: number) {
    this.clearMessages();
    
    if (index < 0 || index >= this.rawData.length) {
      this.showError('Invalid row index');
      return;
    }
    
    const headerRow = this.rawData[index];
    if (!headerRow || headerRow.length === 0 || headerRow.every(cell => !cell)) {
      this.showError('Selected row is empty. Please choose a valid header row.');
      return;
    }
    
    this.headerRowIndex = index;
    this.headers = headerRow.map(h => String(h || '').trim());
    this.selectedColumns = [];

    for (let col = 0; col < this.headers.length; col++) {
      let validCount = 0;
      const maxRows = Math.min(index + 10, this.rawData.length);

      for (let row = index + 1; row < maxRows; row++) {
        const cell = this.rawData[row]?.[col];
        if (this.cleanMobile(cell)) validCount++;
      }

      if (validCount >= this.MIN_VALID_MOBILES) this.selectedColumns.push(col);
    }

    // Auto-detect name column
    this.autoDetectNameColumn();

    if (this.selectedColumns.length > 0) {
      this.showSuccess(`Auto-detected ${this.selectedColumns.length} mobile column(s)`);
    } else {
      this.showSuccess('Header row set. Please manually select mobile columns.');
    }
  }

  toggleColumn(index: number) {
    const pos = this.selectedColumns.indexOf(index);
    if (pos !== -1) {
      this.selectedColumns.splice(pos, 1);
    } else {
      this.selectedColumns.push(index);
    }
  }

  private autoDetectNameColumn() {
    this.selectedNameColumn = null;
    for (let col = 0; col < this.headers.length; col++) {
      const header = this.headers[col].toLowerCase();
      if (header.includes('name') || header.includes('customer') || header.includes('contact')) {
        this.selectedNameColumn = col;
        return;
      }
    }
  }

  /* ================= CLEAN & DOWNLOAD (CHUNKED) ================= */

  async cleanAndDownload() {
    this.clearMessages();
    
    if (this.headerRowIndex === null) {
      this.showError('Please select a header row first');
      return;
    }

    if (this.selectedColumns.length === 0) {
      this.showError('Please select at least one mobile column');
      return;
    }

    if (this.exportMode === 'mobile-name' && this.selectedNameColumn === null) {
      this.showError('Please select a name column for mobile-name export');
      return;
    }

    if (this.csvStreamingFile) {
      await this.cleanAndDownloadCsvStreaming();
      return;
    }

    if (this.isProcessing) return;

    try {
      this.isProcessing = true;
      this.progress = 0;
      this.resetStats();
      if (this.DEBUG) console.log(`Starting data processing. Header row index: ${this.headerRowIndex}, Selected columns:`, this.selectedColumns);
      if (this.DEBUG) console.log('Raw data length:', this.rawData.length, 'Headers:', this.headers.length);

      const seenNumbers = new Set<string>();
      const cleaned: any[][] = [];
      const uniqueNumbers: string[] = [];
      const mobileNamePairs: any[][] = []; // Store name-mobile pairs during processing
      const keepAllRows = this.exportMode === 'keep-all';
      const needsFullRows = this.exportMode === 'full' || this.exportMode === 'keep-all';
      const needsMobileNamePairs = this.exportMode === 'mobile-name';
      const exportHeaders = this.headers.map(header => this.toExportHeader(header));
      
      cleaned.push(exportHeaders);

      const totalRows = this.rawData.length - (this.headerRowIndex + 1);
      if (totalRows <= 0) {
        this.showError('No data rows found after header row.');
        this.isProcessing = false;
        return;
      }

      if (this.DEBUG) console.log(`Total rows to process: ${totalRows}`);
      this.stats.total = totalRows;
      const fastMode = totalRows >= this.FAST_MODE_ROWS;
      const chunkSize =
        totalRows >= 250000 ? 6000 :
        totalRows >= 120000 ? 4000 :
        totalRows >= 60000 ? 2500 : 1200;
      let processed = 0;
      let sliceStartMs = Date.now();

      for (let start = this.headerRowIndex + 1; start < this.rawData.length; start += chunkSize) {
        const end = Math.min(start + chunkSize, this.rawData.length);

          for (let i = start; i < end; i++) {
            try {
              const row = needsFullRows ? [...(this.rawData[i] || [])] : (this.rawData[i] || []);
              if (needsFullRows) {
                while (row.length < this.headers.length) row.push('');
              }

              // Capture name BEFORE modifying the row
              const name = this.selectedNameColumn !== null ? (row[this.selectedNameColumn] || '') : '';

              const validMobiles: string[] = [];
              const perColumn: (string | null)[] = [];

              for (const col of this.selectedColumns) {
                const detail = this.cleanMobileDetailed(row[col]);
                if (needsFullRows) {
                  perColumn.push(detail.cleanedNumbers.length > 0 ? detail.cleanedNumbers.join(' / ') : null);
                }
                const headerName = this.headers[col] || `Column_${col + 1}`;
                if (detail.cleanedNumbers.length > 0) {
                  validMobiles.push(...detail.cleanedNumbers);
                  if (!fastMode) {
                    for (const cleanedNumber of detail.cleanedNumbers) {
                      this.addLimitedStatRow('valid', {
                        row: i + 1,
                        column: headerName,
                        original: String(row[col] ?? ''),
                        cleaned: cleanedNumber
                      });
                    }
                  }
                } else if (detail.reason === 'invalidPattern') {
                  this.stats.invalidPattern++;
                  if (!fastMode) {
                    this.addLimitedStatRow('invalidPattern', {
                      row: i + 1,
                      column: headerName,
                      value: String(row[col] ?? '')
                    });
                  }
                } else if (detail.reason === 'invalidLength') {
                  this.stats.invalidLength++;
                  if (!fastMode) {
                    this.addLimitedStatRow('invalidLength', {
                      row: i + 1,
                      column: headerName,
                      value: String(row[col] ?? '')
                    });
                  }
                }
              }

              const rowMobiles = Array.from(new Set(validMobiles));

              if (keepAllRows) {
                this.selectedColumns.forEach((col, idx) => {
                  if (perColumn[idx]) row[col] = perColumn[idx];
                });
                if (rowMobiles.length > 0) this.stats.valid += rowMobiles.length;
                cleaned.push(row);
                processed++;
                continue;
              }

              if (rowMobiles.length === 0) {
                processed++;
                continue;
              }

              const unseenMobiles: string[] = [];
              const duplicateMobiles: string[] = [];
              for (const mobile of rowMobiles) {
                const normalized = mobile.replace('+91', '');
                if (seenNumbers.has(normalized)) duplicateMobiles.push(mobile);
                else unseenMobiles.push(mobile);
              }

              if (unseenMobiles.length === 0) {
                this.stats.duplicates++;
                if (!fastMode) {
                  for (const duplicate of duplicateMobiles) {
                    this.addLimitedStatRow('duplicates', { row: i + 1, mobile: duplicate });
                  }
                }
                processed++;
                continue;
              }

              for (const number of unseenMobiles) {
                const normalized = number.replace('+91', '');
                seenNumbers.add(normalized);
                uniqueNumbers.push(number);
              }

              if (needsFullRows) {
                const primaryNumber = unseenMobiles[0];
                const fallback = primaryNumber;
                this.selectedColumns.forEach((col, idx) => {
                  row[col] = perColumn[idx] ?? fallback;
                });
                cleaned.push(row);
              }

              if (needsMobileNamePairs) {
                for (const number of unseenMobiles) {
                  mobileNamePairs.push([name, number]);
                }
              }

              this.stats.valid += unseenMobiles.length;
              processed++;

              // Time-slice processing so the browser does not show "wait/exit" prompts.
              if (Date.now() - sliceStartMs >= 20) {
                await this.yieldToBrowser();
                sliceStartMs = Date.now();
              }
            } catch (rowError) {
              console.error(`Error processing row ${i}:`, rowError);
              this.stats.invalidLength++;
              processed++;
              continue;
            }
          }

        this.progress = Math.round((processed / totalRows) * 100);
        this.cdr.detectChanges();
        if (this.DEBUG && this.progress % 10 === 0) {
          console.log(`Processing progress: ${this.progress}%, Rows: ${processed}/${totalRows}, Valid: ${this.stats.valid}`);
        }
        await this.yieldToBrowser();
      }

      if (!keepAllRows && cleaned.length <= 1 && uniqueNumbers.length === 0) {
        if (this.exportMode === 'mobile-name' && mobileNamePairs.length > 0) {
          // mobile-name export does not populate cleaned rows
        } else if (this.exportMode === 'unique' && uniqueNumbers.length > 0) {
          // unique export does not populate cleaned rows
        } else {
        console.warn('No valid mobile numbers found');
        this.showError('No valid mobile numbers found in the selected columns.');
        this.isProcessing = false;
        this.progress = 0;
        return;
        }
      }

      if (this.DEBUG) {
        console.log(`Processing complete. Valid rows: ${this.stats.valid}, Duplicates: ${this.stats.duplicates}. Generating Excel file...`);
      }

      let sheetAdded = false;
      let fallbackData: any[][] = [];
      let exportRowCount = 0;

      if (this.exportMode === 'full' || this.exportMode === 'keep-all') {
        if (cleaned.length > 1) {
          fallbackData = cleaned;
          exportRowCount = Math.max(0, cleaned.length - 1);
          sheetAdded = true;
        }
      } else if (this.exportMode === 'unique') {
        const uniqueSheet = [[this.toExportHeader('Mobile Number')], ...uniqueNumbers.map(n => [n])];
        fallbackData = uniqueSheet;
        exportRowCount = uniqueNumbers.length;
        sheetAdded = true;
      } else if (this.exportMode === 'mobile-name') {
        if (mobileNamePairs.length > 0) {
          const mobileNameData: any[][] = [[this.toExportHeader('Name'), this.toExportHeader('Mobile Number')], ...mobileNamePairs];
          fallbackData = mobileNameData;
          exportRowCount = mobileNamePairs.length;
          sheetAdded = true;
        }
      }

      if (!sheetAdded) {
        console.error('No sheet was added to workbook');
        this.showError('No data to export.');
        this.isProcessing = false;
        this.progress = 0;
        return;
      }

      if (this.DEBUG) console.log('Writing Excel file...');

      try {
        const preferCsvForLargeExport =
          exportRowCount >= 60000 ||
          (this.exportMode === 'keep-all' && exportRowCount >= 40000);
        if (preferCsvForLargeExport) {
          const csvName = `(${exportRowCount})-${this.fileName}.csv`;
          const csvBlob = this.buildCsvBlob(fallbackData);
          this.downloadBlob(csvBlob, csvName);
          await this.yieldToBrowser();
          this.progress = 100;
          this.showStats = true;
          this.showSuccess(`Large export detected. Downloaded CSV for better stability. ${exportRowCount} rows exported.`);
          return;
        }

        const XLSX = await this.loadXlsx();
        const wb = XLSX.utils.book_new();
        if (this.exportMode === 'full' || this.exportMode === 'keep-all') {
          const ws = XLSX.utils.aoa_to_sheet(fallbackData.map(row => row.map(cell => this.sanitizeForExcelCell(cell))));
          const sheetName = this.exportMode === 'keep-all' ? 'Cleaned Keep All Rows' : 'Cleaned';
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
        } else if (this.exportMode === 'unique') {
          const ws = XLSX.utils.aoa_to_sheet(fallbackData.map(row => row.map(cell => this.sanitizeForExcelCell(cell))));
          XLSX.utils.book_append_sheet(wb, ws, 'Unique Numbers');
        } else {
          const ws = XLSX.utils.aoa_to_sheet(fallbackData.map(row => row.map(cell => this.sanitizeForExcelCell(cell))));
          XLSX.utils.book_append_sheet(wb, ws, 'Mobile & Name');
        }

        const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const fileName = `(${exportRowCount})-${this.fileName}.xlsx`;
        
        if (this.DEBUG) console.log(`Generated Excel file: ${fileName}`, { size: blob.size, rows: cleaned.length });
        
        // Trigger browser-native download outside Angular zone.
        this.ngZone.runOutsideAngular(() => {
          try {
            this.downloadBlob(blob, fileName);
            if (this.DEBUG) console.log('File saved successfully');
          } catch (e) {
            console.error('download error:', e);
            throw e;
          }
        });
        
        // Small delay to ensure download starts before UI updates
        await this.yieldToBrowser();
        
        this.progress = 100;
        this.showStats = true;
        this.showSuccess(
          fastMode
            ? `Processed and downloaded successfully. ${exportRowCount} rows exported. Fast mode enabled for speed (detailed category downloads may be limited).`
            : `Processed and downloaded successfully. ${exportRowCount} rows exported.`
        );
      } catch (downloadError) {
        console.error('XLSX download error:', downloadError);
        try {
          const csvBlob = this.buildCsvBlob(fallbackData);
          const csvName = `(${exportRowCount})-${this.fileName}.csv`;
          this.downloadBlob(csvBlob, csvName);
          this.progress = 100;
          this.showStats = true;
          this.showSuccess(
            fastMode
              ? `XLSX export failed, downloaded CSV instead. ${exportRowCount} rows exported. Fast mode enabled for speed (detailed category downloads may be limited).`
              : `XLSX export failed, downloaded CSV instead. ${exportRowCount} rows exported.`
          );
        } catch (fallbackError) {
          console.error('CSV fallback error:', fallbackError);
          this.showError('Failed to download file. Try a smaller selection or export mode.');
          this.progress = 0;
        }
      } finally {
        // Small delay to ensure download initiates before resetting UI state
        setTimeout(() => {
          this.isProcessing = false;
          this.cdr.detectChanges();
          if (this.DEBUG) console.log('Processing complete, UI reset');
        }, 500);
      }
    } catch (error) {
      this.isProcessing = false;
      this.progress = 0;
      this.showError('Failed to process data. Please try again.');
      console.error('Processing error:', error);
      this.cdr.detectChanges();
    }
  }

  private yieldToBrowser(): Promise<void> {
    return new Promise(resolve => {
      if (typeof requestAnimationFrame === 'function') {
        requestAnimationFrame(() => resolve());
      } else {
        setTimeout(resolve, 0);
      }
    });
  }

  /* ================= MOBILE CLEANER ================= */

  cleanMobile(value: any, trackStats = false): string | null {
    const detail = this.cleanMobileDetailed(value);
    if (trackStats) {
      if (detail.reason === 'invalidPattern') this.stats.invalidPattern++;
      if (detail.reason === 'invalidLength') this.stats.invalidLength++;
    }
    return detail.cleaned;
  }

  private cleanMobileDetailed(value: any): { cleaned: string | null; cleanedNumbers: string[]; reason: 'valid' | 'invalidPattern' | 'invalidLength' | 'empty' } {
    if (value === null || value === undefined || value === '') {
      return { cleaned: null, cleanedNumbers: [], reason: 'empty' };
    }

    const cleanedNumbers = this.extractValidMobiles(value, 2);
    if (cleanedNumbers.length > 0) {
      return { cleaned: cleanedNumbers[0], cleanedNumbers, reason: 'valid' };
    }

    const normalizedValue = this.normalizeMobileSource(value);
    const str = normalizedValue.replace(/[^0-9]/g, '');
    if (!str) {
      return { cleaned: null, cleanedNumbers: [], reason: 'invalidLength' };
    }

    const matches = str.match(/\d{10,12}/g);
    if (!matches) {
      return { cleaned: null, cleanedNumbers: [], reason: 'invalidLength' };
    }

    const hasTenDigitLike = matches.some(num => {
      let n = num;
      if (n.length === 12 && n.startsWith('91')) n = n.slice(2);
      if (n.length === 11 && n.startsWith('0')) n = n.slice(1);
      return n.length === 10;
    });

    return { cleaned: null, cleanedNumbers: [], reason: hasTenDigitLike ? 'invalidPattern' : 'invalidLength' };
  }

  private extractValidMobiles(value: any, maxNumbers = 2): string[] {
    const normalizedValue = this.normalizeMobileSource(value);
    const str = normalizedValue.replace(/[^0-9]/g, '');
    if (!str) return [];

    const matches = str.match(/\d{10,12}/g);
    if (!matches) return [];

    const result: string[] = [];
    for (let num of matches) {
      if (num.length === 12 && num.startsWith('91')) num = num.slice(2);
      if (num.length === 11 && num.startsWith('0')) num = num.slice(1);
      if (num.length !== 10) continue;
      if (!/^[6-9]\d{9}$/.test(num)) continue;
      if (this.isInvalidPattern(num)) continue;

      const cleaned = `+91${num}`;
      if (!result.includes(cleaned)) {
        result.push(cleaned);
        if (result.length >= maxNumbers) break;
      }
    }
    return result;
  }

  private normalizeMobileSource(value: any): string {
    if (typeof value === 'number' && Number.isFinite(value)) {
      return this.numberToPlainString(value);
    }

    const text = String(value).trim();
    if (/^[+-]?\d+(\.\d+)?e[+-]?\d+$/i.test(text)) {
      const parsed = Number(text);
      if (Number.isFinite(parsed)) {
        return this.numberToPlainString(parsed);
      }
    }

    return text;
  }

  private numberToPlainString(value: number): string {
    if (Number.isInteger(value)) return String(value);
    return value.toLocaleString('en-US', { useGrouping: false, maximumFractionDigits: 20 });
  }

  private isInvalidPattern(num: string): boolean {
    return this.INVALID_PATTERNS.some(pattern => pattern.test(num));
  }

  private detectDelimiter(sample: string): string {
    const candidates = [',', '\t', ';', '|'];
    let best = ',';
    let bestCount = -1;
    for (const c of candidates) {
      const count = (sample.split(c).length - 1);
      if (count > bestCount) {
        bestCount = count;
        best = c;
      }
    }
    return best;
  }

  private parseLine(line: string, delim: string): string[] {
    // Generic parser that handles quoted fields with the given delimiter
    const out: string[] = [];
    let cur = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') {
          cur += '"';
          i++; // skip escaped quote
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === delim && !inQuotes) {
        out.push(cur);
        cur = '';
      } else {
        cur += ch;
      }
    }
    out.push(cur);
    return out.map(s => s.replace(/^"|"$/g, '').trim());
  }

  private sanitizeForExcelCell(value: any): string | number | boolean | Date | null {
    if (value === undefined || value === null) return null;
    if (value instanceof Date) return value;
    if (typeof value === 'number' || typeof value === 'boolean') return value;
    const text = String(value);
    // Excel cell text limit is 32767 characters.
    return text.length > 32767 ? text.slice(0, 32767) : text;
  }

  private buildCsvBlob(data: any[][]): Blob {
    const rows = data.map(row =>
      row.map(cell => {
        const normalized = this.sanitizeForExcelCell(cell);
        const text = normalized === null ? '' : String(normalized);
        if (/[",\n\r]/.test(text)) {
          return `"${text.replace(/"/g, '""')}"`;
        }
        return text;
      }).join(',')
    );
    const csv = '\uFEFF' + rows.join('\r\n');
    return new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  }

  private csvEscape(value: any): string {
    const normalized = this.sanitizeForExcelCell(value);
    const text = normalized === null ? '' : String(normalized);
    if (/[",\n\r]/.test(text)) {
      return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
  }

  private async loadCsvPreview(file: File, onProgress?: (progress: number) => void): Promise<{ rows: string[][]; delimiter: string }> {
    const rows: string[][] = [];
    let delimiter = ',';
    let delimiterDetected = false;
    let rowCount = 0;

    await this.streamCsvLines(
      file,
      (line) => {
        if (!delimiterDetected) {
          delimiter = this.detectDelimiter(line);
          delimiterDetected = true;
        }
        rows.push(this.parseLine(line, delimiter));
        rowCount++;
        return rowCount < this.MAX_PREVIEW_ROWS;
      },
      onProgress
    );

    return { rows, delimiter };
  }

  private async streamCsvLines(
    file: File,
    onLine: (line: string) => boolean | void | Promise<boolean | void>,
    onProgress?: (progress: number) => void
  ): Promise<void> {
    const decoder = new TextDecoder('utf-8');
    let offset = 0;
    let leftover = '';
    let shouldContinue = true;

    while (offset < file.size && shouldContinue) {
      const next = Math.min(offset + this.CSV_STREAM_CHUNK_SIZE, file.size);
      const buffer = await file.slice(offset, next).arrayBuffer();
      offset = next;

      const text = decoder.decode(buffer, { stream: offset < file.size });
      const merged = leftover + text;
      const lines = merged.split(/\r\n|\n/);
      leftover = lines.pop() ?? '';

      for (const line of lines) {
        const result = await onLine(line);
        if (result === false) {
          shouldContinue = false;
          break;
        }
      }

      if (onProgress) {
        onProgress(Math.round((offset / file.size) * 100));
      }
      await this.yieldToBrowser();
    }

    if (shouldContinue && leftover.length > 0) {
      await onLine(leftover);
      if (onProgress) onProgress(100);
    }
  }

  private addLimitedStatRow(type: 'valid' | 'duplicates' | 'invalidPattern' | 'invalidLength', row: any) {
    const bucket = this.statDownloads[type] as any[];
    if (bucket.length < this.MAX_STAT_DOWNLOAD_ROWS) {
      bucket.push(row);
    }
  }

  private async cleanAndDownloadCsvStreaming() {
    if (!this.csvStreamingFile) return;
    if (this.isProcessing) return;

    this.isProcessing = true;
    this.progress = 0;
    this.resetStats();

    try {
      const file = this.csvStreamingFile;
      const keepAllRows = this.exportMode === 'keep-all';
      const needsFullRows = this.exportMode === 'full' || this.exportMode === 'keep-all';
      const needsMobileNamePairs = this.exportMode === 'mobile-name';
      const seenNumbers = new Set<string>();
      let delimiter = this.csvDelimiter || ',';
      let delimiterDetected = false;
      let rowIndex = -1;
      let exportRowCount = 0;
      let headerWritten = false;
      let dataRowsProcessed = 0;

      const csvParts: string[] = ['\uFEFF'];
      let csvBuffer = '';
      let bufferedLines = 0;
      const flushCsvBuffer = () => {
        if (csvBuffer.length === 0) return;
        csvParts.push(csvBuffer);
        csvBuffer = '';
        bufferedLines = 0;
      };
      const appendCsvLine = (line: string) => {
        csvBuffer += line;
        bufferedLines++;
        if (bufferedLines >= 2000 || csvBuffer.length >= 2 * 1024 * 1024) {
          flushCsvBuffer();
        }
      };

      await this.streamCsvLines(
        file,
        async (line) => {
          rowIndex++;
          if (!delimiterDetected) {
            delimiter = this.detectDelimiter(line);
            delimiterDetected = true;
          }

          const row = this.parseLine(line, delimiter);
          if (rowIndex < (this.headerRowIndex ?? 0)) return true;

          if (rowIndex === this.headerRowIndex) {
            if (this.exportMode === 'full' || this.exportMode === 'keep-all') {
              const exportHeaders = this.headers.map(header => this.toExportHeader(header));
              appendCsvLine(exportHeaders.map(v => this.csvEscape(v)).join(',') + '\r\n');
            } else if (this.exportMode === 'unique') {
              appendCsvLine(`${this.csvEscape(this.toExportHeader('Mobile Number'))}\r\n`);
            } else {
              appendCsvLine(`${this.csvEscape(this.toExportHeader('Name'))},${this.csvEscape(this.toExportHeader('Mobile Number'))}\r\n`);
            }
            headerWritten = true;
            return true;
          }

          if (!headerWritten) return true;
          dataRowsProcessed++;
          this.stats.total = dataRowsProcessed;

          if (row.length < this.headers.length) {
            while (row.length < this.headers.length) row.push('');
          }

          const name = this.selectedNameColumn !== null ? (row[this.selectedNameColumn] || '') : '';
          const validMobiles: string[] = [];
          const perColumn: (string | null)[] = [];

          for (const col of this.selectedColumns) {
            const detail = this.cleanMobileDetailed(row[col]);
            if (needsFullRows) {
              perColumn.push(detail.cleanedNumbers.length > 0 ? detail.cleanedNumbers.join(' / ') : null);
            }
            const headerName = this.headers[col] || `Column_${col + 1}`;

            if (detail.cleanedNumbers.length > 0) {
              validMobiles.push(...detail.cleanedNumbers);
              for (const cleanedNumber of detail.cleanedNumbers) {
                this.addLimitedStatRow('valid', {
                  row: rowIndex + 1,
                  column: headerName,
                  original: String(row[col] ?? ''),
                  cleaned: cleanedNumber
                });
              }
            } else if (detail.reason === 'invalidPattern') {
              this.stats.invalidPattern++;
              this.addLimitedStatRow('invalidPattern', {
                row: rowIndex + 1,
                column: headerName,
                value: String(row[col] ?? '')
              });
            } else if (detail.reason === 'invalidLength') {
              this.stats.invalidLength++;
              this.addLimitedStatRow('invalidLength', {
                row: rowIndex + 1,
                column: headerName,
                value: String(row[col] ?? '')
              });
            }
          }

          const rowMobiles = Array.from(new Set(validMobiles));
          if (keepAllRows) {
            this.selectedColumns.forEach((col, idx) => {
              if (perColumn[idx]) row[col] = perColumn[idx] as string;
            });
            if (rowMobiles.length > 0) this.stats.valid += rowMobiles.length;
            appendCsvLine(row.map(v => this.csvEscape(v)).join(',') + '\r\n');
            exportRowCount++;
            return true;
          }

          if (rowMobiles.length === 0) return true;

          const unseenMobiles: string[] = [];
          const duplicateMobiles: string[] = [];
          for (const mobile of rowMobiles) {
            const normalized = mobile.replace('+91', '');
            if (seenNumbers.has(normalized)) duplicateMobiles.push(mobile);
            else unseenMobiles.push(mobile);
          }

          if (unseenMobiles.length === 0) {
            this.stats.duplicates++;
            for (const duplicate of duplicateMobiles) {
              this.addLimitedStatRow('duplicates', { row: rowIndex + 1, mobile: duplicate });
            }
            return true;
          }

          for (const number of unseenMobiles) {
            seenNumbers.add(number.replace('+91', ''));
          }
          this.stats.valid += unseenMobiles.length;

          if (this.exportMode === 'unique') {
            for (const number of unseenMobiles) {
              appendCsvLine(this.csvEscape(number) + '\r\n');
              exportRowCount++;
            }
            return true;
          }

          if (needsMobileNamePairs) {
            for (const number of unseenMobiles) {
              appendCsvLine(`${this.csvEscape(name)},${this.csvEscape(number)}\r\n`);
              exportRowCount++;
            }
            return true;
          }

          // full export
          const primary = unseenMobiles[0];
          this.selectedColumns.forEach((col, idx) => {
            row[col] = perColumn[idx] ?? primary;
          });
          appendCsvLine(row.map(v => this.csvEscape(v)).join(',') + '\r\n');
          exportRowCount++;
          return true;
        },
        (progress) => {
          this.progress = progress;
          this.cdr.detectChanges();
        }
      );

      if (exportRowCount <= 0) {
        this.showError('No valid mobile numbers found in the selected columns.');
        this.progress = 0;
        this.isProcessing = false;
        return;
      }

      flushCsvBuffer();
      const blob = new Blob(csvParts, { type: 'text/csv;charset=utf-8;' });
      const fileName = `(${exportRowCount})-${this.fileName}.csv`;
      this.downloadBlob(blob, fileName);
      this.progress = 100;
      this.showStats = true;
      this.showSuccess(`Processed and downloaded successfully. ${exportRowCount} rows exported.`);
    } catch (error) {
      console.error('CSV streaming processing error:', error);
      this.showError('Failed to process large CSV. Please try again with a split file.');
      this.progress = 0;
    } finally {
      this.isProcessing = false;
      this.cdr.detectChanges();
    }
  }

  private toExportHeader(header: string): string {
    return String(header || '')
      .trim()
      .replace(/\s+/g, '_');
  }

  private resetStats() {
    this.stats = {
      total: 0,
      valid: 0,
      duplicates: 0,
      invalidPattern: 0,
      invalidLength: 0
    };
    this.statDownloads = {
      valid: [],
      duplicates: [],
      invalidPattern: [],
      invalidLength: []
    };
    this.showStats = false;
  }

  async downloadStatReport(type: 'valid' | 'duplicates' | 'invalidPattern' | 'invalidLength') {
    this.clearMessages();
    const rows = this.statDownloads[type];
    if (!rows.length) {
      this.showError('No rows available for this category.');
      return;
    }

    try {
      const XLSX = await this.loadXlsx();
      const wb = XLSX.utils.book_new();
      let data: any[][] = [];
      let sheetName = '';

      if (type === 'valid') {
        data = [[this.toExportHeader('Row'), this.toExportHeader('Column'), this.toExportHeader('Original_Value'), this.toExportHeader('Cleaned_Number')]];
        data.push(...rows.map((r: any) => [r.row, r.column, r.original, r.cleaned]));
        sheetName = 'Valid Numbers';
      } else if (type === 'duplicates') {
        data = [[this.toExportHeader('Row'), this.toExportHeader('Duplicate_Mobile')]];
        data.push(...rows.map((r: any) => [r.row, r.mobile]));
        sheetName = 'Duplicates Removed';
      } else if (type === 'invalidPattern') {
        data = [[this.toExportHeader('Row'), this.toExportHeader('Column'), this.toExportHeader('Invalid_Value')]];
        data.push(...rows.map((r: any) => [r.row, r.column, r.value]));
        sheetName = 'Invalid Patterns';
      } else {
        data = [[this.toExportHeader('Row'), this.toExportHeader('Column'), this.toExportHeader('Invalid_Value')]];
        data.push(...rows.map((r: any) => [r.row, r.column, r.value]));
        sheetName = 'Invalid Length Format';
      }

      const ws = XLSX.utils.aoa_to_sheet(data);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const fileName = `${this.toExportHeader(type)}_${this.fileName}.xlsx`;
      this.downloadBlob(blob, fileName);
      this.showSuccess(`Downloaded ${rows.length} row(s) for ${type}.`);
    } catch (error) {
      console.error('Stat download error:', error);
      this.showError('Failed to download category file.');
    }
  }

  /* ================= UTILITY METHODS ================= */

  private assessSheetSize(sheet: WorkSheet, xlsx: typeof import('xlsx')): { error?: string; warning?: string } {
    const ref = sheet['!ref'];
    if (!ref) {
      return { error: 'Selected sheet is empty' };
    }

    const range = xlsx.utils.decode_range(ref);
    const rowCount = range.e.r - range.s.r + 1;
    const colCount = range.e.c - range.s.c + 1;
    const estimatedCells = rowCount * colCount;

    if (rowCount > this.MAX_SHEET_ROWS) {
      return { error: `Sheet has ${rowCount} rows. Maximum supported rows are ${this.MAX_SHEET_ROWS}.` };
    }

    if (colCount > this.MAX_SHEET_COLUMNS) {
      return { error: `Sheet has ${colCount} columns. Maximum supported columns are ${this.MAX_SHEET_COLUMNS}.` };
    }

    if (estimatedCells > this.MAX_SHEET_CELLS) {
      return { error: `Sheet is too large (${estimatedCells} cells). Maximum supported cells are ${this.MAX_SHEET_CELLS}.` };
    }

    if (rowCount >= this.LARGE_SHEET_WARNING_ROWS) {
      return { warning: 'Large sheet detected. Processing may take longer.' };
    }

    return {};
  }

  private async loadXlsx(): Promise<typeof import('xlsx')> {
    if (!this.xlsxModule) {
      this.xlsxModule = await import('xlsx');
    }
    return this.xlsxModule;
  }

  private async loadJsZip(): Promise<typeof JSZip> {
    if (!this.jsZipModule) {
      const module = await import('jszip');
      this.jsZipModule = module.default;
    }
    return this.jsZipModule;
  }

  private async preflightXlsxComplexity(data: Uint8Array): Promise<{ error?: string; warning?: string }> {
    try {
      const JSZipLib = await this.loadJsZip();
      const zip = await JSZipLib.loadAsync(data);
      const worksheetEntries = Object.values((zip as any).files || {}).filter((entry: any) => {
        const name = String(entry?.name || '');
        return /^xl\/worksheets\/sheet\d+\.xml$/i.test(name);
      }) as any[];
      if (!worksheetEntries.length) return {};

      let maxSheetXmlSize = 0;
      for (const entry of worksheetEntries) {
        const size = Number(entry?._data?.uncompressedSize || 0);
        if (Number.isFinite(size) && size > maxSheetXmlSize) {
          maxSheetXmlSize = size;
        }
      }
      if (maxSheetXmlSize <= 0) return {};

      const TOO_LARGE_XML = 250 * 1024 * 1024; // 250MB uncompressed worksheet XML
      const LARGE_XML_WARNING = 120 * 1024 * 1024; // 120MB warning

      if (maxSheetXmlSize > TOO_LARGE_XML) {
        return {
          error: 'This XLSX is too complex for browser memory. Please save/export it as CSV and upload CSV, or split the workbook into smaller files.'
        };
      }
      if (maxSheetXmlSize > LARGE_XML_WARNING) {
        return {
          warning: 'Large XLSX detected. Processing may be slow; CSV format is recommended for best performance.'
        };
      }
      return {};
    } catch {
      // If preflight fails, continue with normal parse path.
      return {};
    }
  }

  private async repairZipContainer(data: Uint8Array): Promise<Uint8Array> {
    const JSZipLib = await this.loadJsZip();
    const zip = await JSZipLib.loadAsync(data);
    const repaired = await zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    return repaired;
  }

  private downloadBlob(blob: Blob, fileName: string) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    anchor.style.display = 'none';
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  private resetState() {
    this.workbook = null;
    this.showPreview = false;
    this.isProcessing = false;
    this.progress = 0;
    this.headerRowIndex = null;
    this.headers = [];
    this.selectedColumns = [];
    this.selectedNameColumn = null;
    this.csvStreamingFile = null;
    this.csvDelimiter = ',';
    this.rawData = [];
    this.previewData = [];
    this.resetStats();
  }

  private clearMessages() {
    this.errorMessage = '';
    this.successMessage = '';
  }

  private showError(message: string) {
    this.errorMessage = message;
    this.successMessage = '';
    if (this.errorTimer) clearTimeout(this.errorTimer);
    this.errorTimer = setTimeout(() => this.errorMessage = '', 5000);
  }

  private showSuccess(message: string) {
    this.successMessage = message;
    this.errorMessage = '';
    if (this.successTimer) clearTimeout(this.successTimer);
    this.successTimer = setTimeout(() => this.successMessage = '', 5000);
  }

  private isValidExcelFile(file: File): boolean {
    // Accept common spreadsheet formats including CSV, TSV, TXT, XML and OpenDocument spreadsheets
    return /\.(xls|xlsx|xlsm|xlsb|csv|tsv|txt|xml|ods)$/i.test(file.name);
  }
}




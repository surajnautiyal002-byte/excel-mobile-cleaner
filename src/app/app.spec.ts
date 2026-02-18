import { TestBed } from '@angular/core/testing';
import { AppComponent } from './app';
import * as XLSX from 'xlsx';

describe('AppComponent', () => {
  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [AppComponent],
    }).compileComponents();
  });

  it('should create the app', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    expect(app).toBeTruthy();
  });

  it('should initialize with empty file name', async () => {
    const fixture = TestBed.createComponent(AppComponent);
    await fixture.whenStable();
    const app = fixture.componentInstance;
    expect(app.fileName).toBe('');
  });

  it('should clean mobile from scientific notation string', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    expect(app.cleanMobile('9.19818202888E+11')).toBe('+919818202888');
  });

  it('should normalize mixed mobile formats to same dedupe-ready value', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance;
    const m1 = app.cleanMobile('919818202888');
    const m2 = app.cleanMobile('+91 98182-02888');
    expect(m1).toBe('+919818202888');
    expect(m2).toBe('+919818202888');
    expect(m1).toBe(m2);
  });

  it('should show empty-sheet error when selected sheet has no cells', async () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance as any;

    app.workbook = {
      SheetNames: ['Sheet1'],
      Sheets: { Sheet1: {} as XLSX.WorkSheet }
    } as XLSX.WorkBook;
    app.selectedSheet = 'Sheet1';

    await app.previewSheet();

    expect(app.errorMessage).toContain('empty');
  });

  it('should block oversized sheet before parsing all data', async () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance as any;

    app.workbook = {
      SheetNames: ['Big'],
      Sheets: { Big: { '!ref': 'A1:Z200001' } as XLSX.WorkSheet }
    } as XLSX.WorkBook;
    app.selectedSheet = 'Big';

    await app.previewSheet();

    expect(app.errorMessage).toContain('Maximum supported rows');
  });

  it('should parse xlsx round-trip and preserve sheet data', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['Name', 'Mobile Number'], ['A', 919818202888]]);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    const out = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const parsed = XLSX.read(out, { type: 'buffer' });
    const rows = XLSX.utils.sheet_to_json(parsed.Sheets['Sheet1'], { header: 1, raw: true }) as any[][];

    expect(parsed.SheetNames).toEqual(['Sheet1']);
    expect(rows.length).toBe(2);
    expect(rows[1][1]).toBe(919818202888);
  });

  it('should parse xls round-trip and preserve sheet data', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['Name', 'Mobile Number'], ['A', 919818202888]]);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    const out = XLSX.write(wb, { type: 'buffer', bookType: 'xls' });
    const parsed = XLSX.read(out, { type: 'buffer' });
    const rows = XLSX.utils.sheet_to_json(parsed.Sheets['Sheet1'], { header: 1, raw: true }) as any[][];

    expect(parsed.SheetNames).toEqual(['Sheet1']);
    expect(rows.length).toBe(2);
    expect(rows[1][1]).toBe(919818202888);
  });

  it('should parse csv lines using app delimiter detection and parser', () => {
    const fixture = TestBed.createComponent(AppComponent);
    const app = fixture.componentInstance as any;
    const csv = 'Name,Mobile Number\nA,919818202888\nB,919313123456\n';
    const lines = csv.trim().split(/\r\n|\n/);
    const delim = app.detectDelimiter(lines[0]);
    const rows = lines.map((line: string) => app.parseLine(line, delim));

    expect(delim).toBe(',');
    expect(rows.length).toBe(3);
    expect(rows[1][0]).toBe('A');
  });
});

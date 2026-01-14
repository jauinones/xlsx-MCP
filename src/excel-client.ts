import ExcelJS from "exceljs";
import { HyperFormula, ConfigParams, RawCellContent } from "hyperformula";

interface WorkbookEntry {
  workbook: ExcelJS.Workbook;
  hf: HyperFormula;
  path?: string;
}

interface SheetInfo {
  name: string;
  index: number;
  rowCount: number;
  columnCount: number;
  actualRowCount: number;
  actualColumnCount: number;
}

interface CellValue {
  value: unknown;
  calculatedValue: unknown;
  formula?: string;
}

interface ColumnInfo {
  column: string;
  header: string | null;
  index: number;
}

// HyperFormula configuration
const hfConfig: Partial<ConfigParams> = {
  licenseKey: "gpl-v3",
};

export class ExcelClient {
  private workbooks: Map<string, WorkbookEntry> = new Map();
  private nextId: number = 1;

  // Workbook Management

  createWorkbook(): string {
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet("Sheet1");

    // Create HyperFormula instance with one sheet
    const hf = HyperFormula.buildEmpty(hfConfig);
    hf.addSheet("Sheet1");

    const id = `wb_${this.nextId++}`;
    this.workbooks.set(id, { workbook, hf });
    return id;
  }

  async openWorkbook(filePath: string): Promise<string> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    // Build HyperFormula from workbook data
    const hf = HyperFormula.buildEmpty(hfConfig);

    // Sync each sheet to HyperFormula
    for (const sheet of workbook.worksheets) {
      const sheetId = hf.addSheet(sheet.name);
      if (sheetId === undefined) continue;

      // Get sheet data and sync to HyperFormula
      this.syncSheetToHyperFormula(sheet, hf, sheet.name);
    }

    const id = `wb_${this.nextId++}`;
    this.workbooks.set(id, { workbook, hf, path: filePath });
    return id;
  }

  async saveWorkbook(workbookId: string, filePath?: string): Promise<string> {
    const entry = this.getWorkbookEntry(workbookId);
    const savePath = filePath || entry.path;

    if (!savePath) {
      throw new Error("No file path specified and workbook has no associated path");
    }

    // Before saving, sync calculated values back to ExcelJS
    this.syncCalculatedValuesToWorkbook(entry);

    await entry.workbook.xlsx.writeFile(savePath);
    entry.path = savePath;
    return savePath;
  }

  closeWorkbook(workbookId: string): void {
    const entry = this.workbooks.get(workbookId);
    if (!entry) {
      throw new Error(`Workbook ${workbookId} not found`);
    }
    entry.hf.destroy();
    this.workbooks.delete(workbookId);
  }

  listWorkbooks(): Array<{ id: string; path?: string; sheetCount: number }> {
    const result: Array<{ id: string; path?: string; sheetCount: number }> = [];
    for (const [id, entry] of this.workbooks) {
      result.push({
        id,
        path: entry.path,
        sheetCount: entry.workbook.worksheets.length,
      });
    }
    return result;
  }

  // Sheet Operations

  listSheets(workbookId: string): Array<{ name: string; index: number }> {
    const entry = this.getWorkbookEntry(workbookId);
    return entry.workbook.worksheets.map((sheet, index) => ({
      name: sheet.name,
      index: index + 1,
    }));
  }

  createSheet(workbookId: string, name?: string): { name: string; index: number } {
    const entry = this.getWorkbookEntry(workbookId);
    const sheetName = name || `Sheet${entry.workbook.worksheets.length + 1}`;
    const sheet = entry.workbook.addWorksheet(sheetName);

    // Add sheet to HyperFormula
    entry.hf.addSheet(sheetName);

    return {
      name: sheet.name,
      index: entry.workbook.worksheets.length,
    };
  }

  deleteSheet(workbookId: string, sheetIdentifier: string | number): void {
    const entry = this.getWorkbookEntry(workbookId);
    const sheet = this.getSheet(entry.workbook, sheetIdentifier);
    const sheetName = sheet.name;

    entry.workbook.removeWorksheet(sheet.id);

    // Remove from HyperFormula
    const hfSheetId = entry.hf.getSheetId(sheetName);
    if (hfSheetId !== undefined) {
      entry.hf.removeSheet(hfSheetId);
    }
  }

  renameSheet(workbookId: string, sheetIdentifier: string | number, newName: string): void {
    const entry = this.getWorkbookEntry(workbookId);
    const sheet = this.getSheet(entry.workbook, sheetIdentifier);
    const oldName = sheet.name;

    sheet.name = newName;

    // Rename in HyperFormula
    const hfSheetId = entry.hf.getSheetId(oldName);
    if (hfSheetId !== undefined) {
      entry.hf.renameSheet(hfSheetId, newName);
    }
  }

  getSheetInfo(workbookId: string, sheetIdentifier: string | number): SheetInfo {
    const entry = this.getWorkbookEntry(workbookId);
    const sheet = this.getSheet(entry.workbook, sheetIdentifier);

    return {
      name: sheet.name,
      index: entry.workbook.worksheets.indexOf(sheet) + 1,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      actualRowCount: sheet.actualRowCount,
      actualColumnCount: sheet.actualColumnCount,
    };
  }

  // Cell Operations

  readCell(workbookId: string, sheetIdentifier: string | number, cellAddress: string): CellValue {
    const entry = this.getWorkbookEntry(workbookId);
    const sheet = this.getSheet(entry.workbook, sheetIdentifier);
    const cell = sheet.getCell(cellAddress);

    // Get calculated value from HyperFormula
    const hfSheetId = entry.hf.getSheetId(sheet.name);
    let calculatedValue: unknown = cell.value;

    if (hfSheetId !== undefined) {
      const ref = this.parseCellReference(cellAddress);
      const hfValue = entry.hf.getCellValue({ sheet: hfSheetId, row: ref.row - 1, col: ref.col - 1 });
      calculatedValue = hfValue;
    }

    return {
      value: cell.value,
      calculatedValue,
      formula: cell.formula ? `=${cell.formula}` : undefined,
    };
  }

  writeCell(
    workbookId: string,
    sheetIdentifier: string | number,
    cellAddress: string,
    value: string | number | boolean | Date
  ): void {
    const entry = this.getWorkbookEntry(workbookId);
    const sheet = this.getSheet(entry.workbook, sheetIdentifier);
    const cell = sheet.getCell(cellAddress);
    const ref = this.parseCellReference(cellAddress);

    // If value starts with '=', treat as formula
    if (typeof value === "string" && value.startsWith("=")) {
      cell.value = { formula: value.substring(1) };

      // Update HyperFormula with formula
      const hfSheetId = entry.hf.getSheetId(sheet.name);
      if (hfSheetId !== undefined) {
        entry.hf.setCellContents({ sheet: hfSheetId, row: ref.row - 1, col: ref.col - 1 }, value);
      }
    } else {
      cell.value = value;

      // Update HyperFormula with value
      const hfSheetId = entry.hf.getSheetId(sheet.name);
      if (hfSheetId !== undefined) {
        entry.hf.setCellContents({ sheet: hfSheetId, row: ref.row - 1, col: ref.col - 1 }, value);
      }
    }
  }

  readRange(
    workbookId: string,
    sheetIdentifier: string | number,
    range: string
  ): Array<Array<unknown>> {
    const entry = this.getWorkbookEntry(workbookId);
    const sheet = this.getSheet(entry.workbook, sheetIdentifier);

    const [startCell, endCell] = range.split(":");
    const startRef = this.parseCellReference(startCell);
    const endRef = this.parseCellReference(endCell);

    const hfSheetId = entry.hf.getSheetId(sheet.name);
    const result: Array<Array<unknown>> = [];

    for (let row = startRef.row; row <= endRef.row; row++) {
      const rowData: Array<unknown> = [];
      for (let col = startRef.col; col <= endRef.col; col++) {
        // Get calculated value from HyperFormula
        if (hfSheetId !== undefined) {
          const hfValue = entry.hf.getCellValue({ sheet: hfSheetId, row: row - 1, col: col - 1 });
          rowData.push(hfValue);
        } else {
          const cell = sheet.getCell(row, col);
          rowData.push(cell.value);
        }
      }
      result.push(rowData);
    }

    return result;
  }

  writeRange(
    workbookId: string,
    sheetIdentifier: string | number,
    startCell: string,
    values: Array<Array<string | number | boolean | Date | null>>
  ): void {
    const entry = this.getWorkbookEntry(workbookId);
    const sheet = this.getSheet(entry.workbook, sheetIdentifier);

    const startRef = this.parseCellReference(startCell);
    const hfSheetId = entry.hf.getSheetId(sheet.name);

    // Prepare data for HyperFormula batch update
    const hfValues: RawCellContent[][] = [];

    for (let rowOffset = 0; rowOffset < values.length; rowOffset++) {
      const rowData = values[rowOffset];
      const hfRowData: RawCellContent[] = [];

      for (let colOffset = 0; colOffset < rowData.length; colOffset++) {
        const cell = sheet.getCell(startRef.row + rowOffset, startRef.col + colOffset);
        const value = rowData[colOffset];

        if (value === null) {
          cell.value = null;
          hfRowData.push(null);
        } else if (typeof value === "string" && value.startsWith("=")) {
          cell.value = { formula: value.substring(1) };
          hfRowData.push(value);
        } else if (value instanceof Date) {
          cell.value = value;
          hfRowData.push(value.getTime());
        } else {
          cell.value = value;
          hfRowData.push(value as RawCellContent);
        }
      }
      hfValues.push(hfRowData);
    }

    // Batch update HyperFormula
    if (hfSheetId !== undefined) {
      entry.hf.setCellContents(
        { sheet: hfSheetId, row: startRef.row - 1, col: startRef.col - 1 },
        hfValues
      );
    }
  }

  listColumns(workbookId: string, sheetIdentifier: string | number): ColumnInfo[] {
    const entry = this.getWorkbookEntry(workbookId);
    const sheet = this.getSheet(entry.workbook, sheetIdentifier);

    const columns: ColumnInfo[] = [];
    const headerRow = sheet.getRow(1);

    for (let col = 1; col <= sheet.columnCount; col++) {
      const cell = headerRow.getCell(col);
      if (cell.value !== null && cell.value !== undefined) {
        columns.push({
          column: this.columnIndexToLetter(col),
          header: String(cell.value),
          index: col,
        });
      }
    }

    return columns;
  }

  // Formula Support

  getFormula(workbookId: string, sheetIdentifier: string | number, cellAddress: string): string | null {
    const entry = this.getWorkbookEntry(workbookId);
    const sheet = this.getSheet(entry.workbook, sheetIdentifier);
    const cell = sheet.getCell(cellAddress);

    return cell.formula ? `=${cell.formula}` : null;
  }

  recalculate(workbookId: string): { sheetsRecalculated: number; cellsRecalculated: number } {
    const entry = this.getWorkbookEntry(workbookId);

    // Re-sync all data to HyperFormula and recalculate
    let cellsRecalculated = 0;

    for (const sheet of entry.workbook.worksheets) {
      const hfSheetId = entry.hf.getSheetId(sheet.name);
      if (hfSheetId === undefined) continue;

      // Clear and re-sync sheet data
      const dimensions = entry.hf.getSheetDimensions(hfSheetId);
      if (dimensions.width > 0 && dimensions.height > 0) {
        cellsRecalculated += dimensions.width * dimensions.height;
      }
    }

    return {
      sheetsRecalculated: entry.workbook.worksheets.length,
      cellsRecalculated,
    };
  }

  // Charts

  createChart(
    workbookId: string,
    sheetIdentifier: string | number,
    options: {
      type: "bar" | "line" | "pie" | "scatter";
      dataRange: string;
      title?: string;
      position?: { col: number; row: number };
    }
  ): void {
    const entry = this.getWorkbookEntry(workbookId);
    const sheet = this.getSheet(entry.workbook, sheetIdentifier);

    // Store chart metadata (ExcelJS has limited chart support)
    // @ts-ignore - Adding custom chart metadata
    if (!entry.workbook.properties) {
      // @ts-ignore
      entry.workbook.properties = { date1904: false };
    }
    // @ts-ignore
    if (!entry.workbook.properties.charts) {
      // @ts-ignore
      entry.workbook.properties.charts = [];
    }
    // @ts-ignore
    entry.workbook.properties.charts.push({
      sheet: sheet.name,
      type: options.type,
      dataRange: options.dataRange,
      title: options.title,
      position: options.position,
    });
  }

  deleteChart(workbookId: string, sheetIdentifier: string | number, chartName: string): void {
    const entry = this.getWorkbookEntry(workbookId);

    // Remove from stored chart metadata
    // @ts-ignore
    if (entry.workbook.properties?.charts) {
      // @ts-ignore
      entry.workbook.properties.charts = entry.workbook.properties.charts.filter(
        (c: { title: string }) => c.title !== chartName
      );
    }
  }

  // Pivot Tables

  createPivotTable(
    workbookId: string,
    sourceSheet: string | number,
    options: {
      sourceRange: string;
      destinationSheet?: string;
      destinationCell?: string;
      rowFields: string[];
      columnFields?: string[];
      dataFields: Array<{ field: string; aggregation: "sum" | "count" | "average" | "min" | "max" }>;
    }
  ): { sheet: string; cell: string } {
    const entry = this.getWorkbookEntry(workbookId);
    const srcSheet = this.getSheet(entry.workbook, sourceSheet);

    // Read source data using HyperFormula for calculated values
    const [startCell, endCell] = options.sourceRange.split(":");
    const startRef = this.parseCellReference(startCell);
    const endRef = this.parseCellReference(endCell);

    const hfSheetId = entry.hf.getSheetId(srcSheet.name);

    // Get headers from first row
    const headers: string[] = [];
    for (let col = startRef.col; col <= endRef.col; col++) {
      if (hfSheetId !== undefined) {
        const val = entry.hf.getCellValue({ sheet: hfSheetId, row: startRef.row - 1, col: col - 1 });
        headers.push(String(val || `Column${col}`));
      } else {
        const cell = srcSheet.getCell(startRef.row, col);
        headers.push(String(cell.value || `Column${col}`));
      }
    }

    // Read data rows using calculated values
    const data: Array<Record<string, unknown>> = [];
    for (let row = startRef.row + 1; row <= endRef.row; row++) {
      const rowData: Record<string, unknown> = {};
      for (let col = startRef.col; col <= endRef.col; col++) {
        if (hfSheetId !== undefined) {
          const val = entry.hf.getCellValue({ sheet: hfSheetId, row: row - 1, col: col - 1 });
          rowData[headers[col - startRef.col]] = val;
        } else {
          const cell = srcSheet.getCell(row, col);
          rowData[headers[col - startRef.col]] = cell.value;
        }
      }
      data.push(rowData);
    }

    // Create or get destination sheet
    const destSheetName = options.destinationSheet || `PivotTable_${Date.now()}`;
    let destSheet: ExcelJS.Worksheet;

    try {
      destSheet = entry.workbook.addWorksheet(destSheetName);
      entry.hf.addSheet(destSheetName);
    } catch {
      destSheet = this.getSheet(entry.workbook, destSheetName);
    }

    const destCell = options.destinationCell || "A1";
    const destRef = this.parseCellReference(destCell);

    // Group data by row fields
    const grouped = new Map<string, Array<Record<string, unknown>>>();

    for (const row of data) {
      const key = options.rowFields.map(f => String(row[f] || "")).join("|");
      if (!grouped.has(key)) {
        grouped.set(key, []);
      }
      grouped.get(key)!.push(row);
    }

    // Write pivot table header
    let col = destRef.col;
    for (const field of options.rowFields) {
      destSheet.getCell(destRef.row, col).value = field;
      col++;
    }
    for (const dataField of options.dataFields) {
      destSheet.getCell(destRef.row, col).value = `${dataField.aggregation.toUpperCase()}(${dataField.field})`;
      col++;
    }

    // Write aggregated data
    let row = destRef.row + 1;
    for (const [key, rows] of grouped) {
      col = destRef.col;
      const keyParts = key.split("|");

      // Write row field values
      for (const part of keyParts) {
        destSheet.getCell(row, col).value = part;
        col++;
      }

      // Calculate and write aggregations
      for (const dataField of options.dataFields) {
        const values = rows
          .map(r => r[dataField.field])
          .filter(v => typeof v === "number") as number[];

        let result: number = 0;
        switch (dataField.aggregation) {
          case "sum":
            result = values.reduce((a, b) => a + b, 0);
            break;
          case "count":
            result = rows.length;
            break;
          case "average":
            result = values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0;
            break;
          case "min":
            result = values.length > 0 ? Math.min(...values) : 0;
            break;
          case "max":
            result = values.length > 0 ? Math.max(...values) : 0;
            break;
        }

        destSheet.getCell(row, col).value = result;
        col++;
      }

      row++;
    }

    // Sync destination sheet to HyperFormula
    this.syncSheetToHyperFormula(destSheet, entry.hf, destSheetName);

    return {
      sheet: destSheetName,
      cell: destCell,
    };
  }

  // Helper Methods

  private syncSheetToHyperFormula(sheet: ExcelJS.Worksheet, hf: HyperFormula, sheetName: string): void {
    const hfSheetId = hf.getSheetId(sheetName);
    if (hfSheetId === undefined) return;

    // Build 2D array of cell values
    const data: RawCellContent[][] = [];

    sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      while (data.length < rowNumber) {
        data.push([]);
      }
      const rowData: RawCellContent[] = [];

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        while (rowData.length < colNumber - 1) {
          rowData.push(null);
        }

        if (cell.formula) {
          rowData.push(`=${cell.formula}`);
        } else {
          // Convert ExcelJS cell value to HyperFormula compatible type
          const value = cell.value;
          if (value === null || value === undefined) {
            rowData.push(null);
          } else if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
            rowData.push(value);
          } else if (value instanceof Date) {
            rowData.push(value.getTime());
          } else {
            // For complex objects, convert to string
            rowData.push(String(value));
          }
        }
      });

      data[rowNumber - 1] = rowData;
    });

    if (data.length > 0) {
      hf.setSheetContent(hfSheetId, data);
    }
  }

  private syncCalculatedValuesToWorkbook(entry: WorkbookEntry): void {
    // For cells with formulas, ensure the cached result is updated
    for (const sheet of entry.workbook.worksheets) {
      const hfSheetId = entry.hf.getSheetId(sheet.name);
      if (hfSheetId === undefined) continue;

      sheet.eachRow((row) => {
        row.eachCell((cell) => {
          if (cell.formula) {
            const ref = this.parseCellReference(cell.address);
            const calculatedValue = entry.hf.getCellValue({
              sheet: hfSheetId,
              row: ref.row - 1,
              col: ref.col - 1,
            });

            // Update the cell's result value
            if (typeof cell.value === "object" && cell.value !== null && "formula" in cell.value) {
              (cell.value as { formula: string; result?: unknown }).result = calculatedValue;
            }
          }
        });
      });
    }
  }

  private getWorkbookEntry(workbookId: string): WorkbookEntry {
    const entry = this.workbooks.get(workbookId);
    if (!entry) {
      throw new Error(`Workbook ${workbookId} not found`);
    }
    return entry;
  }

  private getSheet(workbook: ExcelJS.Workbook, identifier: string | number): ExcelJS.Worksheet {
    let sheet: ExcelJS.Worksheet | undefined;

    if (typeof identifier === "number") {
      sheet = workbook.worksheets[identifier - 1];
    } else {
      sheet = workbook.getWorksheet(identifier);
    }

    if (!sheet) {
      throw new Error(`Sheet "${identifier}" not found`);
    }

    return sheet;
  }

  private parseCellReference(ref: string): { row: number; col: number } {
    const match = ref.match(/^([A-Z]+)(\d+)$/i);
    if (!match) {
      throw new Error(`Invalid cell reference: ${ref}`);
    }

    const colStr = match[1].toUpperCase();
    const row = parseInt(match[2], 10);

    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
      col = col * 26 + (colStr.charCodeAt(i) - 64);
    }

    return { row, col };
  }

  private columnIndexToLetter(index: number): string {
    let result = "";
    while (index > 0) {
      const remainder = (index - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      index = Math.floor((index - 1) / 26);
    }
    return result;
  }
}

/**
    Coded by Dave Cook
    www.davecookcodes.com
*/

namespace SpreadsheetManagerTypes {
  export interface Options {
    headerRow: number;
    lastColumn: number;
  }
  export interface RowHeaders {
    [key: string]: number;
  }
}

interface SpreadsheetManager {
  row: Array<string | number | Date>;
  wb: GoogleAppsScript.Spreadsheet.Spreadsheet;
  sheet: GoogleAppsScript.Spreadsheet.Sheet | null;
  values: Array<string | number | Date>[];
  rowHeaders: SpreadsheetManagerTypes.RowHeaders;
  headerRow: number;
  lastColumn: number;
}

type SpreadsheetManagerGenericRowValue = string | number | Date | undefined;

interface SpreadsheetManagerGenericRowObject {
  [key: string]: SpreadsheetManagerGenericRowValue;
}

class SpreadsheetManager {
  constructor(
    wb: GoogleAppsScript.Spreadsheet.Spreadsheet,
    sheetName: string,
    options?: SpreadsheetManagerTypes.Options
  ) {
    const headerRow: number = options ? options.headerRow : 1;

    this.headerRow = headerRow;
    this.wb = wb;
    this.sheet = this.wb.getSheetByName(sheetName);
    const lastColumn: number = options
      ? options.lastColumn
      : (this.sheet?.getLastColumn() as number);
    this.lastColumn = lastColumn;
    if (!this.sheet) return;
    this.values = this.getSheetValues();
    this.rowHeaders = this.getRowHeaders(this.values[0]);
  }

  /**
   *
   *
   * @param variable[][] or variable[] rows
   * @memberof SpreadsheetManager
   */
  addNewRows(rows: Array<string | number | Date>[]) {
    const { sheet } = this;
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length);
    range.setValues(rows);
  }

  /**
   * @param Object[] Array of object representing data to be entered into each row
   * each attribute of each object must be equivalent to an attribute in rowheaders
   * @memberof SpreadsheetManager
   */
  addNewRowsFromObjects(objects: SpreadsheetManagerGenericRowObject[]) {
    const { rowHeaders } = this;
    const newRows = objects.map((obj) => {
      const newRow: Array<string | number | Date> = [];
      for (let header in rowHeaders) {
        const colIndex: number = rowHeaders[header];
        newRow[colIndex] = obj[header] || "";
      }
      return newRow;
    });
    this.addNewRows(newRows);
  }

  /**
   *
   *
   * @param _Row[] row
   * @returns object of values with column headers as keys
   * @memberof SpreadsheetManager
   */
  createObjectFromRow(row: _Row) {
    const { rowHeaders } = this;
    const obj: { [key: string]: string | number | Date | undefined } = {};
    for (let key in rowHeaders) {
      try {
        obj[key] = row.col(key);
      } catch (err) {
        Logger.log(err);
      }
    }
    return obj;
  }

  clearSheet(): void {
    if (!this.sheet) return;
    const { sheet, headerRow } = this;
    sheet
      .getRange(headerRow + 1, 1, sheet.getLastRow(), sheet.getLastColumn())
      .clearContent();
    SpreadsheetApp.flush();
  }

  /**
   *
   *
   * @memberof SpreadsheetManager
   */
  clearSheetAndPasteValues() {
    if (!this.sheet) return;
    const { sheet, values } = this;
    sheet.getDataRange().clearContent();
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    SpreadsheetApp.flush();
  }

  /**
   *
   * @desc loops through all rows
   * @param {function} callback
   * @param {object} options can specify 'bottomUp' as true to reverse direction of loop
   * @memberof SpreadsheetManager
   */
  forEachRow(callback: Function, options?: { bottomUp?: boolean }) {
    if (options?.bottomUp) {
      for (let i = this.values.length - 1; i > 0; i--) {
        const row = new _Row(
          this.values[i],
          this.rowHeaders,
          this,
          i + this.headerRow
        );
        const val = callback(row, i);
        if (val) return val;
      }
    } else {
      for (let i = 1; i < this.values.length; i++) {
        const row = new _Row(
          this.values[i],
          this.rowHeaders,
          this,
          i + this.headerRow
        );
        const val = callback(row, i);
        if (val) return val;
      }
    }
  }
  /**
   *
   *
   * @returns array
   * @memberof SpreadsheetManager
   */
  getLastRow() {
    const { values } = this;
    const lastRowIndex = values.length - 1;
    if (lastRowIndex >= 0) {
      return values[lastRowIndex];
    }
  }
  /**
   * @desc creates an array to reference column number by header name
   * @param string[] topRow
   * @return obj - {header:int,header:int,...}
   */
  getRowHeaders(
    topRow: Array<string | number | Date>
  ): SpreadsheetManagerTypes.RowHeaders {
    const obj: SpreadsheetManagerTypes.RowHeaders = {};
    for (let c = 0; c < topRow.length; c++) {
      //removes line breaks and multiple spaces
      const cell: string = String(topRow[c])
        .replace(/(\r\n|\n|\r)/gm, " ")
        .replace(/\s\s+/g, " ");
      obj[cell] = c;
    }
    return obj;
  }

  getRowsAsObjects() {
    const objects: any[] = [];
    this.forEachRow((row: _Row) => {
      const obj = row.createObject();
      objects.push(obj);
    });
    return objects;
  }
  /**
   * @desc sets values attribute for object
   * @return array of data from sheet
   */
  getSheetValues(): Array<string | number | Date>[] {
    if (!this.sheet) return [[]];
    const lastRow = this.sheet.getLastRow() + 1;
    const lastColumn = this.lastColumn;
    const values: Array<string | number | Date>[] = this.sheet
      .getRange(this.headerRow, 1, lastRow - this.headerRow, lastColumn)
      .getValues();
    return values;
  }
  /**
   * @desc gets values in column by column header name
   * @param string  headerName
   * @param bool valuesOnly = when true, function returns 1d array. When false, 2d array
   * @return array of data from sheet
   */
  getValuesInColumn(headerName: string, valuesOnly = false) {
    const { values, rowHeaders } = this;
    if (rowHeaders.hasOwnProperty(headerName)) {
      const columnIndex = rowHeaders[headerName];

      return values.slice(1).map((row) => {
        const cell = valuesOnly ? row[columnIndex] : [row[columnIndex]];
        return cell;
      });
    } else {
      Logger.log(`${headerName} not found in row headers`);
      return false;
    }
  }
  /**
   * @desc paste formatted column into sheet by header name
   * @param string  headerName
   */
  pasteValuesToColumn(
    headerName: string,
    columnArray: Array<string | number | Date>[]
  ) {
    if (!this.sheet) return;
    const { sheet, rowHeaders } = this;
    if (rowHeaders.hasOwnProperty(headerName)) {
      const columnIndex = rowHeaders[headerName];

      const pasteRange = sheet.getRange(
        2,
        columnIndex + 1,
        columnArray.length,
        1
      );
      const pasteAddress = pasteRange.getA1Notation();
      pasteRange.setValues(columnArray);
    } else {
      Logger.log(`${headerName} not found in row headers`);
      return false;
    }
  }
  /**
   * @desc updates sheet with values from this.values;
   */
  updateAllValues() {
    if (!this.sheet) return;
    const { values, sheet } = this;
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    SpreadsheetApp.flush();
  }
}

interface _Row {
  _rowIndex: number;
  values: Array<string | number | Date>;
  headers: SpreadsheetManagerTypes.RowHeaders;
  parent: SpreadsheetManager;
}
class _Row {
  /**
   *Creates an instance of _Row.
   * @param string[] row
   * @param object headers
   * @memberof _Row
   */
  constructor(
    row: Array<string | number | Date>,
    headers: SpreadsheetManagerTypes.RowHeaders,
    parent: SpreadsheetManager,
    rowIndex: number
  ) {
    this._rowIndex = rowIndex;
    this.values = row;
    this.headers = headers;
    this.parent = parent;
  }

  createObject() {
    const { values, headers } = this;
    const obj: { [key: string]: any } = {};
    for (let header in headers) {
      const index = headers[header];
      obj[header] = values[index];
    }
    return obj;
  }

  col(headerName: string, value?: any): string | number | Date {
    const colIndex = this.headers[headerName];
    try {
      if (value) {
        this.values[colIndex] = value;
        return value;
      } else {
        return this.values[colIndex];
      }
    } catch (err) {
      Logger.log(`${headerName} isn't a column`, err);
      return "";
    }
    return "";
  }

  update() {
    const { parent } = this;
    parent.sheet
      ?.getRange(this._rowIndex, 1, 1, this.values.length)
      .setValues([this.values]);
  }
}

// module.exports = SpreadsheetManager;

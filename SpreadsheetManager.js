/**
    Coded by Dave Cook
    www.davecookcodes.com
*/

class SpreadsheetManager {
  constructor(wb, sheetName, options) {
    const headerRow = options ? options.headerRow : 1;
    this.headerRow = headerRow;
    this.wb = wb;
    this.sheet = this.wb.getSheetByName(sheetName);
    if (!this.sheet) return;
    this.values = this.getSheetValues(headerRow);
    this.rowHeaders = this.getRowHeaders(this.values[0]);
  }
  

  /**
   *
   *
   * @param variable[][] or variable[] rows
   * @memberof SpreadsheetManager
   */
  addNewRows(rows) {
    if (!typeof rows[0] === "object") {
      rows = [rows];
    }
    const { sheet } = this;
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length);
    range.setValues(rows);
  }

  /**
   * @param Object[] Array of object representing data to be entered into each row
   * each attribute of each object must be equivalent to an attribute in rowheaders
   * @memberof SpreadsheetManager
   */
  addNewRowsFromObjects(objects = []){
    const { rowHeaders } = this;
    const newRows =  objects.map(obj => {
      const newRow = [];
      for (let header in rowHeaders){
        const colIndex = rowHeaders[header];
        newRow[colIndex] = obj[header] || '';
      }
      return newRow;
    })

    t

  /**
   *
   *
   * @param _Row[] row
   * @returns object of values with column headers as keys
   * @memberof SpreadsheetManager
   */
  createObjectFromRow(row) {
    const { rowHeaders } = this;
    const obj = {};
    for (let key in rowHeaders) {
      try {
        obj[key] = row.col(rowHeaders[key]);
      } catch (err) {
        Logger.log(err);
      }
    }
    return obj;
  }

  /**
   *
   *
   * @memberof SpreadsheetManager
   */
  clearSheetAndPasteValues() {
    const { sheet, values } = this;
    sheet.getDataRange().clearContent();
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    SpreadsheetApp.flush();
  }

  /**
   *
   * @desc loops through all rows
   * @param {*} callback
   * @memberof SpreadsheetManager
   */
  forEachRow(callback) {
    for (let i = 1; i < this.values.length; i++) {
      const row = new _Row(this.values[i], this.rowHeaders);
      callback(row, i);
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
  getRowHeaders(topRow) {
    const obj = {};
    for (let c = 0; c < topRow.length; c++) {
      //removes line breaks and multiple spaces
      const cell = topRow[c]
        .replace(/(\r\n|\n|\r)/gm, " ")
        .replace(/\s\s+/g, " ");
      obj[cell] = c;
    }
    return obj;
  }
  /**
   * @desc sets values attribute for object
   * @return array of data from sheet
   */
  getSheetValues() {
    const lastRow = this.sheet.getLastRow() + 1;
    const lastColumn = this.sheet.getLastColumn();
    const values = this.sheet
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
  getValuesInColumn(headerName, valuesOnly = false) {
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
  pasteValuesToColumn(headerName, columnArray) {
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
    const { values, sheet } = this;
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    SpreadsheetApp.flush();
  }
}

class _Row {
  /**
   *Creates an instance of _Row.
   * @param string[] row
   * @param object headers
   * @memberof _Row
   */
  constructor(row, headers) {
    this.values = row;
    this.headers = headers;
  }

  createObject() {
    const { values, headers } = this;
    const obj = {};
    for (let header in headers) {
      const index = headers[header];
      obj[header] = values[index];
    }
    return obj;
  }

  col(headerName) {
    const colIndex = this.headers[headerName];
    try {
      return this.values[colIndex];
    } catch (err) {
      Logger.log(`${headerName} isn't a column in ${row.toString()}`, err);
    }
  }
}

// module.exports = SpreadsheetManager;

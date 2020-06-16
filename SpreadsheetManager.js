class SpreadsheetManager{
  constructor(wb, sheetName) {
    this.wb = wb;
    this.sheet = this.wb.getSheetByName(sheetName);
    this.values = this.getSheetValues();
    this.rowHeaders = this.getRowHeaders(this.values[0]);
  }
  /**
    * @desc creates an array to reference column number by header name
    * @param string[] topRow
    * @return obj - {header:int,header:int,...}
  */
  getRowHeaders(topRow){
    const obj = {};
    for (let c = 0; c < topRow.length; c++){
      const cell = topRow[c];
      obj[cell] = c;
    }
    return obj;
  }
  /**
    * @desc sets values attribute for object
    * @return array of data from sheet
  */
  getSheetValues(){
    const { sheet } = this;
    const values = sheet.getDataRange().getValues();
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
    if (rowHeaders.hasOwnProperty(headerName)){
      const columnIndex = rowHeaders[headerName];
      
      return values.slice(1,).map(row => {
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
  pasteValuesToColumn(headerName, columnArray){
    const { sheet, rowHeaders } = this;
    if (rowHeaders.hasOwnProperty(headerName)){
      const columnIndex = rowHeaders[headerName];
      
      const pasteRange = sheet.getRange(2,columnIndex + 1,columnArray.length,1);
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
    sheet.getRange(1,1,values.length,values[0].length).setValues(values);
  }
  
}

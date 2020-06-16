# SpreadsheetManager
A useful class to import into Google Apps Script projects when working with Google Sheets.

To use this in your project add your workbook and the sheet name as parameters:

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const managedSheet = new SpreadsheetManager(ss, "Sheet1");
    
When looping through rows, use column headers to make referencing cells easier

    const h = managedSheet.rowHeaders;
    managedSheet.values.forEach(row => {
      const id = row[h.id]; //column title is 'id'
      const name = row[h.name]; //column title is 'name'
    });

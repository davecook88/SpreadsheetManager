**I am available for freelance work or full time opportunities**

🌎 davecookcodes.com
Upwork: https://www.upwork.com/freelancers/~0141d415013a6613d0

# Spreadsheet Manager

The first class to import into any Google Apps Script project using Google Sheets!

I was fed up of typing out the same lines of code with every project so created this class to abstract out any repetitive logic to make life much easier when working in Google Apps Script ⏩

## Init 🚀

To use this in your project add your workbook and the sheet name as parameters:

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const managedSheet = new SpreadsheetManager(ss, "Sheet1");

The default is for the column headers to be in row 1 of your spreadsheet. If this is not the case, set an alternative header row by adding an additional parameter to the constructor.

    const managedSheet = new SpreadsheetManager(ss, "Sheet1", {headerRow:2});
    
## Updating values ✏
    
Instead of always having to call `getDataRange()` and then `getValues()` when working with a new sheet, this is all taken care of in the class constructor.

The `SpreadsheetManager.values` attribute holds a 2d array of your spreadsheets values for you to manipulate without updating the sheet.

When you are finished, run `SpreadsheetManager.updateAllValues()` to apply the changes to the sheet. In my scripts, this is nearly always the last line in my main functions. This method also calls `SpreadsheetApp.flush()` so it can be used at any point in the script without running into issues where the sheet hasn't fully updated before the next line is run.

## Loop through rows 🔁

The number one use of Google Apps Script in Sheets is iterating through all of the rows in a dataset to get or update values. Normally this is done in a for loop with the column indexes referenced by number like so:

    for (let rowNumber = 0; rowNumber < dataValues.length; rowNumber++){
        var valueFromColumnA = dataValues[rowNumber][0]
        var valueFromColumnB = dataValues[rowNumber][1]
        var valueFromColumnC = dataValues[rowNumber][2]
    }
    
 Not only is this cumbersome but it is also hard to maintain. If columns are added to or removed from the sheet, you have to update all of the column numbers. It's especially difficult to map the numbers to the column letters.
 
SpreadsheetManager makes this so much easier. The `forEachRow()` function allows you to iterate through the rows via a callback function to which a custom `_Row` class is passed. `_Row` contains the `col` method which references cells by column name. For example:

    managedSpreadsheet.forEachRow((row, rowIndex) => {
        const id = row.col('id'); // gets the value from the 'id' column in the sheet
        const name = row.col('name'); // gets the value from the 'name' column in the sheet
        
        // To update a value, simply pass a second argument to the col method;
        const newPrice = 1000;
        row.col('price', newPrice);
        // This change is only applied to the managedSpreadsheet.values property. Don't forget to call updateAllValues() at the and to apply changes to the sheet.
    });
    
## Rows as objects 👩‍💻

By default, Google Apps Script treats spreadsheets as 2d arrays. SpreadsheetManager allows us to pass objects into a sheet to make it easier to keep track of what values mean and to allow for spreadsheets to be updated without breaking scripts.

#### Get a row as an object

In the `_Row` object described above, it is possible to get an entire row's values as a single object:

    managedSpreadsheet.forEachRow((row, rowIndex) => {
        const rowObj = row.createObject();
        // now the rowObj acts as a dictionary with values mapped to column headers
        
        const newPrice = rowObj.price + 10;
        // note any changes to the rowObj will not be reflected in the values property of the SpreadsheetManager class.
    });
    
####  Adding objects to rows

The `addNewRowsFromObjects` method allows you to pass an array of objects directly to the sheet's values, without needing to map anything. If the object's keys match the column headers, that row will be updated with the relevant value. If the key doesn't exist, it is ignored.

This is particularly useful when working with API responses.

    const response = UrlFetchApp('http.....'); // Let's get an array of users from an API
    const json = JSON.parse(response);
    \*
    Assuming that the attributes in the API response match the column headers in our sheet, we can pass the response straight in.
    If not, mapping to an array of objects is just one additional step.
    */
    managedSpreadsheet.addNewRowsFromObjects(json.users);
    managedSpreadsheet.updateAllValues()
    
## Values in a column 📈

If you only need the values from a single column, `getValuesInColumn` makes life easy. Simply enter the column header name and decide if you want the results as a 2d or flat array.
    
    const ids_2d_array = managedSpreadsheet.getValuesInColumn('id') // returns all values from the id column [[1],[2],[3]] 
    const ids_flat_array = managedSpreadsheet.getValuesInColumn('id', true) // returns [1,2,3] 
    
To just update a single column, use `pasteValuesToColumn`.
    
    // Changing ids from 1,2,3 to 10001, 10002, 10003
    const newIds = ids_flat_array.map(oldId => [oldId + 10000]);
    const newIdColumn = managedSpreadsheet.getValuesInColumn('id', newIds)
    
# I hope this helps! 😀

Please feel free to get in touch with feedback or make pull requests if you see any bugs or want to add new features
    

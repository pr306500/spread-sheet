# Spread Sheet

  The spread sheet helps in by adding rows and reading rows by providing 'from_row', 'to_row' params present in spread sheet.


##In order to add a row, a person must following info :-

1. Rows (you want to add in spread sheet)
2. Filepath (path where your spreadsheet resides)
3. Sheet Name (Name of your sheet where yoou want to add row)


### Detailed steps

#### Step 1. After installing "spread_sheet" module, do [var spreadSheet = require('spread_sheet')].

#### Step 2. spreadSheet.addRow(row,filePath,sheetName,function(err,result){}).

#### Step 3. Arguments behaviour:-

      1. Row

         It must be string.

         var row = [[1,2,3],['a',45,56]]; // To add multiple rows.

      2. filePath

         var filePath = '/home/pranjal/Desktop/test.xlsx';

      3. sheetName

         var sheetName = 'Sheet1';

      4. Last argument is the callback (cb), which accepts error and result.   


## In order to get rows, follow the following steps :-

1. Filepath (path where your spreadsheet resides)
2. Sheet Name (Name of your sheet where yoou want to add row)
3. From Row (Add value of row ex. 1 or 5 .., to get info from specific row)
4. To Row (Add value of row ex. 1,2 or 6 .., to get info from specific row)


### Detailed steps

#### Step 1. After installing "spread_sheet" module, do [var spreadSheet = require('spread_sheet')].

#### Step 2. spreadSheet.getRows(filepath,sheet_name,from_row,to_row,function(err,result){}).

#### Step 3. Arguments behaviour:-

      1. fromRow

         It must be string/numeric.

         var from_row = 4;

      2. filePath

         var filePath = '/home/pranjal/Desktop/test.xlsx';

      3. sheetName

         var sheetName = 'Sheet1';

      4. toRow

         It must be string/numeric.

         var to_row = 8;  

      5. Last argument is the callback (cb), which accepts error and result.   

 It will fetch only those rows mentioned in from_row to to_row from spreadsheet.     

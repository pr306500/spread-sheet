exports.addRow = function(row,filepath,sheet_name,output){

var XLSX = require('xlsx');

    var data = String(row).split(',').map((i)=>{
      return i.trim();
    }).filter((i)=>{
      return i;
    });

    data = Array(data);

    var ws_name = "SheetJS";
    var workbook;
    var path = filepath;
    var last_column;

    /*sheet_from_array_of_arrays function will
      convert the row we want to add into worksheet format(ws),
      in order to make the workbook.
    */

    function sheet_from_array_of_arrays(data) {
     if (!Array.isArray(data)) {
       return output('Row entered is not in proper format')
     }
     var ws = {};
     var range = {
       s: {
         c: 10000000,
         r: 10000000
       },
       e: {
         c: 0,
         r: 0
       }
     };
     for (var R = 0; R != data.length; ++R) {
       for (var C = 0; C != data[R].length; ++C) {
         if (range.s.r > R) range.s.r = R;
         if (range.s.c > C) range.s.c = C;
         if (range.e.r < R) range.e.r = R;
         if (range.e.c < C) range.e.c = C;
         var cell = {
           v: data[R][C]
         };
         if (cell.v == null) continue;
         var cell_ref = XLSX.utils.encode_cell({
           c: C,
           r: R
         });

         if (typeof cell.v === 'number') cell.t = 'n';
         else if (typeof cell.v === 'boolean') cell.t = 'b';
         else if (cell.v instanceof Date) {
           cell.t = 'n';
           cell.z = XLSX.SSF._table[14];
           cell.v = datenum(cell.v);
         } else cell.t = 's';

         ws[cell_ref] = cell;
       }
     }
     if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
     return ws;
   }

   /*Through this function we will convert the worksheet (ws)
     into the workbook format*/

   function Workbook() {
     if (!(this instanceof Workbook)) return new Workbook();
     this.SheetNames = [];
     this.Sheets = {};
   }

   var wb = new Workbook(),
     ws = sheet_from_array_of_arrays(data);
   wb.SheetNames.push(ws_name);
   wb.Sheets[ws_name] = ws;


   format_workbook(path, function(err, wb) {

     Object.keys(wb.Sheets.SheetJS).forEach(function(m) {
       var len = Object.keys(wb.Sheets.SheetJS);

       if (len[len.length - 1] === m) {
         delete wb.Sheets.SheetJS[m]
       }
     });

     modify_final_workbook(wb, workbook, function(err, workbook) {
       if (!err) {
         if (!XLSX.writeFile(workbook, path)) {
           workbook = XLSX.readFile(path);
           workbook_new_state = Object.keys(workbook.Sheets[sheet_name]).length;
           if (workbook_new_state > workbook_old_state) {

             return output(null, {
               result: 'New Row Added'
             })

           }
         }
       }
     })

   });

   /*In this function we will first read the spread sheet and the format will be as
     per workbook standard (wb), which will be taken care by the readFile method of
    XLSX object.
    */
   function format_workbook(path, cb) {

     workbook = XLSX.readFile(path);
     if (Object.keys(workbook.Sheets[sheet_name]).length < 2) {
       workbook_old_state = Object.keys(workbook.Sheets[sheet_name]).length;
       if (!XLSX.writeFile(wb, path)) {

         workbook = XLSX.readFile(path);
         workbook_new_state = Object.keys(workbook.Sheets[sheet_name]).length;
         if (workbook_new_state > workbook_old_state) {

           return output(null, {
             result: 'New Row Added'
           })

         }
       }
     } else {

       workbook_old_state = Object.keys(workbook.Sheets[sheet_name]).length;
       var key_array = Object.keys(workbook.Sheets[sheet_name]);

       last_column = key_array[key_array.length - 2].split('');
       check_condition(last_column, cb);

     }


   }

   /*This function will check if the row entered at position greater than
     A10, then this function will get implemented*/
   function check_condition(last_column, cb) {

     var new_arr = [];
     if (last_column.length > 2) {
       for (var i = 1; i < last_column.length; i++) {
         new_arr.push(last_column[i])
       }
       last_column = new_arr.join('');

       delete workbook.Sheets[sheet_name]["!ref"];
       modifiy_column_name(wb, last_column, cb);

     } else {
       last_column = last_column[last_column.length - 1];

       delete workbook.Sheets[sheet_name]["!ref"];
       modifiy_column_name(wb, last_column, cb);

     }
   }

   /*This function will convert the row we want to add
     as per the readFile data, which means that if readFile brings the last column as
     'D9', the modifiy_column_name function will add the new row by first converting from
     "A1" to A10.*/

   function modifiy_column_name(wb, last_column, cb) {

     Object.keys(wb.Sheets[Object.keys(wb.Sheets)]).forEach(function(m) {

       var val = m.split('');

       var index = val.indexOf(val[val.length - 1]);
       if (index) {

         val[index] = parseInt(last_column) + 1;

         m_new = val.join('');

         wb.Sheets.SheetJS[m_new] = wb.Sheets.SheetJS[m];
         delete wb.Sheets.SheetJS[m];

       }
     });

     return cb(null, wb)

   }
   /*This function will take care of final formatting of the workbook after adding the new
     row with the total rows we fetched ftom readFile method and will add the field by manipulating
     the final workbook eg. "A1:D10", where A1 is the starting column and D10 is the last column.*/
   function modify_final_workbook(wb, workbook, cb) {

     Object.keys(wb.Sheets.SheetJS).forEach(function(m) {
         console.log('hello',Object.keys(workbook.Sheets))
       workbook.Sheets[sheet_name][m] = wb.Sheets.SheetJS[m]

     });
     var val = Object.keys(workbook.Sheets[sheet_name]);

     var ref = val[0] + ":" + val[val.length - 1];
     workbook.Sheets[sheet_name]["!ref"] = ref;

     return cb(null, workbook);
   }

}
exports.getRows = function(filepath,sheet_name,from_row,to_row,output){
  var XLSX = require('xlsx');
  var first_cell = 1;
    var allRow = [];
    var eachRow = [];
    var spreadsheet = [];

    wb = XLSX.readFile(filepath);
    wb = wb['Sheets'];
    wb = wb[sheet_name];

    if(typeof(wb) === 'undefined'){
      return output('Sheet name does not exist, please enter correct sheet name');
    }

    delete wb["!ref"];

    if (isNaN(from_row) || isNaN(to_row)) {
      return output('From Row and To Row field should be a number');
    }

    from_row = Number(from_row);
    to_row = Number(to_row);

    if (from_row > to_row) {
      return output('From Row should be lower or an equal to To Row field')
    }

    Object.keys(wb).forEach(function(row_column_pair) {
      num = row_column_pair.replace(/[^\d.]/g, '');
      if (first_cell == parseInt(num)) {
        eachRow.push(row_column_pair);
      }else if (first_cell != parseInt(num)) {
        allRow.push(eachRow);
        eachRow = [];
        first_cell = parseInt(row_column_pair.replace(/[^\d.]/g, ''));
        eachRow.push(row_column_pair);
      }

    });
    allRow.push(eachRow);
    allRow.forEach(function(individual_row, i) {
      eachRow = [];
      individual_row.forEach(function(each_cell, j) {
        eachRow.push(wb[each_cell][Object.keys(wb[each_cell])[Object.keys(wb[each_cell]).length - 1]]);
      });
      spreadsheet.push(eachRow);
    });

    if (spreadsheet.length > 0) {
      final_sheet = [];
      for (var i = from_row - 1; i <= to_row - 1; i++) {
        final_sheet.push(spreadsheet[i]);
      }
      return output(null, {
        response: final_sheet
      });
    }
}

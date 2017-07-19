exports.addRow = function(row,filepath,sheet_name,cb){

var XLSX = require('xlsx');
var excelrange = require('excelrange');
  var mapAlpha = {
    A:1,
    B:2,
    C:3,
    D:4,
    E:5,
    F:6,
    G:7,
    H:8,
    I:9,
    J:10,
    K:11,
    L:12,
    M:13,
    N:14,
    O:15,
    P:16,
    Q:17,
    R:18,
    S:19,
    T:20,
    U:21,
    V:22,
    W:23,
    X:24,
    Y:25,
    Z:26
  };

  var data = row

  var wb = XLSX.readFile(filepath);
  var worksheet = wb.Sheets[sheet_name]
  var ref = worksheet['!ref']
  var rowNumber = 1
  var keyPrefix = null
  var refDefined = ref ? true : false

  if(refDefined){
    var lastRow = ref.match(/(\d+$)/)

    if(lastRow){
      rowNumber = Number(lastRow[0]) + 1
    }
      ref = ref.split(':')
    keyPrefix = ref[1].match(/^([a-zA-Z])+/)
    if(keyPrefix){
      keyPrefix = keyPrefix[0];
    }
  }else{
    ref = ['A1', 'A1']
  }

  data.forEach((row)=>{
    //app.logger.info(row)
    row.forEach((cell, index)=>{
      var key = excelrange(index+1);
      worksheet[key + rowNumber] = {
        "t": isNaN(cell) ? "s" : "n",
        "v": cell,
        "h": cell,
        "w": cell
      }
      //app.logger.info((key + rowNumber)  +  cell)
      ref[1] = (
        refDefined ?
        bigger(key, keyPrefix, (index+1)) :
        (keyPrefix ?
          keyPrefix
          : key )
        )+ rowNumber
    })
    ++rowNumber
  });

  worksheet['!ref'] = ref.join(':')
  //app.logger.info(worksheet)
  XLSX.writeFile(wb, filepath);
  return cb(null, {
    result: 'New Row Added'
  })

function bigger(a, b, aVal, bVal){
if(!aVal){
  aVal = a.split('').reduce((initial, i, j, k)=>{
    return (initial + mapAlpha[i] + ((j < (k.length-1)) ? (26 - mapAlpha[i]): 0))
  }, 0)
}
if(!bVal){
  bVal = b.split('').reduce((initial, i, j, k)=>{
    return (initial + mapAlpha[i] + ((j < (k.length-1)) ? (26 - mapAlpha[i]): 0))
  }, 0)
}
if(aVal > bVal){
  return a
}
return b
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

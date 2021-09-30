 var XLSX = require("xlsx")
 var wb = XLSX.readFile("RkNumber.xlsx")
 var ws = wb.Sheets.RkNumber;
 XLSX.utils.sheet_to_json(ws, {header:1}) // this is the full sheet
[ [ 'fX100', 'fInt', 'sample' ],
  [ '0', '0', '10000000' ],
  [ '0', '1', '1200455' ],
  [ '1', '0', '0.01' ],
  [ '1', '1', '12004.55' ] ]
 ws['!ref'] = "A2:B4" // change the sheet range to A2:B4
 XLSX.utils.sheet_to_json(ws, {header:1}) // will only use cells within the new range
[ [ '0', '0' ],
  [ '0', '1' ],
    ['1', '0']]
  console.log(ws)
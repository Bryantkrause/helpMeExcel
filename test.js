//  https://stackoverflow.com/questions/51441138/how-to-write-into-a-particular-cell-using-xlsx-npm-package
const XLSX = require('xlsx');

// read from a XLS file
let workbook = XLSX.readFile('butts.xlsx');

// get first sheet
let first_sheet_name = workbook.SheetNames[0];
let worksheet = workbook.Sheets[first_sheet_name];

// read value in D4 
let cell = worksheet['D4'].v;
console.log(cell)

// modify value in D4
worksheet['D4'].v = 'NEW VALUE from NODE';

// modify value if D4 is undefined / does not exists
XLSX.utils.sheet_add_aoa(worksheet, [['NEW VALUE from NODE']], {origin: 'D4'});

// write to new file
// formatting from OLD file will be lost!
XLSX.writeFile(workbook, 'test2.xls');
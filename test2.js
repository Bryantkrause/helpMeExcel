// https://github.com/SheetJS/sheetjs/issues/794

/* load module */
const XLSX = require('xlsx');

/* read workbook */
const workbook = XLSX.readFile('test1.xls');

/* get the first worksheet */
const sheet_name_list = workbook.SheetNames;
const worksheet = workbook.Sheets[sheet_name_list[0]];

/* find cell A1 */
let address = 'A1';
let Sheet1A1 = worksheet[address];

/* create a stub cell if it doesn't exist */
if(!Sheet1A1) Sheet1A1 = worksheet[address] = {t:'z'};

/* print out the value in A1 */
console.log(Sheet1A1.v);

// You need to supply your own object with minimum a type and a value. So:

sheet.A1 = { t: 'n', v: 123 };  // Create A1 as a number
sheet.B1 = { t: 's', v: 'foo' };  // Create B1 as a string
// Edit: this method does not seem to create new rows if the rows do not already exist. To write single cells and create new rows as required I am now using:

XLSX.utils.sheet_add_aoa(sheet, [[123]], {origin: 'A1'});
XLSX.utils.sheet_add_aoa(sheet, [['foo']], {origin: 'B1'});

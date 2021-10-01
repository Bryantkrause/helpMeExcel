const XLSX = require('xlsx');
const table = XLSX.readFile('test1.xlsx');
// Use first sheet
const sheet = table.Sheets[table.SheetNames[0]];
// Option 1: If you have numeric row and column indexes
sheet[XLSX.utils.encode_cell({r: 1 /* 2 */, c: 2 /* C */})] = {t: 's' /* type: string */, v: 'abc123' /* value */};
// Option 2: If you have a cell coordinate like 'C2' or 'D15'
// sheet['A1'] = {t: 's' /* type: string */, v: 'abc123' /* value */};
sheet['A2'] = sheet['A1']
// if the file name below matches file name above this will currently write on the requested cell
// if below is different file name will rewrite with updated data on the requested cell in the new file name
XLSX.writeFile(table, 'test1.xlsx');

//  try this place https://stackoverflow.com/questions/57479988/how-can-i-get-the-count-of-rows-in-my-uploaded-excel
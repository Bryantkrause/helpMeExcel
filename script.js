var xlsx = require("xlsx");
const path = require("path");

const AtHome = true;

if (AtHome) {
	CoolFile = "C:/Users/bryan/Documents/Workdoc.xlsx";
} else {
	CoolFile =
		"J:/2021 Finance/OT/Temp Analysis/2021 Temp Working hours for Budget - Fullerton.xlsx";
}
//var wb = xlsx.readFile("butts.xlsx", {cellDates: true})

var wb = xlsx.readFile(CoolFile, { cellDates: true });

var ws = wb.Sheets["Fullerton 111"];

// var first_sheet_name = wb.Sheets["Fullerton 111"]
// var address_of_cell = {s:{c:1, r:2}, e:{c:7, r:58}};

// /* Get worksheet */
// var worksheet = wb.Sheets[first_sheet_name];
// /* Find desired cell */
// var desired_cell = worksheet[address_of_cell];

/* Get the value */
// var desired_value = (desired_cell ? desired_cell.v : undefined);

let citiStaff111Full = { s: { c: 1, r: 2 }, e: { c: 7, r: 58 } };

// var range = xlsx.utils.decode_range(ws['!ref']);
// var num_rows = range.e.r - range.s.r + 1
// var num_cols = range.e.r - range.s.r + 1

// let findMe = path.dirname(CoolFile)
// let seeMe = path.basename(CoolFile)
// let helpMe = path.extname(CoolFile)
// console.log(findMe)
// console.log(seeMe)
// console.log(helpMe)

// console.log(num_rows)
// console.log(range)
// console.log(num_cols)
console.log(citiStaff111Full);

var range = xlsx.utils.decode_range(wb.Sheets["Fullerton 111"]["!ref"]);
range.s.c = 1; // 0 == XLSX.utils.decode_col("A")
range.e.c = 6; // 6 == XLSX.utils.decode_col("G")
var new_range = xlsx.utils.encode_range(range);
var excelInJSON = xlsx.utils.sheet_to_json(wb.Sheets["Fullerton 111"], {
	defval: "",
	range: { s: { c: 1, r: 1 }, e: { c: 7, r: 58 } },
});

console.log(excelInJSON);

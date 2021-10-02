const xlsx = require("xlsx");
const path = require("path");

const AtHome = true;

data = [];

if (AtHome) {
	CoolFile = "C:/Users/bryan/Documents/Workdoc2.xlsx";
} else {
	CoolFile =
		"J:/2021 Finance/OT/Temp Analysis/2021 Temp Working hours for Budget - Fullerton.xlsx";
}

let wb = xlsx.readFile(CoolFile, { cellDates: true });

let firstSheet = wb.SheetNames[0];

let excelRows = xlsx.utils.sheet_to_row_object_array(wb.Sheets[firstSheet]);

// let parseit = JSON.parse(JSON.stringify(excelRows));
// data.push(parseit);

// let fullerton111 = excelRows.map((excelRows) => {
// 	if (excelRows.Location = "Fullerton") {
// 		return { ...excelRows };
// 	}
// });

let result = excelRows.filter(
	(excelRows) =>
		excelRows.Location === "Fullerton" && excelRows.Department === 105
);
console.log(result);

const xlsx = require("xlsx");
const path = require("path");

const AtHome = false;
// const filePath = "C:/Users/bryan/Documents/testydoc.xlsx";
data = [];

if (AtHome) {
	CoolFile = "C:/Users/bryan/Documents/Workdoc4.xlsx";
	filePath = "C:/Users/bryan/Documents/testydoc.xlsx";
} else {
	CoolFile = "J:/2021 Finance/OT/Temp Analysis/2021 Temp Working hours for Budget - Fullerton.xlsx";
	filePath = "C:/Users/bkrause/Documents/testydoc.xlsx";
}

let wb = xlsx.readFile(CoolFile, { cellDates: true , raw: true}, );

let firstSheet = wb.SheetNames[0];

let excelRows = xlsx.utils.sheet_to_row_object_array(wb.Sheets[firstSheet]);

const locations = ["Fullerton", "Downey", "Cerritos"];

const departments = [105, 110, 111];

const columnNames = [
	"Invoice_date",
	"Employee",
	"Regular_hrs",
	"OT_hrs",
	"Regular_pay",
	"OT_pay",
	"Holiday_pay",
	"Sick_Payment",
	"ACA_Charge",
	"TotalAmt",
	"OT_percentage",
	"Location",
	"Department",
	"Temp_Agency",
];

// separate raw data by location and department into new arrays
let fullerton105 = excelRows
	.map((report) => {
		if (report.Location === "Fullerton" && report.Department === 105) {
			return {
				OT_percentage: report.OT_percentage.z = "0.00",
				// TotalAmt: report.Regular_hrs + report.OT_hrs,
				...report,
			};
		}
	})
	.filter((report) => !!report);

let fullerton111 = excelRows
	.map((report) => {
		if (report.Location === "Fullerton" && report.Department === 111) {
			return {
				// TotalAmt: report.Regular_hrs + report.OT_hrs,
				...report,
			};
		}
	})
	.filter((report) => !!report);

let downey = excelRows
	.map((report) => {
		if (report.Location === "Downey") {
			return {
				...report,
			};
		}
	})
	.filter((report) => !!report);

// let cerritos = excelRows
// 	.map((report) => {
// 		if (report.Location === "Cerritos") {
// 			return {
// 				...report,
// 			};
// 		}
// 	})
// 	.filter((report) => !!report);
const worksheetName1 = "Fullerton105";
console.log("fullerton105", fullerton105, "end fullerton105");
console.log("fullerton111", fullerton111, "end fullerton111");
console.log("downey", downey, "end downey");
// console.log("cerritos", cerritos, "end cerritos");

const exportExcel = (data, columnNames, filePath, worksheetName1) => {
	const workBook = xlsx.utils.book_new();
	const workSheetData = [columnNames, ...data];
	const worksheet = xlsx.utils.aoa_to_sheet(workSheetData);
	xlsx.utils.book_append_sheet(workBook, worksheet, worksheetName1);
	xlsx.writeFile(workBook, path.resolve(filePath));
};

const exportDataToExcel = (
	fullerton105,
	columnNames,
	filePath,
	worksheetName1
) => {
	const data = fullerton105.map((report) => {
		return [
			report.Invoice_date,
			report.Employee,
			report.Regular_hrs,
			report.OT_hrs,
			report.Regular_pay,
			report.OT_pay,
			report.Holiday_pay,
			report.Sick_Payment,
			report.ACA_Charge,
			report.TotalAmt,
			report.OT_percentage.z ,
			report.Location,
			report.Department,
			report.Temp_Agency,
		];
	});
	exportExcel(data, columnNames, filePath, worksheetName1);
};

exportDataToExcel(fullerton105, columnNames, filePath, worksheetName1);

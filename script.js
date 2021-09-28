var xlsx = require("xlsx")

var wb = xlsx.readFile("butts.xlsx", {cellDates: true})

var ws = wb.Sheets["Sheet1"]

console.log(ws)
const xlsx = require("xlsx")

const wb = xlsx.readFile("./../input/references/TIME.xlsx", { cellFormula: true })
const timeSheet = wb.Sheets[wb.SheetNames[0]]

const rows = xlsx.utils.sheet_to_json(timeSheet)
const percentage = Object.values(rows[4])[1]

console.log(percentage)

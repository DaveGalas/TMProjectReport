const { resourcesData } = require("./scripts/data/resourcesData")
const { projectsData } = require("./scripts/data/projectsData")
const { noPlanData } = require("./scripts/data/noPlanData")
const { files } = require("./scripts/excel/filePaths")
const { extractData } = require("./scripts/excel/extractData")
const { writeBook } = require("./scripts/excel/writeBook")
const { timeExtract, toUnix } = require("./scripts/utils")

const [TTRdata, OTBMdata, PPMCdata, RATESdata, TIMEdata] = extractData(files)
const { reportDate, today, eoPrevMonth } = timeExtract(TIMEdata)
const percentage = 0.85

const projects = projectsData(OTBMdata, TTRdata, eoPrevMonth)
const resources = resourcesData(TTRdata, PPMCdata, RATESdata, projects, percentage)
const noPlan = noPlanData(OTBMdata, TTRdata, PPMCdata, RATESdata, eoPrevMonth)

writeBook(projects, resources, noPlan, new Date(today), percentage, toUnix(reportDate))

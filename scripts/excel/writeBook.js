const excel = require("xlsx")
const { toUnix } = require("../utils")
const { readFileSync, writeFileSync } = require("fs")

exports.writeBook = ( projects, resources, noPlan, today, percentage, reportDate ) => {
    console.log("...Writing Excel file");

    const projectsSheetData = projects
    const resourcesSheetData = [...resources, ...noPlan]

    const filePath = `./output/PG T&M Projects Report ${
        today.getMonth() + 1
    }-${today.getDate()}-${today.getFullYear()}.xlsx`

    const projectsSheet = excel.utils.json_to_sheet(projectsSheetData)
    const resourcesSheet = excel.utils.json_to_sheet(resourcesSheetData)

    const wb = excel.utils.book_new()

    excel.utils.book_append_sheet(wb, projectsSheet, "T&M Projects List")
    excel.utils.book_append_sheet(wb, resourcesSheet, "Resources List")

    excel.writeFile( wb, filePath )
    writeFileSync("./output/dynamic/projects.json", JSON.stringify(projectsSheetData))
    writeFileSync("./output/dynamic/resources.json", JSON.stringify(resourcesSheetData))
    
    const onTrackCount = resourcesSheetData.filter((e) => e["Audit"] == "On Track").length
    const undertrackingCount = resourcesSheetData.filter((e) => e["Audit"] == "Undertracking").length
    const noTrackingCount = resourcesSheetData.filter((e) => e["Audit"] == "No Tracking").length
    const noPlanCount = resourcesSheetData.filter((e) => e["Audit"] == "No Plan").length

    const options = { year: "numeric", month: "long", day: "numeric" }

    const log = JSON.parse(readFileSync("./misc/log.json", "utf-8"))

    log.push({
        createdAt: today,
        reportDate,
        percentage,
        onTrackCount,
        undertrackingCount,
        noTrackingCount,
        noPlanCount,
    })

    writeFileSync( "./misc/log.json", JSON.stringify( log ) )

    console.log(`\nUtilization rate: ${percentage.toLocaleString("en-GB", { style: "percent" })}`,
        `\n\nGeneration Date: ${today.toLocaleDateString("en-US", options)}`,
        `\nReport Date: ${reportDate.toLocaleDateString("en-US", options)}`,
        `\n\nOn Track Resources: ${onTrackCount}`,
        `\nUndertracking Resources: ${undertrackingCount}`,
        `\nNo Tracking Resources: ${noTrackingCount}`,
        `\nNo Plan Resources: ${noPlanCount}`,
        `\n\nReport generated - PG T&M Projects Report ${
            today.getMonth() + 1
        }-${today.getDate()}-${today.getFullYear()}.xlsx`
    )
}

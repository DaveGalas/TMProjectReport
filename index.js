const excel = require("xlsx")
const fs = require("fs")
const utils = require("./misc/utils")

const today = new Date()
const eoPrevMonth = new Date(today.getFullYear(), today.getMonth(), 1)

// Filenames
const TTRFileName = "./input/TTR.xlsx"
const OTBMFileName = "./input/OTBM.xlsx"
const PPMCFileName = "./input/PPMC.xlsx"
const RATESFileName = "./input/references/RATES.xlsx"
const TIMEFileName = "./input/references/TIME.xlsx"
//

console.log("... Reading Excel files.")

// Data extraction function
const extractData = (arg) => {
    const args = [...arg]
    const fileData = []

    args.forEach((e) => {
        const temp = []
        const tempFile = excel.readFile(e)
        const tempSheets = tempFile.SheetNames

        tempSheets.forEach((f) => {
            const sheetRows = excel.utils.sheet_to_json(tempFile.Sheets[f])

            sheetRows.forEach((g) => {
                temp.push(g)
            })
        })

        fileData.push(temp)
    })

    return fileData
}

const [TTRdata, OTBMdata, PPMCdata, RATESdata, TIMEdata] = extractData([
    TTRFileName,
    OTBMFileName,
    PPMCFileName,
    RATESFileName,
    TIMEFileName,
])
//

// Percentage calculation
const percentage = Object.values(TIMEdata[4])[1]

console.log(`Utilization rate: ${percentage.toLocaleString("en-GB", { style: "percent" })}`)
//

// No plan empty array
let noPlan
//

// Projects list data
const projectsData = () => {
    // Filter project names with 'SCPT'
    const OTBMdataRef = OTBMdata.filter((e) => !e["Project record entry - Project Name"].includes("SCPT"))
    const TTRdataRef = TTRdata
    const PPMCdataRef = PPMCdata.filter((e) => !e["Project Name"]?.includes("SCPT"))
    //

    // Filter OTBM projects with BM CODE TM and month prev month and beyond
    const otbmTM = OTBMdataRef.filter((e) => e["BM Code"].includes("TM"))
    const otbmTMrecent = otbmTM.filter(
        (e) =>
            Number(e["BM Code"].slice(-4)) >=
            Number(`${String(eoPrevMonth.getFullYear()).slice(-2)}${utils.toTwoDigits(eoPrevMonth.getMonth())}`)
    )
    //

    // Filter T&M TTR rows.
    const ttrTMSA = TTRdataRef.filter(
        (e) => e["Engagement Type"] == "Projects - T&M" || e["Engagement Type"] == "SA - TM"
    )
    const ttrUMSA = TTRdataRef.filter(
        (e) => e["Engagement Type"] == "Unmapped" || e["Engagement Type"] == "Service Agreement"
    )
    const ttrUMSATM = ttrUMSA.filter((e) => {
        return otbmTMrecent.map((f) => f["HP Project #"]).includes(e["Project Name"])
    })
    const ttrTM = [...ttrUMSATM, ...ttrTMSA]
    //

    // Catch no-plan TTR rows from PPMC data
    const PGIDregex = /PG\d{2}P\d{5}/

    const ppmcWithProjNum = PPMCdataRef.map((e) => {
        const projectNumber = e["Project Name"]?.match(PGIDregex)
        return {
            ...e,
            ProjectNumber: projectNumber ? projectNumber[0] : "",
        }
    })

    noPlan = ttrTM.filter((e) => {
        return !ppmcWithProjNum.some((f) => {
            return e["Project Name"] === f["ProjectNumber"] && e["Employee ID"] === f["Employee ID"]
        })
    })
    //

    // Filter duplcate TTRTMs
    const uniqueTtrTM = ttrTM.filter((obj, index, arr) => {
        return arr.map((mapObj) => mapObj["Project Name"]).indexOf(obj["Project Name"]) === index
    })
    //

    // Map additional columns, remove empty Project Names
    const mainProjectsData = uniqueTtrTM
        .map((e) => {
            const PGID = e["Project Name"]
            const proj = otbmTMrecent.find((e) => e["HP Project #"] === PGID)
            return {
                ...e,
                // PName: proj ? proj["Project record entry - Project Name"] : "No Name",
                PName: proj ? proj["Project record entry - Project Name"] : "No Name",
                "Project Classification": proj ? proj["Project Classification"] : "No Classification",
                "COMPASS FMO WBS": proj ? proj["COMPASS FMO WBS"] : "No WBS",
                "DXC Project Manager": proj ? proj["DXC Project Manager"] : "No Project Manager",
                "Project Stage": proj ? proj["Project Stage"] : "No Stage",
            }
        })
        .filter((e) => {
            const PName = e["PName"]
            return PName != "No Name" && (!PName || !PName.includes("SCPT"))
        })

    return mainProjectsData
}
//

// Resources list data
const resourcesData = () => {
    // Files
    const projectsDataRef = projectsData().filter((obj) => obj.hasOwnProperty("Project Name"))
    const noPlanDataRef = noPlan
    const OTBMdataRef = OTBMdata.filter((e) => !e["Project record entry - Project Name"].includes("SCPT"))
    const TTRdataRef = TTRdata
    const PPMCdataRef = PPMCdata.filter((e) => !e["Project Name"]?.includes("SCPT"))
    const RATESdataRef = RATESdata
    //

    const PGIDregex = /PG\d{2}P\d{5}/

    //  Filter out No Resource Allocated PPMC rows
    const ppmcWithResources = PPMCdataRef.filter((e) => {
        const projectName = e["Project Name"]
        const regexedName = projectName?.match(PGIDregex)?.[0] ?? null
        return regexedName !== null
    }).filter((e) => {
        return e["Allocation status"] != "No Resource Allocated"
    })
    //

    // Extract PGID from Project Name
    const ppmcWithPGID = ppmcWithResources.map((e) => {
        const projectName = e["Project Name"]
        const regexedName = projectName?.match(PGIDregex)?.[0] ?? null
        return {
            ...e,
            PGID: regexedName,
        }
    })
    //

    //
    const projectNumbers = projectsDataRef.map((e) => e["Project Name"])
    const mapProjectstoPPMC = ppmcWithPGID.filter((e) => projectNumbers.includes(e["PGID"]))

    const uniqueProjects = []
    const seenProjects = new Set()

    mapProjectstoPPMC.forEach((e) => {
        const key = e["PGID"] + "-" + e["Employee ID"]
        if (!seenProjects.has(key)) {
            uniqueProjects.push(e)
            seenProjects.add(key)
        }
    })
    //

    // ** uniqueProjects is the final list of all resources **

    // The following are additional columns

    /// FTE calculation
    const billedFTEsummed = uniqueProjects.map((e) => {
        const { PGID: PGID, "Employee ID": employeeId } = e
        const totalBilledFTE = TTRdataRef.reduce((acc, curr) => {
            if (curr["Project Name"] === PGID && curr["Employee ID"] === employeeId) {
                return acc + curr["Billed FTE"]
            }
            return acc
        }, 0)
        return { ...e, "Sum of Actual FTE": totalBilledFTE }
    })

    const ppmcFTEsummed = billedFTEsummed.map((e) => {
        const { PGID: PGID, "Employee ID": employeeId } = e
        const totalPPMCFTE = ppmcWithPGID.reduce((acc, curr) => {
            if (curr["PGID"] === PGID && curr["Employee ID"] === employeeId) {
                return acc + curr["PPMC FTE"]
            }
            return acc
        }, 0)
        return { ...e, "Sum of PPMC FTE": totalPPMCFTE }
    })

    const calculatedFTEstats = ppmcFTEsummed.map((e) => {
        const totalActualFTE = e["Sum of Actual FTE"]
        const totalPPMCFTE = e["Sum of PPMC FTE"]
        const FTEvariance = totalPPMCFTE - totalActualFTE
        const utilizationRate = totalActualFTE / totalPPMCFTE
        return { ...e, "FTE Variance": FTEvariance, "Utilization Rate": utilizationRate }
    })
    ///

    /// Hours calculation
    const actualHours = calculatedFTEstats.map((e) => {
        const { PGID: projectNumber, "Employee ID": employeeId } = e
        const totalHours = TTRdataRef.reduce((acc, curr) => {
            if (curr["Project Name"] === projectNumber && curr["Employee ID"] === employeeId) {
                return acc + curr["Hours"]
            }
            return acc
        }, 0)
        return { ...e, "Actual Hours": totalHours }
    })

    const hoursCalculation = actualHours.map((e) => {
        let standardHours
        e["Country"].toLowerCase() == "India".toLowerCase() ? (standardHours = 207) : (standardHours = 184)
        const plannedHours = standardHours * e["Sum of PPMC FTE"]
        const hoursVariance = plannedHours - e["Actual Hours"]
        return { ...e, "Planned Hours": plannedHours, "Hours Variance": hoursVariance }
    })
    ///

    /// Revenue calculation
    const revenueCalculation = hoursCalculation.map((e) => {
        const key = e["DXC Job Level"] + e["Country"]
        const rateRow = RATESdataRef.find((e) => e["CSJL"] == key)
        const rate = rateRow ? rateRow["Rate"] : null
        const actualRevenue = rate * e["Actual Hours"]
        const plannedRevenue = rate * e["Planned Hours"]
        const revenueVariance = plannedRevenue - actualRevenue
        return {
            ...e,
            "Hourly Revenue Rates (USD)": rate,
            "Actual Revenue": actualRevenue,
            "Planned Revenue": plannedRevenue,
            "Revenue Variance": revenueVariance,
        }
    })
    ///

    /// Audit calculation
    const audited = revenueCalculation.map((e) => {
        const uRate = e["Utilization Rate"]
        let audit
        if (uRate == 0) audit = "No Tracking"
        if (uRate < percentage && uRate > 0) audit = "Undertracking"
        if (uRate >= percentage) audit = "On Track"
        return { ...e, Audit: audit }
    })

    const mainResourcesData = audited
    ///

    // No Plan data
    const noPlanData = noPlanDataRef.map((e) => {
        const matchingOTBMProject = OTBMdataRef.find((p) => p["HP Project #"] === e["Project Name"])

        if (matchingOTBMProject) {
            e["Project Name2"] = matchingOTBMProject["Project record entry - Project Name"]
        }

        // -- FTE calculation
        const billedFTEsum = e["Billed FTE"]

        const totalPPMCFTE = ppmcWithPGID.reduce((acc, curr) => {
            if (curr["ProjectNumber"] === e["Employee ID"] && curr["Employee ID"] === e["Employee ID"]) {
                return acc + curr["PPMC FTE"]
            }
            return acc
        }, 0)

        const fteVar = totalPPMCFTE - billedFTEsum

        const uRate = totalPPMCFTE !== 0 ? billedFTEsum / totalPPMCFTE : 0
        //

        // -- Hours calculation
        const actualHours = e["Hours"]

        let standardHours

        e["Country"].toLowerCase() == "India".toLowerCase() ? (standardHours = 207) : (standardHours = 184)

        const plannedHours = standardHours * totalPPMCFTE || 0
        const hoursVar = plannedHours - actualHours
        //

        // -- Revenue calculation
        const key = e["DXC Job Level"] + e["Country"]
        const rateRow = RATESdataRef.find((e) => e["CSJL"] == key)
        const rate = rateRow ? rateRow["Rate"] : null
        const actualRevenue = rate * actualHours
        const plannedRevenue = rate * plannedHours
        const revenueVariance = plannedRevenue - actualRevenue
        //
        return {
            Segment: e["Segment"],
            ADM: e["ADM"],
            "Project ID": e["Project Name"],
            "Project Name": e["Project Name2"],
            "Manager Name": e["Manager Name"],
            "Employee ID": e["Employee ID"],
            Employee: e["Employee"],
            "PG Job Level": e["Position Role"],
            "DXC Job Level": e["DXC Job Level"],
            Country: e["Country"],
            "Sum of Actual FTE": billedFTEsum,
            "Sum of PPMC FTE": totalPPMCFTE || 0,
            "FTE Variance": fteVar,
            "Utilization Rate": uRate,
            "Actual Hours": actualHours,
            "Planned Hours": plannedHours,
            "Hours Variance": hoursVar,
            "Hourly Revenue Rates (USD)": rate,
            "Actual Revenue": actualRevenue,
            "Planned Revenue": plannedRevenue,
            "Revenue Variance": revenueVariance,
            Audit: "No Plan",
        }
    })
    //

    return [
        ...mainResourcesData.map((e) => {
            return {
                Segment: e["Segment"],
                ADM: e["ADM"],
                "Project Name": e["Project Name"],
                "Project ID": e["PGID"],
                "Manager Name": e["Manager Name"],
                "Employee ID": Number(e["Employee ID"]),
                Employee: e["Employee"],
                "PG Job Level": e["Job Level"],
                "DXC Job Level": e["DXC Job Level"],
                Country: e["Country"],
                "Sum of Actual FTE": Number(parseFloat(e["Sum of Actual FTE"]).toFixed(2)),
                "Sum of PPMC FTE": Number(parseFloat(e["Sum of PPMC FTE"]).toFixed(2)),
                "FTE Variance": Number(parseFloat(e["FTE Variance"]).toFixed(2)),
                "Utilization Rate": e["Utilization Rate"],
                "Actual Hours": e["Actual Hours"],
                "Planned Hours": e["Planned Hours"],
                "Hours Variance": e["Hours Variance"],
                "Hourly Revenue Rates (USD)": e["Hourly Revenue Rates (USD)"],
                "Actual Revenue": e["Actual Revenue"],
                "Planned Revenue": e["Planned Revenue"],
                "Revenue Variance": e["Revenue Variance"],
                Audit: e["Audit"],
            }
        }),
        ...noPlanData.filter((e) => e["Project Name"]),
    ]
}
//

// Write Excel File
const writeExcel = () => {
    const projectsSheet = excel.utils.json_to_sheet(
        projectsData().map((e) => {
            return {
                Segment: e["Segment"],
                "PG ID": e["Project Name"],
                "Project Name": e["PName"],
                ADM: e["ADM"],
                "DXC Project Manager": e["DXC Project Manager"],
                "Project Stage": e["Project Stage"],
                "Project Classification": e["Project Classification"],
                "COMPASS FMO WBS": e["COMPASS FMO WBS"],
            }
        })
    )
    const resourcesSheet = excel.utils.json_to_sheet(resourcesData())

    const filePath = `./output/PG T&M Projects Report ${
        today.getMonth() + 1
    }-${today.getDate()}-${today.getFullYear()}.xlsx`

    const wb = excel.utils.book_new()
    excel.utils.book_append_sheet(wb, projectsSheet, "T&M Projects List")
    excel.utils.book_append_sheet(wb, resourcesSheet, "Resources List")

    excel.writeFile(wb, filePath)

    const onTrackCount = resourcesData().filter((e) => e["Audit"] == "On Track").length
    const undertrackingCount = resourcesData().filter((e) => e["Audit"] == "Undertracking").length
    const noTrackingCount = resourcesData().filter((e) => e["Audit"] == "No Tracking").length
    const noPlanCount = resourcesData().filter((e) => e["Audit"] == "No Plan").length

    const log = JSON.parse(fs.readFileSync("./misc/log.json", "utf-8"))
    log.push({ createdAt: today, percentage, onTrackCount, undertrackingCount, noTrackingCount, noPlanCount })
    fs.writeFileSync("./misc/log.json", JSON.stringify(log))

    console.log(
        `\nOn Track Resources: ${onTrackCount}`,
        `\nUndertracking Resources: ${undertrackingCount}`,
        `\nNo Tracking Resources: ${noTrackingCount}`,
        `\nNo Plan Resources: ${noPlanCount}`,
        `\n\nReport generated - PG T&M Projects Report ${
            today.getMonth() + 1
        }-${today.getDate()}-${today.getFullYear()}.xlsx`
    )
}
//
writeExcel()

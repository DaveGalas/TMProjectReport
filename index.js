const excel = require("xlsx")
const utils = require("./misc/utils")

const today = new Date()
const eoPrevMonth = new Date(today.getFullYear(), today.getMonth(), 1)

const percentage = process.argv[2] * 1
if (!percentage) {
  throw new Error("Percentage needed to calculate audit.")
} else if (!Number(percentage)) {
}

const TTRFileName = "./input/TTR.xlsx"
const OTBMFileName = "./input/OTBM.xlsx"
const PPMCFileName = "./input/PPMC.xlsx"
const RATESFileName = "./input/references/RATES.xlsx"

// Data extraction
function extractData() {
  let fileData = []
  const args = [...arguments]
  args.forEach((e) => {
    const temp = []
    const tempFile = excel.readFile(e)
    const tempSheets = tempFile.SheetNames

    tempSheets.forEach((tempSheet) => {
      const sheetRows = excel.utils.sheet_to_json(tempFile.Sheets[tempSheet])
      sheetRows.forEach((sheetRow) => {
        temp.push(sheetRow)
      })
    })

    fileData.push(temp)
  })
  return fileData
}
const [TTRdata, OTBMdata, PPMCdata, RATESdata] = extractData(TTRFileName, OTBMFileName, PPMCFileName, RATESFileName)

// Projects list data
let noPlan = []
const projectsData = () => {
  // copy OTBM and TTR data
  const OTBMdataRef = OTBMdata
  const TTRdataRef = TTRdata
  const PPMCdataRef = PPMCdata
  // filter OTBM rows with BM Code containing TM
  const otbmF1 = OTBMdataRef.filter((e) => e["BM Code"].includes("TM"))
  // filter otbmF1 rows with BM month greater than or equal to previous month
  const otbmF2 = otbmF1.filter(
    (e) =>
      Number(e["BM Code"].slice(-4)) >=
      Number(`${String(eoPrevMonth.getFullYear()).slice(-2)}${utils.toTwoDigits(eoPrevMonth.getMonth())}`)
  )
  // filter TTR rows with Engagement Types Projects - T&M and SA - TM
  const ttrTMSA = TTRdataRef.filter(
    (e) => e["Engagement Type"] == "Projects - T&M" || e["Engagement Type"] == "SA - TM"
  )
  // filter TTR rows with Engagement Types Unmapped and Service Agreement
  const ttrUMSA = TTRdataRef.filter(
    (e) => e["Engagement Type"] == "Unmapped" || e["Engagement Type"] == "Service Agreement"
  )
  // filter ttrUMSA with PGID included in otbmF2
  const ttrUMSATM = ttrUMSA.filter((e) => {
    return otbmF2.map((f) => f["HP Project #"]).includes(e["Project Name"])
  })
  // combine Projects - T&M, SA - TM projects, and TM Unmapped, Service Agreement
  const ttrTM = [...ttrUMSATM, ...ttrTMSA]

  //// Catch no-plan TTR rows in PPMC data
  const regex = /PG\d\dP\d\d\d\d\d/
  const PPMCwithProjNum = PPMCdataRef.map((e) => {
    const projectNumber = e["Project Name"]?.match(regex)
    return {
      ...e,
      ProjectNumber: projectNumber ? projectNumber[0] : "",
    }
  })

  noPlan = ttrTM.filter((e) => {
    return !PPMCwithProjNum.some((f) => {
      return e["Project Name"] === f["ProjectNumber"] && e["Employee ID"] === f["Employee ID"]
    })
  })
  ////

  // filter duplicates from ttrTM
  const uniqueTtrTM = ttrTM.filter((obj, index, arr) => {
    return arr.map((mapObj) => mapObj["Project Name"]).indexOf(obj["Project Name"]) === index
  })
  // filter projects with SCPT in their name or if the project doesn't have a name
  const mainProjectsData = uniqueTtrTM
    .map((e) => {
      const PGID = e["Project Name"]
      const proj = otbmF2.find((e) => e["HP Project #"] === PGID)
      return {
        ...e,
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
// Resources list data
const resourcesData = () => {
  const projectsDataRef = projectsData()
  const noPlanDataRef = noPlan
  const OTBMdataRef = OTBMdata
  const PPMCdataRef = PPMCdata
  const TTRdataRef = TTRdata
  const RATESdataRef = RATESdata
  /// Main resources data
  const regex = /PG\d\dP\d\d\d\d\d/
  const filteredProjects = PPMCdataRef.filter((e) => {
    const projectName = e["Project Name"]
    const regexedName = projectName?.match(regex)?.[0] ?? null
    return regexedName !== null
  })
  const projectsWithRegexedName = filteredProjects.map((e) => {
    const projectName = e["Project Name"]
    const regexedName = projectName?.match(regex)?.[0] ?? null
    return {
      ...e,
      regexedName,
    }
  })
  const projectNumbers = projectsDataRef.map((e) => e["Project Name"])
  const filteredProjects2 = projectsWithRegexedName.filter((e) => projectNumbers.includes(e["regexedName"]))
  const uniqueProjects = []
  const seenProjects = new Set()
  filteredProjects2.forEach((e) => {
    const key = e["regexedName"] + "-" + e["Employee ID"]
    if (!seenProjects.has(key)) {
      uniqueProjects.push(e)
      seenProjects.add(key)
    }
  })
  // uniqueProjects is the final list of all resources
  const PPMCwithProjNum = PPMCdata.map((e) => {
    const projectNumber = e["Project Name"]?.match(regex)
    return {
      ...e,
      ProjectNumber: projectNumber ? projectNumber[0] : "",
    }
  })
  /// FTE calculation
  const billedFTEsummed = uniqueProjects.map((e) => {
    const { regexedName: projectNumber, "Employee ID": employeeId } = e
    const totalFTE = TTRdataRef.reduce((acc, curr) => {
      if (curr["Project Name"] === projectNumber && curr["Employee ID"] === employeeId) {
        return acc + curr["Billed FTE"]
      }
      return acc
    }, 0)
    return { ...e, "Sum of Actual FTE": totalFTE }
  })
  const PPMCFTEsummed = billedFTEsummed.map((e) => {
    const { regexedName: projectNumber, "Employee ID": employeeId } = e
    const totalPPMCFTE = PPMCwithProjNum.reduce((acc, curr) => {
      if (curr["ProjectNumber"] === projectNumber && curr["Employee ID"] === employeeId) {
        return acc + curr["PPMC FTE"]
      }
      return acc
    }, 0)
    return { ...e, "Sum of PPMC FTE": totalPPMCFTE }
  })
  const calculatedFTEstats = PPMCFTEsummed.map((e) => {
    const totalActualFTE = e["Sum of Actual FTE"]
    const totalPPMCFTE = e["Sum of PPMC FTE"]
    const FTEvariance = totalPPMCFTE - totalActualFTE
    const utilizationRate = totalActualFTE / totalPPMCFTE
    return { ...e, "FTE Variance": FTEvariance, "Utilization Rate": utilizationRate }
  })
  /// Hours calculation
  const actualHours = calculatedFTEstats.map((e) => {
    const { regexedName: projectNumber, "Employee ID": employeeId } = e
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
  /// Revenue Calculation
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
  /// Audit Calculation
  const audited = revenueCalculation.map((e) => {
    const uRate = e["Utilization Rate"]
    let audit
    if (uRate == 0) audit = "No Tracking"
    if (uRate < percentage && uRate > 0) audit = "Undertracking"
    if (uRate >= percentage) audit = "On Track"
    return { ...e, Audit: audit }
  })
  const mainResourcesData = audited

  /// No Plan Data
  const noPlanData = noPlanDataRef.map((e) => {
    const matchingOTBMProject = OTBMdataRef.find((p) => p["HP Project #"] === e["Project Name"])
    if (matchingOTBMProject) {
      e["Project Name2"] = matchingOTBMProject["Project record entry - Project Name"]
    }
    const billedFTEsum = e["Billed FTE"]
    const totalPPMCFTE = PPMCwithProjNum.reduce((acc, curr) => {
      if (curr["ProjectNumber"] === e["Employee ID"] && curr["Employee ID"] === e["Employee ID"]) {
        return acc + curr["PPMC FTE"]
      }
      return acc
    }, 0)
    const fteVar = totalPPMCFTE - billedFTEsum
    const uRate = totalPPMCFTE !== 0 ? billedFTEsum / totalPPMCFTE : 0
    const actualHours = e["Hours"]
    let standardHours
    e["Country"].toLowerCase() == "India".toLowerCase() ? (standardHours = 207) : (standardHours = 184)
    const plannedHours = standardHours * totalPPMCFTE || 0
    const hoursVar = plannedHours - actualHours
    const key = e["DXC Job Level"] + e["Country"]
    const rateRow = RATESdataRef.find((e) => e["CSJL"] == key)
    const rate = rateRow ? rateRow["Rate"] : null
    const actualRevenue = rate * actualHours
    const plannedRevenue = rate * plannedHours
    const revenueVariance = plannedRevenue - actualRevenue
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

  return [
    ...mainResourcesData.map((e) => {
      return {
        Segment: e["Segment"],
        ADM: e["ADM"],
        "Project Name": e["Project Name"],
        "Project ID": e["regexedName"],
        "Manager Name": e["Manager Name"],
        "Employee ID": e["Employee ID"] * 1 ? e["Employee ID"] * 1 : e["Employee ID"],
        Employee: e["Employee"],
        "PG Job Level": e["Job Level"],
        "DXC Job Level": e["DXC Job Level"],
        Country: e["Country"],
        "Sum of Actual FTE": e["Sum of Actual FTE"],
        "Sum of PPMC FTE": e["Sum of PPMC FTE"],
        "FTE Variance": e["FTE Variance"],
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
    ...noPlanData,
  ]
}

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

  // excel.utils.book_append_sheet(wb, descriptionSheet, "Report")

  excel.writeFile(wb, filePath)
  console.log(
    `\nReport generated - PG T&M Projects Report ${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}.xlsx`
  )
}

writeExcel()

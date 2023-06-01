const { toTwoDigits } = require("../utils")

exports.noPlanData = (OTBMdata, TTRdata, PPMCdata, RATESdata, eoPrevMonth) => {
    console.log("...Extracting resources without a plan")

    const OTBMdataRef = OTBMdata.filter((e) => !e["Project record entry - Project Name"].includes("SCPT"))
    const TTRdataRef = TTRdata
    const PPMCdataRef = PPMCdata.filter((e) => !e["Project Name"]?.includes("SCPT"))
    const RATESdataRef = RATESdata

    const PGIDregex = /PG\d{2}P\d{5}/

    const otbmPrime = OTBMdataRef.filter((e) => {
        const { "BM Code": bmCode } = e

        const includesTM = bmCode.includes("TM")
        const isGTEeoPrev =
            Number(bmCode.slice(-4)) >=
            Number(`${String(eoPrevMonth.getFullYear()).slice(-2)}${toTwoDigits(eoPrevMonth.getMonth())}`)

        return includesTM && isGTEeoPrev
    })

    const ttrPrime = TTRdataRef.filter((e) => {
        const isSATM = e["Engagement Type"] == "Projects - T&M" || e["Engagement Type"] == "SA - TM"
        const isSAUM =
            (e["Engagement Type"] == "Unmapped" || e["Engagement Type"] == "Service Agreement") &&
            otbmPrime.some((f) => f["HP Project #"] === e["Project Name"])

        return isSATM || isSAUM
    })

    const ppmcWithProjNum = []
    for (let i = 0; i < PPMCdataRef.length; i++) {
        const e = PPMCdataRef[i]
        const projectNumber = e["Project Name"]?.match(PGIDregex)
        ppmcWithProjNum.push({
            ...e,
            ProjectNumber: projectNumber ? projectNumber[0] : "",
        })
    }

    const ppmcPrime = PPMCdataRef.map((e) => {
        const { "Project Name": projectName } = e
        const regexedName = projectName?.match(PGIDregex)?.[0] ?? null

        return {
            ...e,
            PGID: regexedName,
        }
    }).filter((e) => {
        const { PGID, "Allocation status": alloc } = e
        return PGID && alloc !== "No Resource Allocated"
    })

    const noPlanData = ttrPrime.filter((e) => {
        return !ppmcWithProjNum.some(
            (f) => e["Project Name"] === f["ProjectNumber"] && e["Employee ID"] === f["Employee ID"]
        )
    })

    const noPlanExtract = []
    for (let i = 0; i < noPlanData.length; i++) {
        const e = noPlanData[i]
        const {
            Country,
            Segment,
            Employee,
            Hours: actualHours,
            "DXC Job Level": dxcJL,
            "Billed FTE": BFTE,
            "Project Name": PGID,
            "Employee ID": EID,
            "Manager Name": manager,
            "Position Role": role,
        } = e

        let totalPPMCFTE = 0
        for (let j = 0; j < ppmcPrime.length; j++) {
            const curr = ppmcPrime[j]
            if (curr["ProjectNumber"] === EID && curr["Employee ID"] === EID) {
                totalPPMCFTE += curr["PPMC FTE"]
            }
        }

        const fteVar = totalPPMCFTE - BFTE
        const uRate = totalPPMCFTE !== 0 ? BFTE / totalPPMCFTE : 0
        const standardHours = Country.toLowerCase() == "india" ? 207 : 184
        const plannedHours = standardHours * totalPPMCFTE || 0
        const hoursVar = plannedHours - actualHours
        const key = dxcJL + Country
        const rateRow = RATESdataRef.find((f) => f["CSJL"] == key)
        const rate = rateRow ? rateRow["Rate"] : null
        const actualRevenue = rate * actualHours
        const plannedRevenue = rate * plannedHours
        const revenueVariance = plannedRevenue - actualRevenue

        const matchingOTBMProject = otbmPrime.find((p) => p["HP Project #"] === PGID)
        if (matchingOTBMProject) {
            e["Project Name2"] = matchingOTBMProject["Project record entry - Project Name"]
        }
        if (!e["Project Name2"]) continue

        if (!e["ADM"]) {
            e["ADM"] =
                e["Segment"].toLowerCase() == "voice"
                    ? "Arevalo, Milton"
                    : e["Segment"].toLowerCase() == "infra"
                    ? "Manalac, Von Xavier"
                    : null
        }

        noPlanExtract.push({
            Segment: Segment,
            ADM: e["ADM"],
            "Project ID": PGID,
            "Project Name": e["Project Name2"],
            "Manager Name": manager,
            "Employee ID": EID,
            Employee: Employee,
            "PG Job Level": role,
            "DXC Job Level": dxcJL,
            Country: Country,
            "Sum of Actual FTE": BFTE,
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
        })
    }

    return noPlanExtract
}

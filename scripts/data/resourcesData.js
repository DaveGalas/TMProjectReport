const { fte, hours, revenue, audit } = require("../calculations")

exports.resourcesData = ( TTRdata, PPMCdata, RATESdata, projectsData, percentage ) => {
    console.log("...Extracting tracking resources")

    const TTRdataRef = TTRdata
    const PPMCdataRef = PPMCdata.filter((e) => !e["Project Name"]?.includes("SCPT"))
    const RATESdataRef = RATESdata

    const PGIDregex = /PG\d{2}P\d{5}/

    const ppmcPrime = PPMCdataRef.filter((e) => {
        const projectName = e["Project Name"]
        const regexedName = projectName?.match(PGIDregex)?.[0] ?? null
        return regexedName !== null && e["Allocation status"] != "No Resource Allocated"
    }).map((e) => {
        const projectName = e["Project Name"]
        const regexedName = projectName?.match(PGIDregex)?.[0] ?? null
        return {
            ...e,
            PGID: regexedName,
        }
    })

    const uniqueProjects = ppmcPrime.filter((e, i, a) => {
        const projectNumbers = projectsData.map((e) => e["PGID"])
        const key = e["PGID"] + "-" + e["Employee ID"]

        const isKeyUnique = i === a.findIndex((p) => p["PGID"] + "-" + p["Employee ID"] === key)
        const isInProjectsData = projectNumbers.includes(e["PGID"])

        return isKeyUnique && isInProjectsData
    })

    const fteStats = fte(uniqueProjects, [TTRdataRef, ppmcPrime])
    const hoursStats = hours(fteStats, [TTRdataRef])
    const revenueStats = revenue(hoursStats, [RATESdataRef])
    const audited = audit(revenueStats, percentage)

    const columnedResources = audited.map((e) => {
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
    })

    return columnedResources
}

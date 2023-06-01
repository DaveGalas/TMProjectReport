const { toTwoDigits } = require("../utils")

exports.projectsData = ( OTBMdata, TTRdata, prevMonth ) => {
    console.log("...Extracting T&M projects");

    const OTBMdataRef = OTBMdata.filter((e) => !e["Project record entry - Project Name"].includes("SCPT"))
    const TTRdataRef = TTRdata
    const eoPrevMonth = new Date(prevMonth)

    const otbmPrime = OTBMdataRef.filter((e) => {
        const includesTM = e["BM Code"].includes("TM")
        const isGTEeoPrev =
            Number(e["BM Code"].slice(-4)) >=
            Number(`${String(eoPrevMonth.getFullYear()).slice(-2)}${toTwoDigits(eoPrevMonth.getMonth())}`)

        return includesTM && isGTEeoPrev
    })

    const ttrPrime = TTRdataRef.filter((e, i, a) => {
        const isSATM = ["Projects - T&M", "SA - TM"].includes(e["Engagement Type"])
        const isSAUM =
            ["Unmapped", "Service Agreement"].includes(e["Engagement Type"]) &&
            otbmPrime.map((f) => f["HP Project #"]).includes(e["Project Name"])
        const isUnique = a.map((f) => f["Project Name"]).indexOf(e["Project Name"]) === i

        return (isSATM || isSAUM) && isUnique
    })

    const projects = ttrPrime
        .map((e) => {
            const proj = otbmPrime.find((f) => f["HP Project #"] === e["Project Name"])
            const projectName = proj ? proj["Project record entry - Project Name"] : null

            if (!projectName || projectName == "SCPT") {
                return
            }

            if (!e["ADM"]) {
                e["Segment"].toLowerCase() == "voice"
                    ? (e["ADM"] = "Arevalo, Milton")
                    : e["Segment"].toLowerCase() == "infra"
                    ? (e["ADM"] = "Manalac, Von Xavier")
                    : (e["ADM"] = null)
            }

            return {
                Segment: e["Segment"],
                PGID: proj?.["HP Project #"],
                "Project Name": projectName,
                ADM: e["ADM"],
                "DXC Project Manager": proj?.["DXC Project Manager"],
                "Project Stage": proj?.["Project Stage"],
                "Project Classification": proj?.["Project Classification"],
                "COMPASS FMO WBS": proj?.["COMPASS FMO WBS"],
            }
        })
        .filter((e) => e)

    return projects
}


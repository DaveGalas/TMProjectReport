
exports.fte = ( data, [TTRdataRef, ppmcWithPGID] ) => {
    return data.map((e) => {
        const { PGID: PGID, "Employee ID": employeeId } = e

        const totalBilledFTE = TTRdataRef.reduce((acc, curr) => {
            if (curr["Project Name"] === PGID && curr["Employee ID"] === employeeId) {
                return acc + curr["Billed FTE"]
            }
            return acc
        }, 0)

        const totalPPMCFTE = ppmcWithPGID.reduce((acc, curr) => {
            if (curr["PGID"] === PGID && curr["Employee ID"] === employeeId) {
                return acc + curr["PPMC FTE"]
            }
            return acc
        }, 0)

        
        const FTEvariance = totalPPMCFTE - totalBilledFTE
        const utilizationRate = totalBilledFTE / totalPPMCFTE

        return {
            ...e,
            "Sum of Actual FTE": totalBilledFTE,
            "Sum of PPMC FTE": totalPPMCFTE,
            "FTE Variance": FTEvariance,
            "Utilization Rate": utilizationRate,
        }
    } )
}

exports.hours = ( data, [TTRdataRef] ) => {
    return data.map((e) => {
        const { PGID, "Employee ID": employeeId,  "Sum of PPMC FTE": SoPF, Country} = e

        const standardHours = Country.toLowerCase() == "india" ? 207 : 184
        const plannedHours = standardHours * SoPF

        const totalHours = TTRdataRef.reduce((acc, curr) => {
            if (curr["Project Name"] === PGID && curr["Employee ID"] === employeeId) {
                return acc + curr["Hours"]
            }
            return acc
        }, 0)
        
        const hoursVariance = plannedHours - totalHours

        return {
            ...e,
            "Actual Hours": totalHours,
            "Planned Hours": plannedHours,
            "Hours Variance": hoursVariance,
        }
    })
}

exports.revenue = ( data , [RATESdataRef]) => {
    return data.map( ( e ) => {
        const {"DXC Job Level":DXCJobLevel, Country,"Actual Hours": AHours,"Planned Hours":PHours} = e

        const key = DXCJobLevel + Country
        const rateRow = RATESdataRef.find((f) => f["CSJL"] == key)
        const rate = rateRow ? rateRow["Rate"] : null
        const actualRevenue = rate * AHours
        const plannedRevenue = rate * PHours
        const revenueVariance = plannedRevenue - actualRevenue
        return {
            ...e,
            "Hourly Revenue Rates (USD)": rate,
            "Actual Revenue": actualRevenue,
            "Planned Revenue": plannedRevenue,
            "Revenue Variance": revenueVariance,
        }
    } )
}

exports.audit = ( data, percentage ) => {
    return data.map((e) => {
        const { "Utilization Rate": uRate } = e
        
        let audit
        if (uRate == 0) audit = "No Tracking"
        if (uRate < percentage && uRate > 0) audit = "Undertracking"
        if (uRate >= percentage) audit = "On Track"
        return { ...e, Audit: audit }
    })
}


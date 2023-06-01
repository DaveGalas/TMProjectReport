const excel = require( "xlsx" )

exports.extractData = ( arg ) => {
    console.log( "...Reading Excel Files" )
    
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
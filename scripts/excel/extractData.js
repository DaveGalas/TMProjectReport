const excel = require( "xlsx" )

exports.extractData = (arg) => {
    console.log("...Reading Excel Files")
    const files = [...arg]
    const xlsxData = files.map( e => {
        const xlsxFile = excel.readFile( e )
        const xlsxSheets = xlsxFile.SheetNames
        
        return excel.utils.sheet_to_json(xlsxFile.Sheets[xlsxSheets[0]])
    })

    return xlsxData
}
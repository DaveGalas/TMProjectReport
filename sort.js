exports.sort = (array, preferences) => {
    return array.sort((a, b) => {
        let result = 0

        preferences.forEach(({ property, isAscending }) => {
            const valueA = a[property]
            const valueB = b[property]

            if (valueA < valueB) {
                result = isAscending ? -1 : 1
            } else if (valueA > valueB) {
                result = isAscending ? 1 : -1
            }

            if (result !== 0) {
                return 
            }
        })

        return result
    })
}

exports.toTwoDigits = (n) => {
  if (String(n).length == 1) {
    return `0${n}`
  } else {
    return n
  }
}

exports.onlyUnique = (value, index, array) => {
  return array.indexOf(value) === index
}

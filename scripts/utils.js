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

exports.formatAsPercentage = (num) => {
  return new Intl.NumberFormat("default", {
    style: "percent",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(num / 100)
}

exports.timeExtract = (TIMEdata, date=new Date()) => {
  const { Percentage: percentage, Date: reportDate } = TIMEdata[ 0 ]
  const today = new Date(date)
  const eoPrevMonth = new Date(today.getFullYear(), today.getMonth(), 1)
  
  return { percentage, reportDate, today, eoPrevMonth }
};

exports.toUnix = ( time ) => {
  const msPerDay = 24 * 60 * 60 * 1000
  const excelEpoch = new Date(Date.UTC(1900, 0, 0))
  const daysOffset = Math.floor( time )
  const msOffset = Math.round((time % 1) * msPerDay)
  const date = new Date(excelEpoch.getTime() + daysOffset * msPerDay + msOffset)
  
  return date
}
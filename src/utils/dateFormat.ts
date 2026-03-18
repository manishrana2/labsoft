export const formatDisplayDate = (value: string): string => {
  const trimmedValue = String(value ?? '').trim()
  if (!trimmedValue) {
    return ''
  }

  const date = /^\d{4}-\d{2}-\d{2}$/.test(trimmedValue)
    ? new Date(`${trimmedValue}T00:00:00`)
    : new Date(trimmedValue)

  if (Number.isNaN(date.getTime())) {
    return trimmedValue
  }

  const day = String(date.getDate()).padStart(2, '0')
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const year = date.getFullYear()
  return `${day}-${month}-${year}`
}

export const formatDisplayDateTime = (value: string): string => {
  const trimmedValue = String(value ?? '').trim()
  if (!trimmedValue) {
    return ''
  }

  const date = new Date(trimmedValue)
  if (Number.isNaN(date.getTime())) {
    return trimmedValue
  }

  const formattedDate = formatDisplayDate(trimmedValue)
  const hours = date.getHours()
  const displayHours = String(hours % 12 || 12).padStart(2, '0')
  const minutes = String(date.getMinutes()).padStart(2, '0')
  const period = hours >= 12 ? 'PM' : 'AM'
  return `${formattedDate} ${displayHours}:${minutes} ${period}`
}

const US_DATE = /^(\d{2})\/(\d{2})\/(\d{4})$/;
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;

// `new Date('YYYY-MM-DD')` is UTC midnight and lands on the previous day in US timezones.
function calendarDate(year: number, month: number, day: number) {
  const date = new Date(year, month - 1, day);
  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }
  return date;
}

function usDateParts(value: string) {
  const match = US_DATE.exec(value);
  if (!match) {
    return null;
  }
  const month = Number(match[1]);
  const day = Number(match[2]);
  const year = Number(match[3]);
  const date = calendarDate(year, month, day);
  if (!date) {
    return null;
  }
  return { year, month, day, date };
}

export function isValidUsDate(value: string) {
  return usDateParts(value) !== null;
}

export function usDateToIso(value: string) {
  const parts = usDateParts(value);
  if (!parts) {
    return null;
  }
  const month = String(parts.month).padStart(2, '0');
  const day = String(parts.day).padStart(2, '0');
  return `${parts.year}-${month}-${day}`;
}

export function isoToUsDate(value: string | null | undefined) {
  if (!value) {
    return '';
  }
  const match = ISO_DATE.exec(value);
  if (!match) {
    return '';
  }
  const date = calendarDate(
    Number(match[1]),
    Number(match[2]),
    Number(match[3]),
  );
  if (!date) {
    return '';
  }
  const parts = new Intl.DateTimeFormat('en-US', {
    month: '2-digit',
    day: '2-digit',
    year: 'numeric',
  }).formatToParts(date);
  const month = parts.find((part) => part.type === 'month')?.value;
  const day = parts.find((part) => part.type === 'day')?.value;
  const year = parts.find((part) => part.type === 'year')?.value;
  if (!month || !day || !year) {
    return '';
  }
  return `${month}/${day}/${year}`;
}

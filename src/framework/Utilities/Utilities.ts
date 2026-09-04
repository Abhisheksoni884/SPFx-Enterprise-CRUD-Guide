/**
 * Formats a Date object or ISO string into a localized date string.
 */
export function formatDate(value?: string | Date): string {
  if (!value) {
    return '';
  }
  const dateObj = typeof value === 'string' ? new Date(value) : value;
  return isNaN(dateObj.getTime()) ? '' : dateObj.toLocaleDateString();
}

/**
 * Checks if date B is earlier than date A (ignoring time if desired, or standard compare).
 */
export function isDateBefore(start: Date, end: Date): boolean {
  const s = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const e = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  return e < s;
}

/**
 * Trims strings safely.
 */
export function safeTrim(value?: string): string {
  return value ? value.trim() : '';
}

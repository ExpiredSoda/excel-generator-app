/**
 * Format time for Excel display
 * Converts 24-hour format (HH:MM) to 12-hour format with AM/PM
 */
export function formatTimeForExcel(time) {
  if (!time) return '';
  
  const [hours, minutes] = time.split(':');
  const hour = parseInt(hours);
  const ampm = hour >= 12 ? 'PM' : 'AM';
  const displayHour = hour % 12 || 12;
  
  return `${displayHour}:${minutes} ${ampm}`;
}

/**
 * Calculate total shift hours
 * Returns formatted hours as string (e.g., "8", "8.5")
 */
export function calculateShiftHours(shifts) {
  if (!shifts.start || !shifts.end) return '0';
  
  const startMinutes = timeToMinutes(shifts.start);
  const endMinutes = timeToMinutes(shifts.end);
  
  let totalMinutes = endMinutes - startMinutes;
  if (totalMinutes < 0) totalMinutes += 24 * 60; // Handle overnight shifts
  
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  
  return minutes > 0 ? `${hours}.${Math.round(minutes/60*10)}` : `${hours}`;
}

/**
 * Convert time string to minutes for comparison
 * Used internally by calculateShiftHours
 */
export function timeToMinutes(timeString) {
  const [hours, minutes] = timeString.split(':').map(Number);
  return hours * 60 + minutes;
}

/**
 * Parse date string in local timezone (avoiding UTC conversion issues)
 * This prevents the common issue where "2025-06-01" becomes May 31st due to timezone offset
 * @param {string} dateString - Date in YYYY-MM-DD format
 * @returns {Date} Date object in local timezone
 */
export function parseLocalDate(dateString) {
  if (!dateString) return new Date();
  
  const [year, month, day] = dateString.split('-').map(Number);
  // Create date in local timezone (month is 0-based in Date constructor)
  return new Date(year, month - 1, day);
}

/**
 * Get month name from date string, timezone-safe
 * @param {string} dateString - Date in YYYY-MM-DD format
 * @returns {string} Month name (e.g., "June")
 */
export function getMonthName(dateString) {
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  
  const date = parseLocalDate(dateString);
  return monthNames[date.getMonth()];
}

/**
 * Get year from date string, timezone-safe
 * @param {string} dateString - Date in YYYY-MM-DD format
 * @returns {number} Year (e.g., 2025)
 */
export function getYear(dateString) {
  const date = parseLocalDate(dateString);
  return date.getFullYear();
} 
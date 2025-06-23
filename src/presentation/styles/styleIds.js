// presentation/styles/styleIds.js
// Style ID constants for easy reference across modules

console.log('✓ StyleIds: Module loaded');

/**
 * Style ID constants for easy reference across generators
 * These map to the index positions in the cell formats array
 * Organized by functionality: Instructions (1-7), Tables (8-17), Calendar (18-21), Utility (22-26), Default (0)
 */
export const STYLE_IDS = {
  // DEFAULT
  DEFAULT: 0,           // Default formatting
  
  // INSTRUCTION STYLES (1-7)
  INSTRUCTION_TITLE: 1,           // Main instruction title (14pt bold, green bg, center, bottom border)
  INSTRUCTION_SECTION_HEADER: 2,  // Section headers (bold 11pt, light blue bg, left)
  INSTRUCTION_HIGHLIGHT: 3,       // Highlighted instructions (regular font, light gray bg, left)
  INSTRUCTION_BULLET: 4,          // Regular bullet points (regular 10pt, no bg, left)
  INSTRUCTION_FOOTER: 5,          // Footer/attribution (9pt italic gray text, no bg, left)
  INSTRUCTION_SPACER: 6,          // Empty spacer rows (default)
  INSTRUCTION_CALLOUT: 7,         // Side notes/tips (gray text, no bg, left)
  
  // TABLE STYLES (8-17)
  TABLE_TITLE: 8,       // Table names (green bg + bottom border)
  TABLE_HEADER: 9,      // Column headers (green bg + all borders)
  TABLE_DATA: 10,       // Data cells (plain + all borders)
  TABLE_INPUT: 11,      // User input cells (light blue + borders)
  TABLE_FORMULA: 12,    // Calculated cells (yellow + borders)
  TABLE_EMPLOYEE: 13,   // Employee names (bold + borders)
  TABLE_TOTAL: 14,      // Sum/total rows (yellow + bold + borders)
  TABLE_ALT_ROW: 15,    // Alternating row color (light gray + borders)
  TABLE_HIGHLIGHT: 16,  // Special emphasis (light green + borders)
  TABLE_WARNING: 17,    // Alerts/errors (light red + borders)
  
  // CALENDAR STYLES (18-21)
  CALENDAR_TITLE: 18,   // Month/year titles (green + bottom border)
  CALENDAR_HEADER: 19,  // Days of week headers (yellow + borders)
  CALENDAR_DAY: 20,     // Date numbers (plain + borders)
  CALENDAR_LEGEND: 21,  // Legend items (lavender + borders)
  // CALENDAR_EVENT styles start at 22+ and are dynamic (created by createCustomCellFormats)
  
  // UTILITY STYLES (22-26)
  HIGHLIGHT_SUCCESS: 22,  // Light green background
  HIGHLIGHT_WARNING: 23,  // Light yellow background
  HIGHLIGHT_INFO: 24,     // Light blue background
  HIGHLIGHT_ERROR: 25,    // Light red background
  DEFAULT_CLEAN: 26       // Absolutely plain (fallback)
};

/**
 * Get the style ID for a custom color (same as getCalendarEventStyleId)
 * @param {number} colorIndex - Index of the custom color (0-based)
 * @returns {number} Style ID for the custom color
 */
export function getCustomStyleId(colorIndex) {
  return getCalendarEventStyleStart() + colorIndex;
}

/**
 * Get the starting style ID for calendar event styles (dynamic colors)
 * Calendar event styles start after all base styles (27 base styles)
 * @returns {number} Starting style ID for calendar events
 */
export function getCalendarEventStyleStart() {
  return 27; // After all 27 base styles (0-26)
}

/**
 * Get style ID for a specific calendar event color
 * @param {number} eventIndex - Index of the event color (0-based)
 * @returns {number} Style ID for the calendar event
 */
export function getCalendarEventStyleId(eventIndex) {
  return getCalendarEventStyleStart() + eventIndex;
}

// presentation/formatting/cellFormats.js
// Complete cell format definitions (combinations of fonts, fills, borders, alignment)

import { FONT_IDS } from '../styles/fonts.js';
import { FILL_IDS } from '../styles/fills.js';
import { BORDER_IDS } from '../styles/borders.js';

/**
 * Base cell format definitions (xf elements)
 * Organized by functionality: Instructions (1-7), Tables (8-17), Calendar (18-21), Utility (22-26), Default (0)
 * Note: CALENDAR_EVENT styles (22+) are dynamic and created by createCustomCellFormats()
 */
export const BASE_CELL_FORMATS = [
  // 0 - Default
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
  
  // INSTRUCTION STYLES (1-7)
  '<xf numFmtId="0" fontId="5" fillId="2" borderId="2" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 1 - INSTRUCTION_TITLE (14pt bold, green bg, center, bottom border)
  '<xf numFmtId="0" fontId="3" fillId="4" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 2 - INSTRUCTION_SECTION_HEADER (bold 11pt, light blue bg, left)
  '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 3 - INSTRUCTION_HIGHLIGHT (regular font, light gray bg, left)
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 4 - INSTRUCTION_BULLET (regular 10pt, no bg, left)
  '<xf numFmtId="0" fontId="9" fillId="0" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 5 - INSTRUCTION_FOOTER (9pt italic gray text, no bg, left)
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 6 - INSTRUCTION_SPACER (default)
  '<xf numFmtId="0" fontId="10" fillId="0" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 7 - INSTRUCTION_CALLOUT (gray text, no bg, left)
    // TABLE STYLES (8-17)
  '<xf numFmtId="0" fontId="3" fillId="2" borderId="2" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 8 - TABLE_TITLE (green bg + bottom border, now centered)
  '<xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 9 - TABLE_HEADER (green bg + all borders, BLACK TEXT)
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 10 - TABLE_DATA (plain + all borders)
  '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 11 - TABLE_INPUT (light blue + borders)
  '<xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 12 - TABLE_FORMULA (yellow + borders)
  '<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 13 - TABLE_EMPLOYEE (bold + borders)
  '<xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 14 - TABLE_TOTAL (yellow + bold + borders)
  '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 15 - TABLE_ALT_ROW (light gray + borders)
  '<xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 16 - TABLE_HIGHLIGHT (light green + borders)
  '<xf numFmtId="0" fontId="0" fillId="12" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 17 - TABLE_WARNING (light red + borders)
  
  // CALENDAR STYLES (18-22)
  '<xf numFmtId="0" fontId="3" fillId="2" borderId="2" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 18 - CALENDAR_TITLE (green + bottom border, month/year)
  '<xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 19 - CALENDAR_HEADER (yellow + borders, days of week)
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 20 - CALENDAR_DAY (plain + borders, date numbers)
  '<xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>', // 21 - CALENDAR_LEGEND (lavender + borders, legend items)
  // 22 - CALENDAR_EVENT will be dynamic custom colors, handled by createCustomCellFormats
  
  // UTILITY STYLES (22-26)
  '<xf numFmtId="0" fontId="0" fillId="9" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 22 - HIGHLIGHT_SUCCESS (light green background)
  '<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 23 - HIGHLIGHT_WARNING (light yellow background)
  '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 24 - HIGHLIGHT_INFO (light blue background)
  '<xf numFmtId="0" fontId="0" fillId="12" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>', // 25 - HIGHLIGHT_ERROR (light red background)
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' // 26 - DEFAULT_CLEAN (absolutely plain fallback)
];

/**
 * Generate custom cell formats for dynamic colors (calendar events, etc.)
 * @param {Array} customColors - Array of color strings
 * @param {number} baseFillCount - Number of base fills to offset custom fill IDs
 * @returns {Array} Array of custom cell format XML strings
 */
export function createCustomCellFormats(customColors = [], baseFillCount = 13) {
  return customColors.map((color, index) => 
    `<xf numFmtId="0" fontId="0" fillId="${baseFillCount + index}" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>`
  );
}

/**
 * Get all cell formats including custom ones
 * Dynamic calendar event styles start at index 27 (after the 27 base styles)
 * @param {Array} customColors - Array of custom color strings
 * @param {number} baseFillCount - Number of base fills
 * @returns {Array} Complete array of cell format definitions
 */
export function getAllCellFormats(customColors = [], baseFillCount = 13) {
  const customFormats = createCustomCellFormats(customColors, baseFillCount);
  return [...BASE_CELL_FORMATS, ...customFormats];
}

/**
 * Helper functions for style creation and management
 */

/**
 * Get style information by ID
 * @param {number} styleId - The style ID to look up
 * @returns {Object} Style information including name, description, and usage
 */
export function getStyleInfo(styleId) {
  const styleMap = {
    0: { name: 'DEFAULT', description: 'Default formatting', usage: 'Empty cells, spacers' },
    
    // Instruction Styles
    1: { name: 'INSTRUCTION_TITLE', description: '14pt bold, green bg, center, bottom border', usage: 'Main instruction sheet titles' },
    2: { name: 'INSTRUCTION_SECTION_HEADER', description: 'Bold 11pt, light blue bg, left', usage: 'Section headers like "1. MAIN TRACKER SHEET:"' },
    3: { name: 'INSTRUCTION_HIGHLIGHT', description: 'Regular font, light gray bg, left', usage: 'Important tips and highlighted instructions' },
    4: { name: 'INSTRUCTION_BULLET', description: 'Regular 10pt, no bg, left', usage: 'Regular bullet point text' },
    5: { name: 'INSTRUCTION_FOOTER', description: '9pt italic gray text, no bg, left', usage: 'Footer attribution text' },
    6: { name: 'INSTRUCTION_SPACER', description: 'Default formatting', usage: 'Empty rows between sections' },
    7: { name: 'INSTRUCTION_CALLOUT', description: 'Gray text, no bg, left', usage: 'Side notes and tips' },
    
    // Table Styles
    8: { name: 'TABLE_TITLE', description: 'Green bg + bottom border', usage: 'Table names and titles' },
    9: { name: 'TABLE_HEADER', description: 'Green bg + all borders', usage: 'Column headers' },
    10: { name: 'TABLE_DATA', description: 'Plain + all borders', usage: 'Regular data cells' },
    11: { name: 'TABLE_INPUT', description: 'Light blue + borders', usage: 'User input cells' },
    12: { name: 'TABLE_FORMULA', description: 'Yellow + borders', usage: 'Calculated cells' },
    13: { name: 'TABLE_EMPLOYEE', description: 'Bold + borders', usage: 'Employee names' },
    14: { name: 'TABLE_TOTAL', description: 'Yellow + bold + borders', usage: 'Sum/total rows' },
    15: { name: 'TABLE_ALT_ROW', description: 'Light gray + borders', usage: 'Alternating row color' },
    16: { name: 'TABLE_HIGHLIGHT', description: 'Light green + borders', usage: 'Special emphasis' },
    17: { name: 'TABLE_WARNING', description: 'Light red + borders', usage: 'Alerts and errors' },
    
    // Calendar Styles
    18: { name: 'CALENDAR_TITLE', description: 'Green + bottom border', usage: 'Month/year titles' },
    19: { name: 'CALENDAR_HEADER', description: 'Yellow + borders', usage: 'Days of week headers' },
    20: { name: 'CALENDAR_DAY', description: 'Plain + borders', usage: 'Date numbers' },
    21: { name: 'CALENDAR_LEGEND', description: 'Lavender + borders', usage: 'Legend items' },
    
    // Utility Styles
    22: { name: 'HIGHLIGHT_SUCCESS', description: 'Light green background', usage: 'Success messages, positive states' },
    23: { name: 'HIGHLIGHT_WARNING', description: 'Light yellow background', usage: 'Warning messages, caution states' },
    24: { name: 'HIGHLIGHT_INFO', description: 'Light blue background', usage: 'Information messages, neutral states' },
    25: { name: 'HIGHLIGHT_ERROR', description: 'Light red background', usage: 'Error messages, negative states' },
    26: { name: 'DEFAULT_CLEAN', description: 'Absolutely plain', usage: 'Fallback, minimal formatting' }
  };
  
  return styleMap[styleId] || { name: 'UNKNOWN', description: 'Unknown style', usage: 'Not defined' };
}

/**
 * Create a custom instruction style with specific formatting
 * @param {Object} options - Style options
 * @param {number} options.fontId - Font ID to use
 * @param {number} options.fillId - Fill ID to use  
 * @param {number} options.borderId - Border ID to use
 * @param {string} options.alignment - Horizontal alignment (left, center, right)
 * @param {boolean} options.wrapText - Whether to wrap text
 * @returns {string} Cell format XML string
 */
export function createInstructionStyle({ fontId = 0, fillId = 0, borderId = 0, alignment = 'left', wrapText = false }) {
  const wrapAttr = wrapText ? ' wrapText="1"' : '';
  return `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"><alignment horizontal="${alignment}" vertical="center"${wrapAttr}/></xf>`;
}

/**
 * Create a custom table style with specific formatting
 * @param {Object} options - Style options
 * @param {number} options.fontId - Font ID to use
 * @param {number} options.fillId - Fill ID to use
 * @param {number} options.borderId - Border ID to use (defaults to 1 for all borders)
 * @param {string} options.alignment - Horizontal alignment (left, center, right)
 * @returns {string} Cell format XML string
 */
export function createTableStyle({ fontId = 0, fillId = 0, borderId = 1, alignment = 'center' }) {
  return `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"><alignment horizontal="${alignment}" vertical="center"/></xf>`;
}

/**
 * Create a custom calendar style with specific formatting
 * @param {Object} options - Style options
 * @param {number} options.fontId - Font ID to use
 * @param {number} options.fillId - Fill ID to use
 * @param {number} options.borderId - Border ID to use (defaults to 1 for all borders)
 * @param {string} options.alignment - Horizontal alignment (defaults to center for calendar)
 * @returns {string} Cell format XML string
 */
export function createCalendarStyle({ fontId = 0, fillId = 0, borderId = 1, alignment = 'center' }) {
  return `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"><alignment horizontal="${alignment}" vertical="center"/></xf>`;
}

/**
 * Create a utility highlight style with specific formatting
 * @param {Object} options - Style options
 * @param {number} options.fontId - Font ID to use
 * @param {number} options.fillId - Fill ID to use
 * @param {number} options.borderId - Border ID to use (defaults to 0 for no borders)
 * @param {string} options.alignment - Horizontal alignment (defaults to left for utility)
 * @returns {string} Cell format XML string
 */
export function createUtilityStyle({ fontId = 0, fillId = 0, borderId = 0, alignment = 'left' }) {
  return `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"><alignment horizontal="${alignment}" vertical="center"/></xf>`;
}

/**
 * Validate style ID is within valid range
 * @param {number} styleId - Style ID to validate
 * @returns {boolean} True if valid, false if out of range
 */
export function isValidStyleId(styleId) {
  return typeof styleId === 'number' && styleId >= 0 && styleId < BASE_CELL_FORMATS.length;
}

/* ========================================
   STYLE GLOSSARY & USAGE GUIDE
   ======================================== */

/**
 * EXCEL GENERATOR STYLE SYSTEM GLOSSARY
 * 
 * This glossary explains when and how to use each style category in the Excel Generator.
 * For technical details (font IDs, fill IDs, etc.), use the getStyleInfo() function.
 * 
 * INSTRUCTION STYLES - For documentation and help sheets
 * 
 * These styles are designed for instruction sheets that explain how to use the generated
 * Excel files. They use minimal borders to prevent visual clutter across wide content areas.
 * 
 * • Title Styles: Use for main headings like "HOW TO USE THIS ATTENDANCE TRACKER"
 * • Section Headers: Use for numbered sections like "1. ENTERING DATA" or "2. VIEWING REPORTS"
 * • Highlights: Use to draw attention to important warnings or tips
 * • Bullets: Use for step-by-step instructions and lists
 * • Footers: Use for attribution, version info, or disclaimers
 * • Spacers: Use for visual separation between content sections
 * • Callouts: Use for side notes, tips, or additional context
 * 
 * TABLE STYLES - For structured data presentation
 * 
 * These styles create professional-looking data tables with consistent borders and colors
 * that clearly distinguish different types of content and user interaction areas.
 * 
 * • Titles: Use above tables to label what the data represents
 * • Headers: Use for column titles like "Employee Name", "Hours Worked", "Total"
 * • Data: Use for regular information that users view but don't typically edit
 * • Input: Use for cells where users are expected to enter information
 * • Formulas: Use for calculated cells to show they're auto-computed
 * • Employee Names: Use for person identifiers to make them stand out
 * • Totals: Use for sum rows and important calculated results
 * • Alternating Rows: Use every other row in large tables for easier reading
 * • Highlights: Use to emphasize important data points or achievements
 * • Warnings: Use for error conditions, missing data, or problems requiring attention
 * 
 * CALENDAR STYLES - For date-based layouts
 * 
 * These styles create clear, readable calendar layouts that balance functionality with
 * visual appeal, making it easy to distinguish dates from events and legends.
 * 
 * • Titles: Use for month/year headers like "DECEMBER 2024"
 * • Headers: Use for day-of-week labels like "MON", "TUE", "WED"
 * • Day Numbers: Use for the actual date numbers (1, 2, 3, etc.)
 * • Legends: Use for explaining what different colors or symbols mean
 * • Events: Use dynamic colors for different types of events or appointments
 * 
 * UTILITY STYLES - For status indication and feedback
 * 
 * These styles provide semantic meaning through color to quickly communicate status,
 * conditions, or feedback to users without requiring detailed reading.
 * 
 * • Success: Use for positive outcomes, completed tasks, or confirmation messages
 * • Warning: Use for caution states, potential issues, or "pay attention" areas
 * • Info: Use for neutral information, helpful tips, or general notices
 * • Error: Use for problems, failures, or critical issues requiring immediate action
 * • Clean Default: Use when you need absolutely minimal formatting as a fallback
 * 
 * CHOOSING THE RIGHT STYLE
 * 
 * Ask yourself these questions:
 * 1. What is the purpose of this content? (instruction, data, calendar, status)
 * 2. How should users interact with it? (read-only, input required, calculated)
 * 3. What level of attention does it need? (title, normal, highlight, warning)
 * 4. Does it need borders? (tables yes, instructions minimal, utilities no)
 * 
 * STYLE COMBINATIONS TO AVOID
 * 
 * • Don't mix instruction styles in data tables (different border patterns)
 * • Don't use utility styles for regular content (colors have semantic meaning)
 * • Don't use table styles in instruction sheets (too many borders)
 * • Don't use calendar styles outside of calendar contexts (alignment assumptions)
 * 
 * ACCESSIBILITY AND CLARITY
 * 
 * • Colors always have semantic meaning (green=good/headers, yellow=calculated/caution, red=problems)
 * • Font weights create clear hierarchy (bold=important, regular=content, gray=supplemental)
 * • Borders define data boundaries (full borders=structured data, minimal=flowing content)
 * • Alignment supports content type (center=data/numbers, left=text/instructions)
 * 
 * COMMON PATTERNS
 * 
 * Instruction Sheet Layout:
 * [TITLE] -> [SECTION_HEADER] -> [BULLET/HIGHLIGHT] -> [SPACER] -> repeat
 * 
 * Data Table Layout:
 * [TABLE_TITLE] -> [TABLE_HEADER row] -> [TABLE_DATA/INPUT/FORMULA rows] -> [TABLE_TOTAL]
 * 
 * Calendar Layout:
 * [CALENDAR_TITLE] -> [CALENDAR_HEADER row] -> [CALENDAR_DAY grid] -> [CALENDAR_LEGEND]
 * 
 * Status/Feedback:
 * Use UTILITY styles sparingly for important status communication only
 */

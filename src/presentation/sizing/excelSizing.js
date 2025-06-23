// presentation/sizing/excelSizing.js
// Universal auto-sizing utilities for Excel worksheets

// Debug: Track module loading
console.log('✓ ExcelSizing: Module loaded');

/**
 * Row height constants for different content types
 */
export const ROW_HEIGHTS = {
  title: 35,        // Main titles, headers
  subtitle: 30,     // Section headers, day-of-week headers
  header: 25,       // Table headers, legend items
  content: 20,      // Standard content
  data: 25,         // Data rows
  event: 40,        // Event rows with potential wrapping
  spacer: 10,       // Empty spacer rows
  instruction: 20   // Instruction text
};

/**
 * Column width constants for different content types
 */
export const COLUMN_WIDTHS = {
  narrow: 7,        // Short codes, numbers
  standard: 10,     // Default column width
  medium: 12,       // Time fields, short text
  wide: 15,         // Calendar days, medium text
  name: 20,         // Employee names, longer text
  description: 25,  // Job titles, descriptions
  spacer: 3,        // Narrow spacer columns
  instruction: 80   // Wide instruction columns
};

/**
 * Creates column definitions with auto-sizing
 * @param {Array} columns - Array of column configurations
 * @returns {Array} Column definitions for Excel
 */
export function createColumnDefinitions(columns) {
  return columns.map(col => {
    const definition = {
      min: col.min,
      max: col.max,
      width: col.width || COLUMN_WIDTHS.standard
    };
    
    // Add auto-sizing attributes if requested
    if (col.bestFit !== false) {
      definition.bestFit = 1;
    }
    
    if (col.customWidth !== false) {
      definition.customWidth = 1;
    }
    
    return definition;
  });
}

/**
 * Creates calendar-specific column layout
 * @returns {Array} Column definitions for calendar sheets
 */
export function createCalendarColumns() {
  const cols = [];
  
  // Calendar day columns (A-G) with auto-sizing
  for (let c = 1; c <= 7; c++) {
    cols.push({
      min: c,
      max: c,
      width: COLUMN_WIDTHS.wide,
      bestFit: true,
      customWidth: true
    });
  }
  
  // Spacer column (H)
  cols.push({
    min: 8,
    max: 8,
    width: COLUMN_WIDTHS.spacer,
    bestFit: false,
    customWidth: true
  });
  
  // Legend description column (I)
  cols.push({
    min: 9,
    max: 9,
    width: COLUMN_WIDTHS.description - 7, // 18
    bestFit: true,
    customWidth: true
  });
  
  // Legend color column (J)
  cols.push({
    min: 10,
    max: 10,
    width: COLUMN_WIDTHS.narrow,
    bestFit: true,
    customWidth: true
  });
  
  return createColumnDefinitions(cols);
}

/**
 * Creates attendance tracker column layout
 * @returns {Array} Column definitions for attendance sheets
 */
export function createAttendanceColumns() {
  const columns = [
    { min: 1, max: 1, width: COLUMN_WIDTHS.name },              // Employee Name (20)
    { min: 2, max: 2, width: COLUMN_WIDTHS.standard },          // ID (10)
    { min: 3, max: 3, width: COLUMN_WIDTHS.description - 7 },   // Job Title (18)
    { min: 4, max: 4, width: COLUMN_WIDTHS.wide + 2 },          // Shift Start (17)
    { min: 5, max: 5, width: COLUMN_WIDTHS.wide + 2 },          // First Break (17)
    { min: 6, max: 6, width: COLUMN_WIDTHS.wide + 2 },          // Lunch Break (17) 
    { min: 7, max: 7, width: COLUMN_WIDTHS.wide + 3 },          // Second Break (18)
    { min: 8, max: 8, width: COLUMN_WIDTHS.wide },              // Shift End (15)
    { min: 9, max: 9, width: COLUMN_WIDTHS.standard },          // Monday (10)
    { min: 10, max: 10, width: COLUMN_WIDTHS.standard },        // Tuesday (10)
    { min: 11, max: 11, width: COLUMN_WIDTHS.standard },        // Wednesday (10)
    { min: 12, max: 12, width: COLUMN_WIDTHS.standard },        // Thursday (10)
    { min: 13, max: 13, width: COLUMN_WIDTHS.standard },        // Friday (10)
    { min: 14, max: 14, width: COLUMN_WIDTHS.standard },        // Saturday (10)
    { min: 15, max: 15, width: COLUMN_WIDTHS.standard },        // Sunday (10)
    { min: 16, max: 16, width: COLUMN_WIDTHS.medium + 3 }       // Total Hours (15)
  ];
  
  return createColumnDefinitions(columns);
}

/**
 * Creates instruction sheet column layout
 * @returns {Array} Column definitions for instruction sheets
 */
export function createInstructionColumns() {
  return createColumnDefinitions([{
    min: 1,
    max: 1,
    width: COLUMN_WIDTHS.instruction,
    bestFit: false,
    customWidth: true
  }]);
}

/**
 * Creates tracker sheet column layout (reference/summary sheets)
 * @returns {Array} Column definitions for tracker sheets
 */
export function createTrackerColumns() {
  return createColumnDefinitions([
    { min: 1, max: 1, width: COLUMN_WIDTHS.name },     // Employee Name
    { min: 2, max: 2, width: COLUMN_WIDTHS.standard }, // ID/Code
    { min: 3, max: 3, width: 50 }                      // Details/Notes (extra wide)
  ]);
}



/**
 * Generates XML for column definitions with auto-sizing support
 * Excel auto-sizing behavior:
 * - bestFit="1" enables Excel to recalculate column widths on open
 * - customWidth="1" indicates the width was manually set
 * - Setting width="0" tells Excel to auto-calculate optimal width
 * 
 * @param {Array} columnDefs - Array of column definitions
 * @param {Object} options - Auto-sizing options
 * @param {boolean} options.enableAutoWidth - Use width="0" for Excel auto-calculation
 * @param {boolean} options.includeSheetFormat - Include sheetFormatPr element
 * @returns {string} XML string for columns with auto-sizing
 */
export function generateColumnsXML(columnDefs, options = {}) {
  if (!columnDefs || columnDefs.length === 0) {
    return '';
  }
  
  const { enableAutoWidth = false, includeSheetFormat = false } = options;
  
  const colsXML = columnDefs.map(col => {
    const width = enableAutoWidth ? 0 : col.width;
    let xml = `<col min="${col.min}" max="${col.max}" width="${width}"`;
    
    if (col.bestFit !== false) xml += ' bestFit="1"';
    if (col.customWidth !== false) xml += ' customWidth="1"';
    
    xml += '/>';
    return xml;
  }).join('');
  
  let result = `<cols>${colsXML}</cols>`;
  
  // Add sheet-level formatting if requested
  if (includeSheetFormat) {
    const defaultColWidth = enableAutoWidth ? 0 : 8.43;
    const sheetFormatPr = `<sheetFormatPr defaultRowHeight="15" defaultColWidth="${defaultColWidth}" baseColWidth="10" customHeight="0"/>`;
    result = `${sheetFormatPr}${result}`;
  }
  
  return result;
}

/**
 * Generates row opening tag with proper height attributes
 * Excel auto-height behavior:
 * - customHeight="1" with ht value = fixed height (no auto-sizing)
 * - customHeight="0" or omitted = tells Excel to auto-size when opened
 * - enableAutoHeight=true = omits customHeight, allowing Excel to calculate on open
 * 
 * @param {number} rowNumber - Row number (1-based)
 * @param {string} heightType - Height type from ROW_HEIGHTS
 * @param {Object} options - Additional row options
 * @param {boolean} options.enableAutoHeight - If true, allows Excel auto-calculation
 * @param {number} options.styleId - Style ID to apply to the row
 * @returns {string} Opening row tag with proper height attributes
 */
export function generateRowXML(rowNumber, heightType = 'content', options = {}) {
  const {
    styleId = null,
    enableAutoHeight = false,
    hidden = false,
    collapsed = false,
    outlineLevel = 0
  } = options;
  
  let xml = `<row r="${rowNumber}"`;
  
  // Handle height attributes based on auto-sizing preference
  if (enableAutoHeight) {
    // Let Excel calculate height when the file is opened
    // Row will appear with default height until opened in Excel
  } else {
    // Set specific height for immediate visual consistency
    const height = ROW_HEIGHTS[heightType] || ROW_HEIGHTS.content;
    xml += ` ht="${height}" customHeight="1"`;
  }
  
  // Add style if provided
  if (styleId !== null) {
    xml += ` s="${styleId}"`;
  }
  
  // Add visibility attributes
  if (hidden) xml += ' hidden="1"';
  if (collapsed) xml += ' collapsed="1"';
  if (outlineLevel > 0) xml += ` outlineLevel="${outlineLevel}"`;
  
  xml += '>';
  return xml;
}

/**
 * Applies auto-sizing configuration to a sheet based on sheet type
 * @param {ExcelSheet} sheet - Sheet instance to configure
 * @param {string} sheetType - Type of sheet (calendar, attendance, tracker, instruction)
 * @param {Object} options - Additional options for auto-sizing
 * @returns {boolean} True if auto-sizing was successfully applied
 */
export function applyAutoSizing(sheet, sheetType = 'standard', options = {}) {
  let columnDefs;
  
  // Get appropriate column definitions based on sheet type
  switch (sheetType) {
    case 'calendar':
      columnDefs = createCalendarColumns();
      break;
    case 'attendance':
      columnDefs = createAttendanceColumns();
      break;
    case 'tracker':
      columnDefs = createTrackerColumns();
      break;
    case 'instruction':
      columnDefs = createInstructionColumns();
      break;
    default:
      columnDefs = createColumnDefinitions([
        { min: 1, max: 1, width: COLUMN_WIDTHS.standard, bestFit: true, customWidth: true }
      ]);
  }
  
  // Apply column definitions
  if (sheet.setCols && columnDefs) {
    sheet.setCols(columnDefs);
  }
  
  // Apply additional options
  if (options.enableAutoFilter && sheet.autoFilter !== undefined) {
    sheet.autoFilter = true;
  }
  
  if (options.enableFreezePanes && sheet.freezePanes !== undefined) {
    sheet.freezePanes = options.freezePanes || { row: 1, col: 0 };
  }
  
  console.log(`✅ AutoSizing: Applied for ${sheetType} sheet (${columnDefs.length} columns)`);
  return true;
}

/**
 * EXCEL AUTO-HEIGHT BEHAVIOR REFERENCE:
 * 
 * Key insight: Excel auto-height is calculated BY EXCEL after opening the file, not during generation.
 * 
 * During file generation:
 * - enableAutoHeight=false → Fixed height with customHeight="1" and ht="value"
 * - enableAutoHeight=true → No height attributes, Excel calculates when opened
 * 
 * When user opens file in Excel:
 * - Fixed height rows → Stay at specified height
 * - Auto-height rows → Excel calculates based on content, font, cell width, text wrapping
 * 
 * This is why auto-height appears to "not work" during generation - it's working correctly,
 * but only becomes visible after opening in Excel.
 */

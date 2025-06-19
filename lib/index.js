// Import all the generator functions for the convenience function
import { buildCalendarSheetWithExcelBuilder } from './generators/calendarSheet.js';
import { getStylesXML } from './generators/stylesXml.js';
import { getContentTypesXML, getRelsXML } from './generators/contentTypesXml.js';
import { getWorkbookXML, getWorkbookRelsXML } from './generators/workbookXml.js';
import { getTrackerSheetXML } from './generators/trackerSheet.js';

// Core classes (from excelCore.js)
export { 
  ExcelBuilder, 
  ExcelSheet, 
  ExcelCell, 
  ExcelRow, 
  ConditionalFormattingRule,
  escapeXml
} from './core/excelCore.js';

// Re-export generators
export { buildCalendarSheetWithExcelBuilder };
export { getStylesXML };
export { getContentTypesXML, getRelsXML };
export { getWorkbookXML, getWorkbookRelsXML };
export { getTrackerSheetXML };

// Utilities
export { createZip } from './utils/zipWriter.js';
export { sanitizeLegendInput, validateLegendInput } from './utils/sanitize.js';
export { validateColorSelection, rgbToHex } from './utils/validation.js';

// Complete Excel Generator - includes tracker option
export function generateCompleteExcel(options = {}) {
  const {
    year = new Date().getFullYear(),
    month = new Date().getMonth(),
    eventRows = 3,
    includeTracker = false,
    legendValues = ['Meeting', 'Holiday', 'Personal'],
    customColors = ['FFDC143C', 'FF228B22', 'FF1E90FF']
  } = options;

  // Now the functions are properly imported and available in scope
  const calendarXML = buildCalendarSheetWithExcelBuilder(
    year, month, eventRows, false, legendValues, customColors
  );
  
  const stylesXML = getStylesXML(eventRows, customColors);
  const contentTypesXML = getContentTypesXML(includeTracker);
  const workbookXML = getWorkbookXML(includeTracker);
  const relsXML = getRelsXML();
  const workbookRelsXML = getWorkbookRelsXML(includeTracker);

  // Base files
  const files = [
    { name: '[Content_Types].xml', content: contentTypesXML },
    { name: '_rels/.rels', content: relsXML },
    { name: 'xl/workbook.xml', content: workbookXML },
    { name: 'xl/worksheets/sheet1.xml', content: calendarXML },
    { name: 'xl/styles.xml', content: stylesXML },
    { name: 'xl/_rels/workbook.xml.rels', content: workbookRelsXML }
  ];

  // Add tracker sheet if requested
  if (includeTracker) {
    const trackerXML = getTrackerSheetXML(legendValues);
    files.push({ name: 'xl/worksheets/sheet2.xml', content: trackerXML });
  }

  return files;
}

// Version info
export const version = '2.0.0';
export const name = 'Excel Generator Library';

// Quick start example
export const quickStart = `
import { 
  ExcelBuilder, 
  buildCalendarSheetWithExcelBuilder, 
  getStylesXML,
  createZip 
} from './lib/index.js';

// Generate a calendar
const calendarXML = buildCalendarSheetWithExcelBuilder(2024, 0, 3, false, 
  ['Meeting', 'Holiday', 'Personal'], 
  ['FFDC143C', 'FF228B22', 'FF1E90FF']
);

const stylesXML = getStylesXML(3, ['FFDC143C', 'FF228B22', 'FF1E90FF']);

// Create complete Excel file
const files = [
  { name: 'xl/worksheets/sheet1.xml', content: calendarXML },
  { name: 'xl/styles.xml', content: stylesXML }
  // ... add other required files
];

const zipBytes = createZip(files);
// Download or use the Excel file
`;
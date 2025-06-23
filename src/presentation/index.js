// presentation/index.js
// Main entry point for all presentation modules

// Re-export styles components
export { FONTS, FONT_IDS } from './styles/fonts.js';
export { BASE_FILLS, FILL_IDS, createCustomFills, getAllFills } from './styles/fills.js';
export { BORDERS, BORDER_IDS } from './styles/borders.js';
export { COLORS } from './styles/colors.js';
export { STYLE_IDS, getCustomStyleId } from './styles/styleIds.js';
export { generateStylesXML, getUniversalStylesXML } from './styles/stylesXml.js';

// Re-export formatting components
export { BASE_CELL_FORMATS, createCustomCellFormats, getAllCellFormats } from './formatting/cellFormats.js';
export { createDxfElements, generateDxfXML } from './formatting/dxfFormats.js';

// Re-export sizing components
export { 
  ROW_HEIGHTS, 
  COLUMN_WIDTHS,   createColumnDefinitions,
  createCalendarColumns,
  createTrackerColumns,
  createAttendanceColumns,
  createInstructionColumns,
  applyAutoSizing,
  generateColumnsXML,
  generateRowXML
} from './sizing/excelSizing.js';

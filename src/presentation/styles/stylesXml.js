// presentation/styles/stylesXml.js
// Main styles XML generator - assembles all components

import { FONTS } from './fonts.js';
import { getAllFills } from './fills.js';
import { BORDERS } from './borders.js';
import { getAllCellFormats } from '../formatting/cellFormats.js';
import { generateDxfXML } from '../formatting/dxfFormats.js';

/**
 * Generate complete styles XML for Excel workbook
 * @param {Array} customColors - Array of custom color strings for calendar legends, etc.
 * @returns {string} Complete styles.xml content
 */
export function generateStylesXML(customColors = []) {
  // Get all component arrays
  const allFills = getAllFills(customColors);
  const allCellFormats = getAllCellFormats(customColors, allFills.length - customColors.length);
  const dxfXML = generateDxfXML(customColors);

  // Assemble complete XML
  return `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="${FONTS.length}">${FONTS.join('')}</fonts>
  <fills count="${allFills.length}">${allFills.join('')}</fills>
  <borders count="${BORDERS.length}">${BORDERS.join('')}</borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${allCellFormats.length}">${allCellFormats.join('')}</cellXfs>
  ${dxfXML}
</styleSheet>`;
}

// Backward compatibility - expose the original function name
export const getUniversalStylesXML = generateStylesXML;

// generators/trackerSheet.js
// Generates the tracker sheet XML for Excel
import { escapeXml } from '../core/excelCore.js';

/**
 * Generate tracker sheet XML with automatic counting formulas
 * @param {Array} legendValues - Array of legend values to track
 * @returns {string} - Complete tracker sheet XML
 */
export function getTrackerSheetXML(legendValues = ['Meeting', 'Holiday', 'Personal']) {
  let xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="20"/>
    <col min="2" max="2" width="10"/>
    <col min="3" max="3" width="50" bestFit="1"/>
  </cols>
  <sheetData>`;
  
  // Header row with bold styling
  xml += `<row r="1">
    <c r="A1" t="inlineStr" s="7"><is><t>Legend Value</t></is></c>
    <c r="B1" t="inlineStr" s="7"><is><t>Count</t></is></c>
    <c r="C1" t="inlineStr" s="7"><is><t>Description</t></is></c>
  </row>`;
    // Data rows with actual legend values and counting formulas
  legendValues.forEach((value, index) => {
    const rowNum = index + 2;
    const cellA = `A${rowNum}`;
    const cellB = `B${rowNum}`;
    const cellC = `C${rowNum}`;
    
    // Clean the legend value for use in formulas
    const cleanValue = escapeTrackerXml(value.toString().trim());
    
    // Create COUNTIF formula with correct Excel syntax (Calendar!A:G not Calendar.A:G)
    const countFormula = `COUNTIF(Calendar!A:G,"${cleanValue}")`;
    
    xml += `<row r="${rowNum}">
    <c r="${cellA}" t="inlineStr"><is><t>${cleanValue}</t></is></c>
    <c r="${cellB}"><f>${countFormula}</f><v>0</v></c>
    <c r="${cellC}" t="inlineStr"><is><t>Count of "${cleanValue}" entries in the calendar sheet</t></is></c>
  </row>`;
  });
  
  // Remove the Total row - it's not in the main app
  xml += `</sheetData>
</worksheet>`;
  
  return xml;
}

/**
 * Escape XML special characters for tracker sheet (renamed to avoid conflicts)
 * @param {string} unsafe - String that may contain XML special characters
 * @returns {string} - XML-safe string
 */
function escapeTrackerXml(unsafe) {
  if (typeof unsafe !== 'string') return '';
  return unsafe.replace(/[<>&'"]/g, function (c) {
    switch (c) {
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '&': return '&amp;';
      case '\'': return '&apos;';
      case '"': return '&quot;';
      default: return c;
    }
  });
}

// legendSheet.js
// Generates the Legends sheet for the Attendance Tracker Excel file

import { ExcelSheet } from '../../core/excelSheet.js';
import { ExcelRow } from '../../core/excelRow.js';
import { ExcelCell } from '../../core/excelCell.js';
import { getCustomStyleId } from '../../../presentation/index.js';

/**
 * Build the Legends sheet XML for Excel
 * @param {Array<{label: string, color: string}>} legends - Array of legend objects with label and color (hex)
 * @returns {string} XML for the Legends worksheet
 */
export function buildLegendSheet(legends) {
  const sheet = new ExcelSheet('Legends');

  // Title row (merged)
  const titleRow = new ExcelRow(1);
  titleRow.addCell(new ExcelCell('A', 1, 'Legend', { style: 1 }));
  titleRow.addCell(new ExcelCell('B', 1, '', { style: 1 }));
  sheet.addRow(titleRow);

  // Header row
  const headerRow = new ExcelRow(2);
  headerRow.addCell(new ExcelCell('A', 2, 'Label', { style: 2 }));
  headerRow.addCell(new ExcelCell('B', 2, 'Color', { style: 2 }));
  sheet.addRow(headerRow);

  // Legend rows (no hex code as text, just label and color cell with fill)
  legends.forEach((legend, i) => {
    const rowNum = i + 3;
    const row = new ExcelRow(rowNum);
    row.addCell(new ExcelCell('A', rowNum, legend.label, { style: 10 }));
    row.addCell(new ExcelCell('B', rowNum, '', { style: getCustomStyleId(i) }));
    sheet.addRow(row);
  });

  // Merge title row across both columns
  sheet.addMerge('A1:B1');

  // Set column widths: A wider for labels, B narrow for color
  sheet.setCols([
    { min: 1, max: 1, width: 28, bestFit: true, customWidth: true }, // Column A: wide for legend names
    { min: 2, max: 2, width: 10, bestFit: true, customWidth: true }   // Column B: narrow for color
  ]);

  return sheet.toXML();
} 
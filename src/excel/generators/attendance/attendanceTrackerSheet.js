// generators/attendanceTrackerSheet.js
// Generates the main shift tracker worksheet XML
import { escapeXml, ExcelSheet, ConditionalFormattingRule } from '../../core/index.js';
import { STYLE_IDS } from '../../../presentation/index.js';
import { calculateEmployeeColumnWidths, generateColumnsXML, generateRowXML, ROW_HEIGHTS } from '../../../presentation/sizing/excelSizing.js';
import { DataValidationRule, generateDataValidationsXML } from '../../core/dataValidation.js';
import { formatTimeForExcel, getMonthName, getYear } from '../../../shared/utils/timeUtils.js';

export function buildShiftTrackerSheet(employees, dates, legends = []) {
  // Calculate month/year from the first date for sheet naming (timezone-safe)
  let sheetName = "Shift Tracker";
  if (dates && dates.length > 0) {
    const month = getMonthName(dates[0]);
    const year = getYear(dates[0]);
    sheetName = `Shift Tracker ${month} ${year}`;
  }

  // Employee info columns
  const infoHeaders = [
    'Employee Name', 'ID', 'Job Title', 'Shift Start', 'First Break', 
    'Lunch Break', 'Second Break', 'Shift End'
  ];
  // Dynamic date columns
  const dateHeaders = dates.map(dateStr => {
    const [year, month, day] = dateStr.split('-').map(Number);
    const dateObj = new Date(year, month - 1, day);
    const weekday = dateObj.toLocaleDateString('en-US', { weekday: 'short' });
    return `${weekday} ${month}/${day}`;
  });
  const headers = [...infoHeaders, ...dateHeaders, 'Total Hours'];

  // Helper to get Excel column letters (supports AA, AB, etc.)
  function getExcelColLetter(n) {
    let s = '';
    while (n >= 0) {
      s = String.fromCharCode((n % 26) + 65) + s;
      n = Math.floor(n / 26) - 1;
    }
    return s;
  }

  // Calculate optimal column width based on longest legend label
  function calculateLegendColumnWidth(legends) {
    if (!legends || legends.length === 0) return 12; // Default width
    
    // Find the longest legend label
    const longestLabel = legends.reduce((longest, legend) => {
      const label = legend.label || '';
      return label.length > longest.length ? label : longest;
    }, '');
    
    // Excel width calculation: approximately 1 character = 1 unit, with some padding
    // Add extra padding for dropdown arrow and cell padding
    const baseWidth = longestLabel.length * 1.2; // 1.2 units per character for better spacing
    const minWidth = 8;  // Minimum usable width
    const maxWidth = 25; // Maximum to prevent overly wide columns
    const padding = 3;   // Extra padding for dropdown arrow and margins
    
    return Math.max(minWidth, Math.min(maxWidth, baseWidth + padding));
  }

  // Generate column definitions: info columns, then for each date two columns (Legend, Hours), then Total Hours
  const dynamicWidths = calculateEmployeeColumnWidths(employees);
  const infoColDefs = [
    { min: 1, max: 1, width: dynamicWidths.name, bestFit: true, customWidth: true },     // Employee Name (dynamic)
    { min: 2, max: 2, width: 10, bestFit: true, customWidth: true },                    // ID
    { min: 3, max: 3, width: dynamicWidths.title, bestFit: true, customWidth: true },  // Job Title (dynamic)
    { min: 4, max: 4, width: 17, bestFit: true, customWidth: true },                   // Shift Start
    { min: 5, max: 5, width: 17, bestFit: true, customWidth: true },                   // First Break
    { min: 6, max: 6, width: 17, bestFit: true, customWidth: true },                   // Lunch Break
    { min: 7, max: 7, width: 18, bestFit: true, customWidth: true },                   // Second Break
    { min: 8, max: 8, width: 15, bestFit: true, customWidth: true }                    // Shift End
  ];
  
  const legendColumnWidth = calculateLegendColumnWidth(legends);
  const dateColDefs = dates.flatMap((_, i) => [
    { min: infoHeaders.length + i * 2 + 1, max: infoHeaders.length + i * 2 + 1, width: legendColumnWidth, bestFit: true, customWidth: true }, // Legend - dynamic width
    { min: infoHeaders.length + i * 2 + 2, max: infoHeaders.length + i * 2 + 2, width: 10, bestFit: true, customWidth: true }  // Hours
  ]);
  const totalColDef = [{ min: infoHeaders.length + dates.length * 2 + 1, max: infoHeaders.length + dates.length * 2 + 1, width: 14, bestFit: true, customWidth: true }];
  const columnDefs = [...infoColDefs, ...dateColDefs, ...totalColDef];
  
  const colsXML = generateColumnsXML(columnDefs, {
    enableAutoWidth: false,
    includeSheetFormat: true
  });

  // Header row
  const headerRowXML = generateRowXML(1, 'subtitle', {
    styleId: STYLE_IDS.TABLE_HEADER,
    enableAutoHeight: false // Fixed height for header
  });
  let headerXML = headerRowXML;
  let colIndex = 0;
  infoHeaders.forEach((header) => {
    const colLetter = getExcelColLetter(colIndex++);
    headerXML += `<c r="${colLetter}1" s="${STYLE_IDS.TABLE_HEADER}" t="inlineStr"><is><t>${escapeXml(header)}</t></is></c>`;
  });
  dates.forEach((dateStr) => {
    const [year, month, day] = dateStr.split('-').map(Number);
    const dateObj = new Date(year, month - 1, day);
    const weekday = dateObj.toLocaleDateString('en-US', { weekday: 'short' });
    const dateLabel = `${weekday} ${month}/${day}`;
    // Legend column
    const legendColLetter = getExcelColLetter(colIndex++);
    headerXML += `<c r="${legendColLetter}1" s="${STYLE_IDS.TABLE_HEADER}" t="inlineStr"><is><t>${escapeXml(dateLabel)}</t></is></c>`;
    // Hours column
    const hoursColLetter = getExcelColLetter(colIndex++);
    headerXML += `<c r="${hoursColLetter}1" s="${STYLE_IDS.TABLE_FORMULA}" t="inlineStr"><is><t>Hours</t></is></c>`;
  });
  // Total Hours column
  const totalColLetter = getExcelColLetter(colIndex);
  headerXML += `<c r="${totalColLetter}1" s="${STYLE_IDS.TABLE_FORMULA}" t="inlineStr"><is><t>Total Hours</t></is></c>`;
  headerXML += '</row>';

  // Add separator row between headers and employee data (merged across all columns)
  const separatorRowXML = generateRowXML(2, 'data', {
    styleId: STYLE_IDS.TABLE_ALT_ROW, // Use alternate row style for visual separation
    enableAutoHeight: false // Fixed height for separator
  });
  let separatorXML = separatorRowXML;
  // Add single cell that will be merged across all columns
  separatorXML += `<c r="A2" s="${STYLE_IDS.TABLE_ALT_ROW}"/>`;
  separatorXML += '</row>';

  // Hidden legend data no longer needed - using cross-sheet references to Legends sheet
  let hiddenLegendXML = '';

  // Employee rows (starting at row 3 due to separator row)
  let employeeRowsXML = '';
  employees.forEach((employee, empIndex) => {
    const rowNum = empIndex + 3;
    const rowStartXML = generateRowXML(rowNum, 'data', {
      styleId: STYLE_IDS.DATA_CELL,
      enableAutoHeight: true // Always auto height for data rows
    });
    let rowXML = rowStartXML;
    // Employee info columns
    const rowData = [
      employee.name,
      employee.id || '',
      employee.title,
      formatTimeForExcel(employee.shifts.start),
      formatTimeForExcel(employee.shifts.firstBreak),
      formatTimeForExcel(employee.shifts.lunch),
      formatTimeForExcel(employee.shifts.secondBreak),
      formatTimeForExcel(employee.shifts.end)
    ];
    let colIdx = 0;
    rowData.forEach((data) => {
      const colLetter = getExcelColLetter(colIdx++);
      let styleId = colIdx === 1 ? STYLE_IDS.TABLE_EMPLOYEE : STYLE_IDS.TABLE_DATA;
      rowXML += `<c r="${colLetter}${rowNum}" s="${styleId}" t="inlineStr"><is><t>${escapeXml(data)}</t></is></c>`;
    });
    // For each date: Legend (dropdown, required) and Hours (manual input, yellow)
    const hoursColLetters = [];
    dates.forEach((_, d) => {
      const legendColLetter = getExcelColLetter(colIdx++);
      rowXML += `<c r="${legendColLetter}${rowNum}" s="${STYLE_IDS.TABLE_INPUT}"/>`;
      const hoursColLetter = getExcelColLetter(colIdx++);
      rowXML += `<c r="${hoursColLetter}${rowNum}" s="${STYLE_IDS.TABLE_FORMULA}"/>`;
      hoursColLetters.push(`${hoursColLetter}${rowNum}`);
    });
    // Total Hours column (sum all hours columns for this row)
    const totalColLetter = getExcelColLetter(colIdx);
    rowXML += `<c r="${totalColLetter}${rowNum}" s="${STYLE_IDS.TABLE_FORMULA}" t="str"><f>SUM(${hoursColLetters.join(",")})</f></c>`;
    rowXML += '</row>';
    employeeRowsXML += rowXML;
  });

  // After generating employeeRowsXML, build data validation rules for all legend columns
  const dataValidationRules = [];
  if (legends.length > 0 && employees.length > 0) {
    const firstDataRow = 3; // Employee data now starts at row 3
    const lastDataRow = Math.max(employees.length + 2, 100); // Apply to sufficient rows for future use
    let legendColIndices = [];
    let colIdx = infoHeaders.length;
    for (let d = 0; d < dates.length; d++) {
      legendColIndices.push(colIdx);
      colIdx += 2; // Skip hours col
    }
    
    // Reference the legend labels from the Legends sheet (column A, starting from row 3)
    const lastLegendRow = legends.length + 2; // +2 because legends start at row 3 (1=title, 2=header)
    const legendRange = `Legends!$A$3:$A$${lastLegendRow}`;
    
    legendColIndices.forEach(colIndex => {
      const colLetter = getExcelColLetter(colIndex);
      const sqref = `${colLetter}${firstDataRow}:${colLetter}${lastDataRow}`;
      dataValidationRules.push(new DataValidationRule(sqref, legendRange, { 
        allowBlank: false, 
        showDropDown: true, // This will be converted to "0" in XML (showing dropdown)
        type: 'list'
      }));
    });
  }
  const dataValidationsXML = generateDataValidationsXML(dataValidationRules);

  // Use ConditionalFormattingRule class with DXF IDs like the calendar builder
  const tempSheet = new ExcelSheet(sheetName);
  
  if (legends.length > 0) {
    const firstDataRow = 3; // Employee data now starts at row 3
    const lastDataRow = Math.max(employees.length + 2, 10); // Keep range small for testing
    let legendColIndices = [];
    let colIdx = infoHeaders.length;
    for (let d = 0; d < dates.length; d++) {
      legendColIndices.push(colIdx);
      colIdx += 2; // Skip hours col
    }
    
    // Create conditional formatting for each legend using the ConditionalFormattingRule class
    legends.forEach((legend, legendIndex) => {
      // Ensure color is in proper Excel RGB format (FFRRGGBB)
      let color;
      if (legend.color.startsWith('#')) {
        // Convert #RRGGBB to FFRRGGBB
        color = 'FF' + legend.color.slice(1).toUpperCase();
      } else if (legend.color.startsWith('FF')) {
        // Already in Excel format
        color = legend.color.toUpperCase();
      } else {
        // Assume it's RRGGBB, add FF prefix
        color = 'FF' + legend.color.toUpperCase();
      }
      
      legendColIndices.forEach(colIndex => {
        const colLetter = getExcelColLetter(colIndex);
        const range = `${colLetter}${firstDataRow}:${colLetter}${lastDataRow}`;
        
        // Use simple cellIs comparison with DXF ID reference
        const formula = `"${legend.label}"`;
        const rule = new ConditionalFormattingRule(range, formula, color, legendIndex + 1, false, legendIndex);
        tempSheet.addConditionalFormatting(rule);
      });
    });
  }

  // Extract conditional formatting XML from the sheet
  let conditionalFormattingXML = '';
  if (tempSheet.conditionalFormatting.length > 0) {
    // Group conditional formatting rules by range
    const rangeGroups = {};
    tempSheet.conditionalFormatting.forEach(cf => {
      if (!rangeGroups[cf.sqref]) rangeGroups[cf.sqref] = [];
      rangeGroups[cf.sqref].push(cf);
    });
    
    conditionalFormattingXML = Object.keys(rangeGroups).map(sqref => 
      `<conditionalFormatting sqref="${sqref}">${rangeGroups[sqref].map(cf => cf.toXML()).join('')}</conditionalFormatting>`
    ).join('');
  }

  // Create merged cells XML for separator row
  const totalColumnsForMerge = infoHeaders.length + dates.length * 2 + 1; // Info + (Legend+Hours per date) + Total
  const lastColLetter = getExcelColLetter(totalColumnsForMerge - 1);
  const mergeCellsXML = `<mergeCells count="1"><mergeCell ref="A2:${lastColLetter}2"/></mergeCells>`;

  // Freeze pane: freeze first column and header/separator rows
  const freezePaneXML = `<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>`;

  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${freezePaneXML}
  ${colsXML}
  <sheetData>
    ${headerXML}
    ${separatorXML}
    ${employeeRowsXML}
    ${hiddenLegendXML}
  </sheetData>
  ${mergeCellsXML}
  ${conditionalFormattingXML}
  ${dataValidationsXML}
</worksheet>`;
}


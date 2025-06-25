// generators/referenceSheet.js
// Generates the employee reference sheet XML
import { escapeXml } from '../../core/index.js';
import { STYLE_IDS, COLUMN_WIDTHS } from '../../../presentation/index.js';
import { createAttendanceColumns, generateColumnsXML, generateRowXML, ROW_HEIGHTS, calculateEmployeeColumnWidths, calculateReferenceSheetWidths } from '../../../presentation/sizing/excelSizing.js';
import { formatTimeForExcel, calculateShiftHours } from '../../../shared/utils/timeUtils.js';

export function buildReferenceSheet(employees, legends = [], sheetName = "Shift Tracker") {
  const headers = ['Employee', 'Shift Hours', 'Daily Total', 'Email'];
  
  // Calculate dynamic widths based on actual data
  const dynamicWidths = calculateReferenceSheetWidths(employees, legends);
  
  // Generate column definitions with proper widths
  const referenceColumns = [
    { min: 1, max: 1, width: dynamicWidths.employee },     // Employee (dynamic)
    { min: 2, max: 2, width: dynamicWidths.shiftHours },   // Shift Hours (dynamic) 
    { min: 3, max: 3, width: dynamicWidths.dailyTotal },   // Daily Total (dynamic)
    { min: 4, max: 4, width: dynamicWidths.email },        // Email (dynamic)
    { min: 5, max: 5, width: dynamicWidths.spacer },       // Spacer (3)
    { min: 6, max: 6, width: dynamicWidths.legend },       // Legend (dynamic)
    { min: 7, max: 7, width: dynamicWidths.usage },        // Usage (dynamic)
    { min: 8, max: 8, width: dynamicWidths.percentage }    // % (dynamic)
  ];
  
  const columnDefs = referenceColumns.map(col => ({
    ...col,
    bestFit: true,
    customWidth: true
  }));
  
  const colsXML = generateColumnsXML(columnDefs, {
    enableAutoWidth: false,
    includeSheetFormat: true
  });

  // Row 1: Title row with both employee and legend titles
  const titleRowXML = generateRowXML(1, 'title', {
    styleId: STYLE_IDS.TABLE_TITLE,
    enableAutoHeight: false
  });
  
  let titleXML = titleRowXML;
  // Employee title (A1-D1)
  titleXML += `<c r="A1" s="${STYLE_IDS.TABLE_TITLE}" t="inlineStr"><is><t>Employee Shift Reference</t></is></c>`;
  titleXML += `<c r="B1" s="${STYLE_IDS.TABLE_TITLE}"/>`;
  titleXML += `<c r="C1" s="${STYLE_IDS.TABLE_TITLE}"/>`;
  titleXML += `<c r="D1" s="${STYLE_IDS.TABLE_TITLE}"/>`;
  
  // Legend Usage title (F1-H1) - merged across FGH if legends exist
  if (legends && legends.length > 0) {
    titleXML += `<c r="F1" s="${STYLE_IDS.TABLE_TITLE}" t="inlineStr"><is><t>Legend Usage Analytics</t></is></c>`;
    titleXML += `<c r="G1" s="${STYLE_IDS.TABLE_TITLE}"/>`;
    titleXML += `<c r="H1" s="${STYLE_IDS.TABLE_TITLE}"/>`;
  }
  titleXML += '</row>';

  // Row 2: Spacer row
  const spacerRowXML = generateRowXML(2, 'spacer', {
    enableAutoHeight: false
  });
  const spacerXML = `${spacerRowXML}</row>`;

  // Row 3: Headers for both employee data AND legend usage
  const headerRowXML = generateRowXML(3, 'header', {
    styleId: STYLE_IDS.TABLE_HEADER,
    enableAutoHeight: false
  });
  let headerXML = headerRowXML;
  
  // Employee headers (A3-D3)
  headers.forEach((header, index) => {
    const colLetter = String.fromCharCode(65 + index);
    headerXML += `<c r="${colLetter}3" s="${STYLE_IDS.TABLE_HEADER}" t="inlineStr"><is><t>${escapeXml(header)}</t></is></c>`;
  });
  
  // Legend usage headers (F3-H3) - only if legends exist
  if (legends && legends.length > 0) {
    headerXML += `<c r="F3" s="${STYLE_IDS.TABLE_HEADER}" t="inlineStr"><is><t>Legend</t></is></c>`;
    headerXML += `<c r="G3" s="${STYLE_IDS.TABLE_HEADER}" t="inlineStr"><is><t>Usage</t></is></c>`;
    headerXML += `<c r="H3" s="${STYLE_IDS.TABLE_HEADER}" t="inlineStr"><is><t>%</t></is></c>`;
  }
  headerXML += '</row>';

  // Employee data rows starting from row 4, with legend usage in the same rows
  let employeeRowsXML = '';
  const maxRows = Math.max(employees.length, legends?.length || 0);
  
  // Calculate the correct range for attendance legend data
  // Employee data starts at row 3 in attendance tracker
  const firstDataRow = 3;
  const lastDataRow = employees.length + 2;
  
  for (let i = 0; i < maxRows; i++) {
    const rowNum = i + 4;
    
    const rowStartXML = generateRowXML(rowNum, 'content', {
      enableAutoHeight: true
    });
    let rowXML = rowStartXML;
    
    // Add employee data if this employee exists
    if (i < employees.length) {
      const employee = employees[i];
      const shiftHours = `${formatTimeForExcel(employee.shifts.start)} - ${formatTimeForExcel(employee.shifts.end)}`;
      const dailyTotal = calculateShiftHours(employee.shifts) + ' hours';
      
      const rowData = [
        employee.name,
        shiftHours,
        dailyTotal,
        employee.email || ''
      ];

      rowData.forEach((data, colIndex) => {
        const colLetter = String.fromCharCode(65 + colIndex);
        const styleId = colIndex === 0 ? STYLE_IDS.TABLE_EMPLOYEE : STYLE_IDS.TABLE_DATA;
        
        rowXML += `<c r="${colLetter}${rowNum}" s="${styleId}" t="inlineStr"><is><t>${escapeXml(data)}</t></is></c>`;
      });
    }
    
    // Add legend usage data with DYNAMIC FORMULAS if this legend exists
    if (i < legends.length) {
      const legend = legends[i];
      const legendLabel = legend.label.replace(/"/g, '""'); // Escape quotes in legend label
      
      // Simple approach: count in the entire attendance data area (columns I onwards)
      // This covers all legend columns regardless of exact structure
      // Use the correct sheet name (with quotes to handle spaces in sheet names)
      const attendanceRange = `'${sheetName}'!I${firstDataRow}:BZ${lastDataRow}`;
      
      // Dynamic formula to count occurrences of this legend
      const countFormula = `=COUNTIF(${attendanceRange},"${legendLabel}")`;
      
      // Dynamic formula to calculate percentage (based on total employees * estimated days)
      const estimatedTotalEntries = employees.length * 30; // 30 days estimate
      const percentageFormula = `=IF(G${rowNum}=0,0,ROUND(G${rowNum}/${estimatedTotalEntries}*100,0))`;
      
      // Legend name (static text)
      rowXML += `<c r="F${rowNum}" s="${STYLE_IDS.TABLE_DATA}" t="inlineStr"><is><t>${escapeXml(legend.label)}</t></is></c>`;
      
      // Usage count (dynamic formula)
      rowXML += `<c r="G${rowNum}" s="${STYLE_IDS.TABLE_DATA}"><f>${countFormula}</f><v>0</v></c>`;
      
      // Percentage (dynamic formula)  
      rowXML += `<c r="H${rowNum}" s="${STYLE_IDS.TABLE_DATA}"><f>${percentageFormula}</f><v>0</v></c>`;
    }
    
    rowXML += '</row>';
    employeeRowsXML += rowXML;
  }

  // Merge title across employee columns (A1:D1) and legend columns (F1:H1) if legends exist
  let mergesXML = '';
  if (legends && legends.length > 0) {
    mergesXML = `<mergeCells count="2"><mergeCell ref="A1:D1"/><mergeCell ref="F1:H1"/></mergeCells>`;
  } else {
    mergesXML = `<mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>`;
  }

  // Simple worksheet without any drawing/chart references
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${colsXML}
  <sheetData>
    ${titleXML}
    ${spacerXML}
    ${headerXML}
    ${employeeRowsXML}
  </sheetData>
  ${mergesXML}
</worksheet>`;
}
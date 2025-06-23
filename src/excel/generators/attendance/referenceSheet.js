// generators/referenceSheet.js
// Generates the employee reference sheet XML
import { escapeXml } from '../../core/index.js';
import { STYLE_IDS, COLUMN_WIDTHS } from '../../../presentation/index.js';
import { createAttendanceColumns, generateColumnsXML, generateRowXML, ROW_HEIGHTS } from '../../../presentation/sizing/excelSizing.js';

export function buildReferenceSheet(employees) {
  const headers = ['Employee', 'Shift Hours', 'Daily Total', 'Email'];
  // Generate column definitions using enhanced auto-sizing with universal constants
  const referenceColumns = [
    { min: 1, max: 1, width: COLUMN_WIDTHS.name },        // Employee (20)
    { min: 2, max: 2, width: COLUMN_WIDTHS.description }, // Shift Hours (25) 
    { min: 3, max: 3, width: COLUMN_WIDTHS.medium },      // Daily Total (15)
    { min: 4, max: 4, width: COLUMN_WIDTHS.description + 5 }  // Email (30)
  ];
  
  const columnDefs = referenceColumns.map(col => ({
    ...col,    bestFit: true,
    customWidth: true
  }));
  
  const colsXML = generateColumnsXML(columnDefs, {
    enableAutoWidth: false,
    includeSheetFormat: true
  });
    // Title row with fixed height for consistency
  const titleRowXML = generateRowXML(1, 'title', {
    styleId: STYLE_IDS.TABLE_TITLE,
    enableAutoHeight: false  // Use fixed height for title
  });
  const titleXML = `${titleRowXML}
    <c r="A1" s="${STYLE_IDS.TABLE_TITLE}" t="inlineStr"><is><t>Employee Shift Reference</t></is></c>
  </row>`;
  // Empty row for spacing with fixed height
  const spacerRowXML = generateRowXML(2, 'spacer', {
    enableAutoHeight: false  // Use fixed height for spacer
  });
  const spacerXML = `${spacerRowXML}</row>`;

  // Header row with fixed height for consistency
  const headerRowXML = generateRowXML(3, 'header', {
    styleId: STYLE_IDS.TABLE_HEADER,
    enableAutoHeight: false  // Use fixed height for headers
  });
  let headerXML = headerRowXML;
  headers.forEach((header, index) => {
    const colLetter = String.fromCharCode(65 + index);
    headerXML += `<c r="${colLetter}3" s="${STYLE_IDS.TABLE_HEADER}" t="inlineStr"><is><t>${escapeXml(header)}</t></is></c>`;
  });
  headerXML += '</row>';
  // Employee data rows with auto-height for content adaptation
  let employeeRowsXML = '';
  employees.forEach((employee, empIndex) => {
    const rowNum = empIndex + 4;
    
    const shiftHours = `${formatTimeForExcel(employee.shifts.start)} - ${formatTimeForExcel(employee.shifts.end)}`;
    const dailyTotal = calculateShiftHours(employee.shifts) + ' hours';
    
    const rowData = [
      employee.name,
      shiftHours,
      dailyTotal,
      employee.email || ''
    ];    // Use row generation with auto-height for employee data
    const rowStartXML = generateRowXML(rowNum, 'content', {
      enableAutoHeight: true  // Enable auto-height for employee data rows
    });
    let rowXML = rowStartXML;
      rowData.forEach((data, colIndex) => {
      const colLetter = String.fromCharCode(65 + colIndex);
      const styleId = colIndex === 0 ? STYLE_IDS.TABLE_EMPLOYEE : STYLE_IDS.TABLE_DATA; // Bold for employee name
      
      rowXML += `<c r="${colLetter}${rowNum}" s="${styleId}" t="inlineStr"><is><t>${escapeXml(data)}</t></is></c>`;
    });
    
    rowXML += '</row>';
    employeeRowsXML += rowXML;
  });

  // Merge title across all columns
  const mergesXML = `<mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>`;

  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
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

/**
 * Format time for Excel display
 */
function formatTimeForExcel(time) {
  if (!time) return '';
  
  const [hours, minutes] = time.split(':');
  const hour = parseInt(hours);
  const ampm = hour >= 12 ? 'PM' : 'AM';
  const displayHour = hour % 12 || 12;
  
  return `${displayHour}:${minutes} ${ampm}`;
}

/**
 * Calculate total shift hours
 */
function calculateShiftHours(shifts) {
  if (!shifts.start || !shifts.end) return '0';
  
  const startMinutes = timeToMinutes(shifts.start);
  const endMinutes = timeToMinutes(shifts.end);
  
  let totalMinutes = endMinutes - startMinutes;
  if (totalMinutes < 0) totalMinutes += 24 * 60; // Handle overnight shifts
  
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  
  return minutes > 0 ? `${hours}.${Math.round(minutes/60*10)}` : `${hours}`;
}

/**
 * Convert time string to minutes for comparison
 */
function timeToMinutes(timeString) {
  const [hours, minutes] = timeString.split(':').map(Number);
  return hours * 60 + minutes;
}
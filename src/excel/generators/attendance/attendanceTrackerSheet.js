// generators/attendanceTrackerSheet.js
// Generates the main shift tracker worksheet XML
import { escapeXml } from '../../core/index.js';
import { STYLE_IDS } from '../../../presentation/index.js';
import { createAttendanceColumns, generateColumnsXML, generateRowXML, ROW_HEIGHTS } from '../../../presentation/sizing/excelSizing.js';

// Debug: Track successful imports and module loading
console.log('✓ AttendanceShiftTracker imports loaded:', {
  escapeXml: typeof escapeXml,
  STYLE_IDS: typeof STYLE_IDS,
  createAttendanceColumns: typeof createAttendanceColumns,
  generateColumnsXML: typeof generateColumnsXML,
  generateRowXML: typeof generateRowXML,
  ROW_HEIGHTS: typeof ROW_HEIGHTS
});

export function buildShiftTrackerSheet(employees) {
  console.log('📊 AttendanceShiftTracker: Building shift tracker sheet...', {
    employeeCount: employees?.length || 0
  });
  const headers = [
    'Employee Name', 'ID', 'Job Title', 'Shift Start', 'First Break', 
    'Lunch Break', 'Second Break', 'Shift End', 'Monday', 'Tuesday', 
    'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday', 'Total Hours'
  ];  // Generate column definitions using auto-sizing
  const columnDefs = createAttendanceColumns();
  
  const colsXML = generateColumnsXML(columnDefs, {
    enableAutoWidth: false,
    includeSheetFormat: true
  });  // Generate header row with height control - use auto-height
  const headerRowXML = generateRowXML(1, 'subtitle', { 
    styleId: STYLE_IDS.TABLE_HEADER,
    enableAutoHeight: true  // Enable auto-height for better sizing
  });
  let headerXML = headerRowXML;
  headers.forEach((header, index) => {
    const colLetter = String.fromCharCode(65 + index);
    headerXML += `<c r="${colLetter}1" s="${STYLE_IDS.TABLE_HEADER}" t="inlineStr"><is><t>${escapeXml(header)}</t></is></c>`;
  });
  headerXML += '</row>';// Generate employee data rows with enhanced height control - use auto-height
  let employeeRowsXML = '';  employees.forEach((employee, empIndex) => {
    const rowNum = empIndex + 2;    const rowStartXML = generateRowXML(rowNum, 'data', {
      styleId: STYLE_IDS.DATA_CELL,
      enableAutoHeight: true  // Enable auto-height for employee rows
    });
    let rowXML = rowStartXML;
    
    // Employee data columns
    const rowData = [
      employee.name,
      employee.id || '',
      employee.title,
      formatTimeForExcel(employee.shifts.start),
      formatTimeForExcel(employee.shifts.firstBreak),
      formatTimeForExcel(employee.shifts.lunch),
      formatTimeForExcel(employee.shifts.secondBreak),
      formatTimeForExcel(employee.shifts.end),
      '', // Monday hours (user input)
      '', // Tuesday hours
      '', // Wednesday hours
      '', // Thursday hours
      '', // Friday hours
      '', // Saturday hours
      '', // Sunday hours
      '' // Total hours (will be formula)
    ];

    rowData.forEach((data, colIndex) => {
      const colLetter = String.fromCharCode(65 + colIndex);      let styleId = STYLE_IDS.TABLE_DATA; // Default data style
      
      // Employee name column - bold style
      if (colIndex === 0) {
        styleId = STYLE_IDS.TABLE_EMPLOYEE;
      }
      // Hours input columns - special style
      else if (colIndex >= 8 && colIndex <= 14) {
        styleId = STYLE_IDS.TABLE_INPUT;
      }
      // Total hours column - formula style
      else if (colIndex === 15) {
        styleId = STYLE_IDS.TABLE_FORMULA;
        // Add SUM formula for total hours
        rowXML += `<c r="${colLetter}${rowNum}" s="${styleId}" t="str"><f>SUM(I${rowNum}:O${rowNum})</f></c>`;
        return;
      }      if (data !== '') {
        rowXML += `<c r="${colLetter}${rowNum}" s="${styleId}" t="inlineStr"><is><t>${escapeXml(data)}</t></is></c>`;
      } else {
        rowXML += `<c r="${colLetter}${rowNum}" s="${styleId}"/>`;
      }
    });
    
    rowXML += '</row>';
    employeeRowsXML += rowXML;
  });  // Add instructions row with height control - use auto-height for text wrapping
  const instructionRowNum = employees.length + 4;
  const instructionRowStartXML = generateRowXML(instructionRowNum, 'instruction', {
    styleId: STYLE_IDS.INSTRUCTION_BULLET,
    enableAutoHeight: true  // Enable auto-height for instruction text wrapping
  });
  const instructionsXML = `${instructionRowStartXML}
    <c r="A${instructionRowNum}" s="${STYLE_IDS.INSTRUCTION_BULLET}" t="inlineStr"><is><t>Instructions: Enter daily hours worked in the Monday-Sunday columns. Use decimal format (e.g., 8.5 for 8 hours 30 minutes). Total hours will calculate automatically.</t></is></c>
  </row>`;

  // Merge instructions across all columns
  const mergesXML = `<mergeCells count="1"><mergeCell ref="A${instructionRowNum}:P${instructionRowNum}"/></mergeCells>`;

  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${colsXML}
  <sheetData>
    ${headerXML}
    ${employeeRowsXML}
    ${instructionsXML}
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
// excel/generators/calendar/calendarBuilderSheet.js
// Builds the main calendar worksheet using ExcelBuilder
import { ExcelSheet, ExcelCell, ExcelRow, ConditionalFormattingRule } from '../../core/index.js';
import { createCalendarColumns, applyAutoSizing, ROW_HEIGHTS } from '../../../presentation/sizing/excelSizing.js';
import { STYLE_IDS, getCustomStyleId } from '../../../presentation/index.js';

// Debug: Track successful imports
console.log('✓ CalendarSheet imports loaded:', {
  ExcelSheet: typeof ExcelSheet,
  ExcelCell: typeof ExcelCell,
  ExcelRow: typeof ExcelRow,
  ConditionalFormattingRule: typeof ConditionalFormattingRule,
  createCalendarColumns: typeof createCalendarColumns,
  applyAutoSizing: typeof applyAutoSizing,
  ROW_HEIGHTS: typeof ROW_HEIGHTS,
  STYLE_IDS: typeof STYLE_IDS,
  getCustomStyleId: typeof getCustomStyleId
});

export function buildCalendarSheetWithExcelBuilder(year, month, eventRows, includeDrawing, legendValues = null, customColors = null) {
  console.log('📅 CalendarSheet: Building calendar...', {
    year, month, eventRows, includeDrawing,
    legendValuesProvided: !!legendValues,
    customColorsProvided: !!customColors
  });
  
  const monthNames = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  const daysOfWeek = ["SUNDAY","MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY"];
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const startDay = new Date(year, month, 1).getDay();

  const defaultLegendValues = [
    "Meeting", "Workout", "Appointment", "Holiday", "Personal",
    "Work", "Travel", "Study", "Event"
  ];  const actualLegendValues = legendValues || defaultLegendValues.slice(0, eventRows);
  
  console.log('📊 CalendarSheet: Calendar data prepared:', {
    monthName: monthNames[month],
    daysInMonth,
    startDay,
    legendValues: actualLegendValues
  });  // Create calendar sheet with auto-sizing
  const sheet = new ExcelSheet("Calendar");
  const columnDefs = createCalendarColumns();
  
  // Apply auto-sizing for calendar
  applyAutoSizing(sheet, 'calendar', { enableAutoFilter: true });
  function colToIndex(col) {
    return col.charCodeAt(0) - 65 + 1;
  }
    // Use the universal auto-sizing row class with optional auto-height
  class UniqueExcelRow extends ExcelRow {
    constructor(r, heightType = 'content', enableAutoHeight = false) {
      super(r);
      this.cellMap = new Map();
      this.heightType = heightType;
      this.enableAutoHeight = enableAutoHeight;
      
      // Set height based on type or enable auto-height
      if (enableAutoHeight) {
        // Don't set custom height - let Excel auto-size
        this.customHeightSet = false;
      } else if (ROW_HEIGHTS[heightType]) {
        this.height = ROW_HEIGHTS[heightType];
        this.customHeightSet = true;
      }
    }
    addCell(cell) {
      this.cellMap.set(cell.col, cell);
    }
    toXML() {
      const sorted = Array.from(this.cellMap.values()).sort((a, b) => colToIndex(a.col) - colToIndex(b.col));
      this.cells = sorted;
      
      // Generate row XML with conditional height attributes
      let xml = `<row r="${this.rowNumber}"`;
      
      if (this.customHeightSet && !this.enableAutoHeight) {
        xml += ` ht="${this.height}" customHeight="1"`;
      }
      // For auto-height rows, don't add height attributes
      
      xml += '>';
      
      // Add cell XML
      for (const cell of this.cells) {
        xml += cell.toXML ? cell.toXML() : cell.toString();
      }
      
      xml += '</row>';
      return xml;
    }
  }
    const rowMap = new Map();
  function getRow(r, heightType = 'content', enableAutoHeight = false) {
    if (!rowMap.has(r)) rowMap.set(r, new UniqueExcelRow(r, heightType, enableAutoHeight));
    return rowMap.get(r);
  }

  let headerRow = 1;
  let calDaysRow = 2;  // Month header (A1:G1, merged) - Title height for prominence
  let monthRow = getRow(headerRow, 'title');
  monthRow.addCell(new ExcelCell('A', headerRow, monthNames[month].toUpperCase() + ' ' + year, {style: STYLE_IDS.CALENDAR_TITLE, align: 'center'}));  // Add empty cells with same style for proper merged cell borders
  for (let col = 1; col < 7; col++) {
    monthRow.addCell(new ExcelCell(String.fromCharCode(65 + col), headerRow, '', {style: STYLE_IDS.CALENDAR_TITLE}));
  }
  sheet.addMerge(`A${headerRow}:G${headerRow}`);
  
  // Legend header (I1:J1, merged) - Same height as month header
  monthRow.addCell(new ExcelCell('I', headerRow, 'Legend', {style: STYLE_IDS.CALENDAR_LEGEND}));
  monthRow.addCell(new ExcelCell('J', headerRow, '', {style: STYLE_IDS.CALENDAR_LEGEND}));
  sheet.addMerge(`I${headerRow}:J${headerRow}`);
  
  // Legend rows with custom colors and merged cells - Header height
  for (let l = 0; l < eventRows; l++) {
    let legendRow = getRow(headerRow + 1 + l, 'header');
    const legendValue = actualLegendValues[l] || `Category ${l + 1}`;    // Use the custom color style: getCustomStyleId gets the correct style for legend colors
    legendRow.addCell(new ExcelCell('I', headerRow + 1 + l, legendValue, {style: getCustomStyleId(l)}));
    legendRow.addCell(new ExcelCell('J', headerRow + 1 + l, '', {style: getCustomStyleId(l)}));
    // Merge the legend cells I and J to create unified colored legend entries
    sheet.addMerge(`I${headerRow + 1 + l}:J${headerRow + 1 + l}`);
  }
  // Day-of-week header (A2:G2) - Subtitle height for readability
  let dowRow = getRow(calDaysRow, 'subtitle');
  for (let d = 0; d < 7; d++) {
    dowRow.addCell(new ExcelCell(String.fromCharCode(65 + d), calDaysRow, daysOfWeek[d], {style: STYLE_IDS.CALENDAR_HEADER}));
  }
  // Calendar grid
  let calGridStartRow = 3;
  let currentRow = calGridStartRow;
  let day = 1;
  let firstWeek = true;  while (day <= daysInMonth) {
    let weekCols = [];
    // Date row - Data height for date numbers
    let weekRow = getRow(currentRow, 'data');    for (let dow = 0; dow < 7; dow++) {
      if ((firstWeek && dow < startDay) || day > daysInMonth) {
        weekRow.addCell(new ExcelCell(String.fromCharCode(65 + dow), currentRow, '', {style: STYLE_IDS.CALENDAR_DAY}));
      } else {
        weekRow.addCell(new ExcelCell(String.fromCharCode(65 + dow), currentRow, day, {type: 'n', style: STYLE_IDS.CALENDAR_DAY}));
        day++;
      }
    }
    firstWeek = false;    // Event rows - Event height for multiple events/text wrapping
    for (let er = 0; er < eventRows; er++) {
      let eventRow = getRow(currentRow + 1 + er, 'event');
      for (let dow = 0; dow < 7; dow++) {
        eventRow.addCell(new ExcelCell(String.fromCharCode(65 + dow), currentRow + 1 + er, '', {style: STYLE_IDS.CALENDAR_DAY}));
      }
    }
    currentRow += 1 + eventRows;
  }

  const allRows = Array.from(rowMap.keys()).sort((a, b) => a - b);
  for (const r of allRows) {
    sheet.addRow(rowMap.get(r));
  }
  const lastCalendarRow = Math.max(...allRows.filter(r => r >= calGridStartRow));
  const calendarEventRange = `A${calGridStartRow + 1}:G${lastCalendarRow}`;

  const defaultPalette = [
    "FFDC143C", "FF228B22", "FF1E90FF", "FFFFA500", "FF800080",
    "FFFFFF00", "FF00CED1", "FF8B4513", "FF4682B4"
  ];
  const palette = customColors || defaultPalette;

  // Add conditional formatting rules for each legend value
  if (eventRows > 0) {
    const legendRowStart = headerRow + 1;
      for (let l = 0; l < eventRows; l++) {
        const legendRow = legendRowStart + l;
        // Use the actual top-left cell of the range for the formula
        const formula = `UPPER(A${calGridStartRow + 1})=UPPER($I$${legendRow})`;
        // Use DXF ID that matches the legend color index
        const rule = new ConditionalFormattingRule(calendarEventRange, formula, palette[l], l + 1, true, l);
        sheet.addConditionalFormatting(rule);
      }
    

  }

  let xml = sheet.toXML();
  // Debug: check if conditional formatting XML is actually in the output
  if (xml.includes('<conditionalFormatting')) {
  }

  return xml;
}

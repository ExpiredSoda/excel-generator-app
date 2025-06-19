// generators/calendarSheet.js
// Builds the main calendar worksheet using ExcelBuilder
import { ExcelSheet, ExcelCell, ExcelRow, ConditionalFormattingRule } from '../core/excelCore.js';

export function buildCalendarSheetWithExcelBuilder(year, month, eventRows, includeDrawing, legendValues = null, customColors = null) {
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
  ];
  const actualLegendValues = legendValues || defaultLegendValues.slice(0, eventRows);

  const cols = [];
  for (let c = 1; c <= 7; c++) cols.push({min: c, max: c, width: 15});
  cols.push({min: 8, max: 8, width: 3}); // H spacer
  cols.push({min: 9, max: 9, width: 18}); // I
  cols.push({min: 10, max: 10, width: 7}); // J

  const sheet = new ExcelSheet("Calendar");
  sheet.setCols(cols);

  function colToIndex(col) {
    return col.charCodeAt(0) - 65 + 1;
  }

  class UniqueExcelRow extends ExcelRow {
    constructor(r) {
      super(r);
      this.cellMap = new Map();
    }
    addCell(cell) {
      this.cellMap.set(cell.col, cell);
    }
    toXML() {
      const sorted = Array.from(this.cellMap.values()).sort((a, b) => colToIndex(a.col) - colToIndex(b.col));
      this.cells = sorted;
      return super.toXML();
    }
  }

  const rowMap = new Map();
  function getRow(r) {
    if (!rowMap.has(r)) rowMap.set(r, new UniqueExcelRow(r));
    return rowMap.get(r);
  }

  let headerRow = 1;
  let calDaysRow = 2;

  // Month header (A1:G1, merged)
  let monthRow = getRow(headerRow);
  monthRow.addCell(new ExcelCell('A', headerRow, monthNames[month].toUpperCase() + ' ' + year, {style: 4, align: 'center'}));
  sheet.addMerge(`A${headerRow}:G${headerRow}`);
  
  // Legend header (I1:J1, merged)
  monthRow.addCell(new ExcelCell('I', headerRow, 'Legend', {style: 5}));
  sheet.addMerge(`I${headerRow}:J${headerRow}`);

  // Legend rows with custom colors and merged cells
  for (let l = 0; l < eventRows; l++) {
    let legendRow = getRow(headerRow + 1 + l);
    const legendValue = actualLegendValues[l] || `Category ${l + 1}`;
    // Use the custom color style: style index 10 + l uses the custom colors
    legendRow.addCell(new ExcelCell('I', headerRow + 1 + l, legendValue, {style: 10 + l}));
    legendRow.addCell(new ExcelCell('J', headerRow + 1 + l, '', {style: 10 + l}));
    // Merge the legend cells I and J to create unified colored legend entries
    sheet.addMerge(`I${headerRow + 1 + l}:J${headerRow + 1 + l}`);
  }

  // Day-of-week header (A2:G2)
  let dowRow = getRow(calDaysRow);
  for (let d = 0; d < 7; d++) {
    dowRow.addCell(new ExcelCell(String.fromCharCode(65 + d), calDaysRow, daysOfWeek[d], {style: 3}));
  }

  // Calendar grid
  let calGridStartRow = 3;
  let currentRow = calGridStartRow;
  let day = 1;
  let firstWeek = true;
  while (day <= daysInMonth) {
    let weekCols = [];
    let weekRow = getRow(currentRow);
    for (let dow = 0; dow < 7; dow++) {
      if ((firstWeek && dow < startDay) || day > daysInMonth) {
        weekRow.addCell(new ExcelCell(String.fromCharCode(65 + dow), currentRow, '', {style: 0}));
      } else {
        weekRow.addCell(new ExcelCell(String.fromCharCode(65 + dow), currentRow, day, {type: 'n', style: 1}));
        day++;
      }
    }
    firstWeek = false;
    for (let er = 0; er < eventRows; er++) {
      let eventRow = getRow(currentRow + 1 + er);
      for (let dow = 0; dow < 7; dow++) {
        eventRow.addCell(new ExcelCell(String.fromCharCode(65 + dow), currentRow + 1 + er, '', {style: 0}));
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

  console.log('Building calendar with custom colors:', customColors);
  console.log('Using legend values:', actualLegendValues);
  
  const defaultPalette = [
    "FFDC143C", "FF228B22", "FF1E90FF", "FFFFA500", "FF800080",
    "FFFFFF00", "FF00CED1", "FF8B4513", "FF4682B4"
  ];
  const palette = customColors || defaultPalette;
  
  console.log('Final palette for conditional formatting:', palette);

  // Add conditional formatting rules for each legend value
  if (eventRows > 0) {
    const legendRowStart = headerRow + 1;
    
    console.log('Adding conditional formatting rules:');
    console.log(`Calendar event range: ${calendarEventRange}`);
    console.log(`Legend starts at row: ${legendRowStart}`);
    console.log(`Event rows: ${eventRows}`);
    
    for (let l = 0; l < eventRows; l++) {
      const legendRow = legendRowStart + l;
      // Use expression type with UPPER function for case-insensitive matching
      const formula = `UPPER(A${calGridStartRow + 1})=UPPER($I$${legendRow})`;
      // Reference DXF by ID (0-based index)
      const rule = new ConditionalFormattingRule(calendarEventRange, formula, palette[l], l + 1, true, l);
      
      console.log(`Adding CF rule ${l + 1}:`);
      console.log(`  Range: ${calendarEventRange}`);
      console.log(`  Formula: ${formula}`);
      console.log(`  Color: ${palette[l]}`);
      console.log(`  DXF ID: ${l}`);
      console.log(`  LegendRow: ${legendRow}`);
      
      sheet.addConditionalFormatting(rule);
    }
    
    console.log(`Total CF rules: ${sheet.conditionalFormatting.length}`);
  }

  let xml = sheet.toXML();
  // Debug: check if conditional formatting XML is actually in the output
  if (xml.includes('<conditionalFormatting')) {
    console.log('✓ Conditional formatting XML found in sheet');
    const cfStart = xml.indexOf('<conditionalFormatting');
    const cfEnd = xml.indexOf('</conditionalFormatting>') + '</conditionalFormatting>'.length;
    console.log('CF XML section:', xml.substring(cfStart, cfEnd));
  } else {
    console.log('✗ Conditional formatting XML NOT found in sheet');
  }

  return xml;
}

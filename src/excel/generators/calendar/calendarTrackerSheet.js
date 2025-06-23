// generators/calendarTrackerSheet.js
// Generates the tracker sheet XML for Excel
import { escapeXml } from '../../core/index.js';
import { STYLE_IDS, getCustomStyleId } from '../../../presentation/index.js';
import { createTrackerColumns, generateColumnsXML, generateRowXML, ROW_HEIGHTS } from '../../../presentation/sizing/excelSizing.js';

// Debug: Track successful imports and module loading
console.log('✓ CalendarTracker imports loaded:', {
  escapeXml: typeof escapeXml,
  STYLE_IDS: typeof STYLE_IDS,
  createTrackerColumns: typeof createTrackerColumns,
  generateColumnsXML: typeof generateColumnsXML,
  generateRowXML: typeof generateRowXML,
  ROW_HEIGHTS: typeof ROW_HEIGHTS
});

export function getTrackerSheetXML(eventRows, legendValues = null) {
  console.log('📊 CalendarTracker: Generating tracker sheet...', {
    eventRows,
    legendValuesProvided: !!legendValues
  });
  const defaultLegendValues = [
    "Meeting", "Workout", "Appointment", "Holiday", "Personal",
    "Work", "Travel", "Study", "Event"
  ];
  const actualLegendValues = legendValues || defaultLegendValues.slice(0, eventRows);
  const columnDefs = createTrackerColumns();

  const colsXML = generateColumnsXML(columnDefs, {
    enableAutoWidth: false,
    includeSheetFormat: true
  });

  // Header row with fixed height for consistency
  const headerRowXML = generateRowXML(1, 'header', {
    enableAutoHeight: false  // Use fixed height for headers
  });
  let rows = `${headerRowXML}
    <c r="A1" t="inlineStr" s="${STYLE_IDS.TABLE_HEADER}"><is><t>Legend Value</t></is></c>
    <c r="B1" t="inlineStr" s="${STYLE_IDS.TABLE_HEADER}"><is><t>Count</t></is></c>
    <c r="C1" t="inlineStr" s="${STYLE_IDS.TABLE_HEADER}"><is><t>Description</t></is></c>
  </row>`;

  for (let i = 0; i < eventRows; i++) {
    const rowNum = i + 2;
    const legendValue = actualLegendValues[i] || `Category ${i + 1}`;
    const legendCellRef = `Calendar!I${i + 2}`;
    const countFormula = `COUNTIF(Calendar!A:G,Calendar!I${i + 2})`;
    const descriptionText = `Auto-counts "${escapeXml(legendValue)}" entries from Calendar sheet`;
    // Data rows with auto-height for content adaptation
    const dataRowXML = generateRowXML(rowNum, 'content', {
      enableAutoHeight: true  // Enable auto-height for data rows
    });
    rows += `${dataRowXML}
      <c r="A${rowNum}" t="str" s="${getCustomStyleId(i)}"><f>=${legendCellRef}</f></c>
      <c r="B${rowNum}" t="str" s="${STYLE_IDS.TABLE_FORMULA}"><f>=${countFormula}</f></c>
      <c r="C${rowNum}" t="inlineStr" s="${STYLE_IDS.TABLE_DATA}"><is><t>${descriptionText}</t></is></c>
    </row>`;
  }
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${colsXML}
  <sheetData>${rows}</sheetData>
</worksheet>`;
}

// All logic for getTrackerSheetXML matches script.js, including formulas and XML structure.
// No changes needed.

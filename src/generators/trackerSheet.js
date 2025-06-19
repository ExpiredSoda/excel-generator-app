// generators/trackerSheet.js
// Generates the tracker sheet XML for Excel
import { escapeXml } from '../core/excelCore.js';

export function getTrackerSheetXML(eventRows, legendValues = null) {
  const defaultLegendValues = [
    "Meeting", "Workout", "Appointment", "Holiday", "Personal",
    "Work", "Travel", "Study", "Event"
  ];
  const actualLegendValues = legendValues || defaultLegendValues.slice(0, eventRows);
  let rows = `<row r="1">
    <c r="A1" t="inlineStr" s="7"><is><t>Legend Value</t></is></c>
    <c r="B1" t="inlineStr" s="7"><is><t>Count</t></is></c>
    <c r="C1" t="inlineStr" s="7"><is><t>Description</t></is></c>
  </row>`;
  for (let i = 0; i < eventRows; i++) {
    const rowNum = i + 2;
    const legendValue = actualLegendValues[i] || `Category ${i + 1}`;
    const legendCellRef = `Calendar!I${i + 2}`;
    const countFormula = `COUNTIF(Calendar!A:G,Calendar!I${i + 2})`;
    const descriptionText = `Auto-counts "${escapeXml(legendValue)}" entries from Calendar sheet`;
    rows += `<row r="${rowNum}">
      <c r="A${rowNum}" t="str" s="8"><f>=${legendCellRef}</f></c>
      <c r="B${rowNum}" t="str" s="9"><f>=${countFormula}</f></c>
      <c r="C${rowNum}" t="inlineStr" s="8"><is><t>${descriptionText}</t></is></c>
    </row>`;
  }
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="20"/>
    <col min="2" max="2" width="10"/>
    <col min="3" max="3" width="50" bestFit="1"/>
  </cols>
  <sheetData>${rows}</sheetData>
</worksheet>`;
}

// All logic for getTrackerSheetXML matches script.js, including formulas and XML structure.
// No changes needed.

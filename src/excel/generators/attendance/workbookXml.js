// generators/attendance/workbookXml.js
// Generates workbook structure XML for shift tracker (workbook only)
import { escapeXml } from '../../core/index.js';

// Debug logs removed

export function getShiftTrackerWorkbookXML(shiftTrackerSheetName = "Shift Tracker") {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Instructions" sheetId="1" r:id="rId1">
      <tabColor rgb="FF7030A0"/>
    </sheet>
    <sheet name="${escapeXml(shiftTrackerSheetName)}" sheetId="2" r:id="rId2">
      <tabColor rgb="FF20B388"/>
    </sheet>
    <sheet name="Quick Reference" sheetId="3" r:id="rId3">
      <tabColor rgb="FF4472C4"/>
    </sheet>
    <sheet name="Legends" sheetId="4" r:id="rId4">
      <tabColor rgb="FFFFC000"/>
    </sheet>
  </sheets>
</workbook>`;
}

export function getShiftTrackerWorkbookRelsXML() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

/**
 * Generate relationships XML for the Quick Reference sheet (sheet3) with drawing
 * @returns {string} Sheet relationships XML
 */
export function getSheetDrawingRelsXML() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`;
}

/**
 * Generate relationships XML for the drawing with chart reference
 * @returns {string} Drawing relationships XML
 */
export function getDrawingChartRelsXML() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`;
}

// Content types and main relationships moved to contentTypesXml.js for better organization
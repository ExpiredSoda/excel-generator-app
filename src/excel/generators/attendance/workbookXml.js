// generators/attendance/workbookXml.js
// Generates workbook structure XML for shift tracker (workbook only)

// Debug logs removed

export function getShiftTrackerWorkbookXML() {
  // Debug logs removed
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Shift Tracker" sheetId="1" r:id="rId1">
      <tabColor rgb="FF20B388"/>
    </sheet>
    <sheet name="Quick Reference" sheetId="2" r:id="rId2">
      <tabColor rgb="FF4472C4"/>
    </sheet>
    <sheet name="Instructions" sheetId="3" r:id="rId3">
      <tabColor rgb="FF7030A0"/>
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
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

// Content types and main relationships moved to contentTypesXml.js for better organization
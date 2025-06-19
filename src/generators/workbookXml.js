// generators/workbookXml.js
// Generates workbook.xml and workbook relationships XML

export function getWorkbookXML(includeTracker) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Instructions" sheetId="1" r:id="rId1"/>
    <sheet name="Calendar" sheetId="2" r:id="rId2">
      <tabColor rgb="FF00B050"/>
    </sheet>
    ${includeTracker ? `<sheet name="Tracker" sheetId="3" r:id="rId3"><tabColor rgb="FF7030A0"/></sheet>` : ''}
  </sheets>
</workbook>`;
}

export function getWorkbookRelsXML(includeTracker) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  ${includeTracker ? `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>` : ''}
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

// All logic for getWorkbookXML and getWorkbookRelsXML matches script.js, including tracker support and XML structure.
// No changes needed.

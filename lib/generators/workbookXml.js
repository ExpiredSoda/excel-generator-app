// generators/workbookXml.js
// Generates workbook.xml and workbook relationships XML

export function getWorkbookXML(includeTracker = false) {
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Calendar" sheetId="1" r:id="rId1"/>`;

  if (includeTracker) {
    xml += `
    <sheet name="Tracker" sheetId="2" r:id="rId2"/>`;
  }

  xml += `
  </sheets>
</workbook>`;
  
  return xml;
}

export function getWorkbookRelsXML(includeTracker = false) {
  let xml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>`;

  if (includeTracker) {
    xml += `
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>`;
  }
  
  xml += `
  <Relationship Id="rId${includeTracker ? '3' : '2'}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  
  return xml;
}

// All logic for getWorkbookXML and getWorkbookRelsXML matches script.js, including tracker support and XML structure.
// No changes needed.

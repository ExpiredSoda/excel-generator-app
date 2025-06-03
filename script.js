// =========================
// ExcelBuilder Library
// =========================
// Provides ExcelCell, ExcelRow, ExcelSheet, and ExcelBuilder classes for building Excel XML.
// --------------------------------------------------
class ExcelCell {
  constructor(col, row, value = '', opts = {}) {
    this.col = col; // e.g. 'A'
    this.row = row; // e.g. 1
    this.value = value;
    this.type = opts.type || 'inlineStr';
    this.style = opts.style || 0;
    this.mergeAcross = opts.mergeAcross || 0;
  }
  get ref() {
    return `${this.col}${this.row}`;
  }
  toXML() {
    let attrs = `r='${this.ref}'`;
    if (this.type) attrs += ` t='${this.type}'`;
    if (this.style) attrs += ` s='${this.style}'`;
    let val = this.value !== '' ? `<is><t>${this.value}</t></is>` : '';
    return `<c ${attrs}>${val}</c>`;
  }
}

class ExcelRow {
  constructor(r) {
    this.r = r;
    this.cells = [];
  }
  addCell(cell) {
    this.cells.push(cell);
  }
  toXML() {
    return `<row r='${this.r}'>${this.cells.map(c => c.toXML()).join('')}</row>`;
  }
}

class ExcelSheet {
  constructor(name) {
    this.name = name;
    this.rows = [];
    this.merges = [];
    this.cols = [];
  }
  addRow(row) {
    this.rows.push(row);
  }
  addMerge(ref) {
    this.merges.push(ref);
  }
  setCols(colDefs) {
    this.cols = colDefs;
  }
  toXML() {
    let colsXML = this.cols.length ? `<cols>${this.cols.map(c => `<col min='${c.min}' max='${c.max}' width='${c.width}'/>`).join('')}</cols>` : '';
    let mergesXML = this.merges.length ? `<mergeCells count='${this.merges.length}'>${this.merges.map(ref => `<mergeCell ref='${ref}'/>`).join('')}</mergeCells>` : '';
    // Add xmlns:r for relationships (needed for <drawing r:id=...>)
    return `<?xml version='1.0' encoding='UTF-8'?>\n<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>\n  ${colsXML}\n  <sheetData>\n    ${this.rows.map(r => r.toXML()).join('\n    ')}\n  </sheetData>\n  ${mergesXML}\n</worksheet>`;
  }
}

class ExcelBuilder {
  constructor() {
    this.sheets = [];
    this.styles = null;
  }
  addSheet(sheet) {
    this.sheets.push(sheet);
  }
  setStyles(stylesXML) {
    this.styles = stylesXML;
  }
  // Example: getSheetXML(0) returns XML for first sheet
  getSheetXML(idx) {
    return this.sheets[idx].toXML();
  }
  // Extend: add workbook.xml, rels, etc. as needed
}
// =========================
// End ExcelBuilder Library
// =========================

// =========================
// ZIP Writer Utilities
// =========================
// Functions for creating a minimal ZIP archive (uncompressed, no CRC).
// --------------------------------------------------

// Converts a string to an array of bytes (UTF-16, basic for ASCII)
function stringToBytes(str) {
  return Array.from(str, c => c.charCodeAt(0));
}

// Converts a number to a little-endian byte array of given length
function toBytesLE(num, len) {
  const arr = [];
  for (let i = 0; i < len; i++) {
    arr.push(num & 0xff);
    num >>= 8;
  }
  return arr;
}

// Minimal ZIP writer: MULTIPLE files (uncompressed, no CRC)
function createZip(files) {
  // files = [{ name: "filename", content: "string content" }]
  let offset = 0;
  let allData = [];
  let centralDir = [];
  files.forEach(file => {
    const fileBytes = stringToBytes(file.content);
    const fileLen = fileBytes.length;
    const filenameBytes = stringToBytes(file.name);

    // Local file header
    const localHeader = [
      0x50,0x4b,0x03,0x04, // Local file header signature
      0x14,0x00,           // Version needed to extract
      0x00,0x00,           // General purpose bit flag
      0x00,0x00,           // Compression method (store)
      0x00,0x00,           // Last mod time
      0x00,0x00,           // Last mod date
      0x00,0x00,0x00,0x00, // CRC32 (set to zero for now)
      ...toBytesLE(fileLen, 4), // Compressed size
      ...toBytesLE(fileLen, 4), // Uncompressed size
      ...toBytesLE(filenameBytes.length, 2), // File name length
      0x00,0x00            // Extra field length
    ];
    const local = [...localHeader, ...filenameBytes, ...fileBytes];
    allData.push(...local);

    // Central directory entry (references offset)
    const central = [
      0x50,0x4b,0x01,0x02, // Central directory file header signature
      0x14,0x00,           // Version made by
      0x14,0x00,           // Version needed to extract
      0x00,0x00,           // General purpose bit flag
      0x00,0x00,           // Compression method (store)
      0x00,0x00,           // Last mod time
      0x00,0x00,           // Last mod date
      0x00,0x00,0x00,0x00, // CRC32
      ...toBytesLE(fileLen, 4), // Compressed size
      ...toBytesLE(fileLen, 4), // Uncompressed size
      ...toBytesLE(filenameBytes.length, 2), // File name length
      0x00,0x00,           // Extra field length
      0x00,0x00,           // File comment length
      0x00,0x00,           // Disk number start
      0x00,0x00,           // Internal file attributes
      0x00,0x00,0x00,0x00, // External file attributes
      ...toBytesLE(offset, 4) // Relative offset of local header
    ];
    centralDir.push(...central, ...filenameBytes);
    offset += local.length;
  });

  // End of central directory record
  const centralDirLen = centralDir.length;
  const allDataLen = allData.length;
  const endCentral = [
    0x50,0x4b,0x05,0x06, // End of central dir signature
    0x00,0x00,           // Number of this disk
    0x00,0x00,           // Number of the disk with the start of the central directory
    ...toBytesLE(files.length, 2), // Number of entries on this disk
    ...toBytesLE(files.length, 2), // Total number of entries
    ...toBytesLE(centralDirLen, 4), // Size of the central directory
    ...toBytesLE(allDataLen, 4),    // Offset of central directory
    0x00,0x00            // Comment length
  ];

  return new Uint8Array([...allData, ...centralDir, ...endCentral]);
}
// =========================
// End ZIP Writer Utilities
// =========================

// =========================
// Excel XML Generators
// =========================
// Functions to generate required XML files for Excel structure.
// --------------------------------------------------

// Required for Excel file structure
function getContentTypesXML(includeTracker) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  ${includeTracker ? `<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` : ''}
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`;
}

function getRelsXML() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

function getSheet1InstructionsXML() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols><col min="2" max="8" width="24"/></cols>
  <sheetData>
    <row r="1"><c r="B1" t="inlineStr" s="1"><is><t>How to Use This Calendar Workbook</t></is></c></row>
    <row r="2"><c r="B2" t="inlineStr"><is><t>This Excel workbook was generated by Free Excel Generators.</t></is></c></row>
    <row r="3"><c r="B3" t="inlineStr"><is><t>- The Calendar sheet lets you organize events by day. Each day has up to 9 slots for event labels.</t></is></c></row>
    <row r="4"><c r="B4" t="inlineStr"><is><t>- Color-code your events according to the legend.</t></is></c></row>
    <row r="5"><c r="B5" t="inlineStr"><is><t>- You can add or edit events directly in the Calendar sheet. Use cell fill color to match the legend, or simply type event names.</t></is></c></row>
    <row r="6"><c r="B6" t="inlineStr"><is><t>- [Your website link here]</t></is></c></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="B1:H1"/></mergeCells>
</worksheet>`;
}

// =========================
// Calendar Sheet Builder
// =========================
// Builds the main calendar worksheet using ExcelBuilder.
// --------------------------------------------------

// Build the calendar sheet using the ExcelBuilder library
function buildCalendarSheetWithExcelBuilder(year, month, eventRows, includeDrawing) {
  const monthNames = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  const daysOfWeek = ["SUNDAY","MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY"];
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const startDay = new Date(year, month, 1).getDay();

  // Setup columns (A-G calendar, I/J legend)
  const cols = [];
  for (let c = 1; c <= 7; c++) cols.push({min: c, max: c, width: 13});
  cols.push({min: 8, max: 8, width: 3}); // H spacer
  cols.push({min: 9, max: 9, width: 18}); // I
  cols.push({min: 10, max: 10, width: 7}); // J

  const sheet = new ExcelSheet("Calendar");
  sheet.setCols(cols);

  // --- Build rows with both legend and calendar cells in the same row when possible ---
  // Row plan:
  // 1: Legend header (I1:J1, merged) and Month header (A1:G1, merged)
  // 2..(1+eventRows): Legend rows (I/J)
  // 2: Day-of-week header (A2:G2)
  // 3+: Calendar grid (A-G only), event rows only for real days

  // Helper: column letter to index (A=1, B=2, ... J=10)
  function colToIndex(col) {
    return col.charCodeAt(0) - 64;
  }

  // ExcelRow with unique columns and sorted output
  class UniqueExcelRow extends ExcelRow {
    constructor(r) {
      super(r);
      this.cellMap = new Map(); // col letter -> cell
    }
    addCell(cell) {
      this.cellMap.set(cell.col, cell); // overwrite if duplicate
    }
    toXML() {
      // Output cells in order A, B, ..., J
      const ordered = Array.from(this.cellMap.values()).sort((a, b) => colToIndex(a.col) - colToIndex(b.col));
      return `<row r='${this.r}'>${ordered.map(c => c.toXML()).join('')}</row>`;
    }
  }

  const rowMap = new Map();
  function getRow(r) {
    if (!rowMap.has(r)) rowMap.set(r, new UniqueExcelRow(r));
    return rowMap.get(r);
  }

  // --- Place legend and calendar at the very top, side by side ---
  // Row 1: Month header (A-G, merged)
  let headerRow = 1;
  let calDaysRow = 2;

  // Month header (A1:G1, merged)
  let monthRow = getRow(headerRow);
  monthRow.addCell(new ExcelCell('A', headerRow, monthNames[month].toUpperCase() + ' ' + year, {style: 4, align: 'center'}));
  sheet.addMerge(`A${headerRow}:G${headerRow}`);
  // Legend header (I1:J1, merged) -- REMOVE legend header in cells, only floating legend now
  // monthRow.addCell(new ExcelCell('I', headerRow, 'Legend', {style: 5}));
  // sheet.addMerge(`I${headerRow}:J${headerRow}`);

  // Legend rows (I/J), rows 2..(1+eventRows) -- REMOVE legend rows in cells, only floating legend now
  // for (let l = 0; l < eventRows; l++) {
  //   let r = legendStartRow + l;
  //   let row = getRow(r);
  //   row.addCell(new ExcelCell('I', r, 'Enter Value Here', {style: 3}));
  //   row.addCell(new ExcelCell('J', r, '', {style: 6 + l}));
  // }

  // Day-of-week header (A2:G2)
  let dowRow = getRow(calDaysRow);
  for (let d = 0; d < 7; d++) {
    dowRow.addCell(new ExcelCell(String.fromCharCode(65 + d), calDaysRow, daysOfWeek[d], {style: 2}));
  }

  // Calendar grid (A-G only), starting at row 3
  let calGridStartRow = 3;
  let currentRow = calGridStartRow;
  let day = 1;
  let firstWeek = true;
  while (day <= daysInMonth) {
    let weekCols = [];
    let weekRow = getRow(currentRow);
    for (let dow = 0; dow < 7; dow++) {
      if (firstWeek && dow < startDay) {
        weekCols.push(null);
        continue;
      }
      if (day > daysInMonth) {
        weekCols.push(null);
        continue;
      }
      weekRow.addCell(new ExcelCell(String.fromCharCode(65 + dow), currentRow, day, {style: 3}));
      weekCols.push(day);
      day++;
    }
    firstWeek = false;
    // Only add event rows for columns with real days (ragged grid, no extra rows)
    for (let er = 0; er < eventRows; er++) {
      currentRow++;
      let eventRow = getRow(currentRow);
      let hasEventCell = false;
      for (let dow = 0; dow < 7; dow++) {
        if (weekCols[dow] !== null && weekCols[dow] !== undefined) {
          eventRow.addCell(new ExcelCell(String.fromCharCode(65 + dow), currentRow, '', {style: 3}));
          hasEventCell = true;
        }
      }
      // Only add the row if it has at least one event cell (no empty event rows)
      if (!hasEventCell) rowMap.delete(currentRow);
    }
    currentRow++;
  }

  // Add all rows to the sheet in order
  const allRows = Array.from(rowMap.keys()).sort((a, b) => a - b);
  for (const r of allRows) {
    // Center align month/year header row
    if (r === headerRow) {
      // Add center alignment to the merged cell
      monthRow.cells.forEach(cell => cell.style = 4); // style 4 is header big+fill+border
    }
    sheet.addRow(rowMap.get(r));
  }
  let xml = sheet.toXML();
  if (includeDrawing) {
    // Insert <drawing r:id="rId1"/> before </worksheet>
    xml = xml.replace(/<\/worksheet>$/, '  <drawing r:id="rId1"/>\n</worksheet>');
  }
  return xml;
}

// Relationship for sheet2.xml to drawing1.xml
function getSheet2RelsXML() {
  return `<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>\n</Relationships>`;
}

// Empty rels for drawing1.xml (shapes only)
function getDrawing1RelsXML() {
  return `<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
}

function getStylesXML(eventRows) {
  const palette = [
    "FFDC143C", "FF228B22", "FF1E90FF", "FFFFA500", "FF800080",
    "FFFFFF00", "FF00CED1", "FF8B4513", "FF4682B4"
  ];
  const fills = [
    `<fill><patternFill patternType="none"/></fill>`,     // 0: none
    `<fill><patternFill patternType="gray125"/></fill>`,  // 1: gray125
    `<fill><patternFill patternType="solid"><fgColor rgb="FFB6D7A8"/><bgColor indexed="64"/></patternFill></fill>`, // 2: header highlight
    `<fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor indexed="64"/></patternFill></fill>` // 3: legend header fill
  ];
  // Add color fills for each legend row
  for (let i = 0; i < eventRows; i++) {
    fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${palette[i]}"/><bgColor indexed="64"/></patternFill></fill>`);
  }
  // Fonts: 0 = normal, 1 = bold, 2 = header big bold, 3 = legend header bold
  const fonts = [
    '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
    '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
    '<font><b/><sz val="16"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
    '<font><b/><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>'
  ];
  // Borders: 0 = none, 1 = thin box
  const borders = [
    '<border/>',
    '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>'
  ];
  // CellXfs: Normal, Bold, Bold+Border, Normal+Border, HeaderBig+Fill+Border, LegendHeader+Fill+Border, then colors+border for legend cells
  const cellXfs = [
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>', // 0: Normal
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>', // 1: Bold
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0"/>', // 2: Bold+border
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>', // 3: Normal+border
    // 4: Header big+fill+border (month) -- add center alignment
    '<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>',
    '<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0"/>'   // 5: Legend header+fill+border
  ];
  for (let i = 0; i < eventRows; i++) {
    cellXfs.push(`<xf numFmtId="0" fontId="0" fillId="${i + 4}" borderId="1" xfId="0"/>`);
  }
  return `<?xml version="1.0" encoding="UTF-8"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n  <fonts count="${fonts.length}">${fonts.join('')}</fonts>\n  <fills count="${fills.length}">${fills.join('')}</fills>\n  <borders count="${borders.length}">${borders.join('')}</borders>\n  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n  <cellXfs count="${cellXfs.length}">${cellXfs.join('')}</cellXfs>\n</styleSheet>`;
}

function getTrackerSheetXML(eventRows) {
  let rows = `<row><c t="inlineStr"><is><t>Legend Value</t></is></c><c t="inlineStr"><is><t>Count</t></is></c></row>`;
  for (let i = 0; i < eventRows; i++) {
    rows += `<row><c t="inlineStr"><is><t>Insert Value Here</t></is></c><c t="inlineStr"><is><t>0</t></is></c></row>`;
  }
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${rows}</sheetData>
</worksheet>`;
}

function getWorkbookXML(includeTracker) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Instructions" sheetId="1" r:id="rId1"/>
    <sheet name="Calendar" sheetId="2" r:id="rId2">
      <tabColor rgb="FF00B050"/>
    </sheet>
    ${includeTracker ? `<sheet name="Tracker" sheetId="3" r:id="rId3"><tabColor rgb="FF7030A0"/></sheet>` : ''}
  </sheets>
</workbook>`;
}

function getWorkbookRelsXML(includeTracker) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  ${includeTracker ? `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>` : ''}
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

// =========================
// App UI & Event Handlers
// =========================
// Handles navigation, form submission, and download logic for the web app.
// --------------------------------------------------

document.addEventListener("DOMContentLoaded", function() {
  // Find all nav items and the main content area
  const navHome = document.getElementById("nav-home");
  const navCalendar = document.getElementById("nav-calendar");
  const navRoundRobin = document.getElementById("nav-roundrobin");
  const mainContent = document.querySelector(".main-content");
  const navItems = document.querySelectorAll(".nav-item");

  // Content templates for each page
  const pages = {
    home: `
      <h2>Welcome to Free Excel Generators!</h2>
      <p>
        This site offers free, easy-to-use tools for creating custom Excel resources like printable calendars and round robin tournament schedules.<br>
        Choose a tool from the sidebar to get started, customize it to your needs, and download your finished Excel file with just a click.
      </p>
    `,
    calendar: `
      <h2>Custom Excel Calendar Builder</h2>
      <form id="calendarForm">
        <label for="year">Year:</label>
        <input type="number" id="year" min="1900" max="2100" value="2024" required>
        <label for="month">Month:</label>
        <select id="month" required>
          <option value="0">January</option>
          <option value="1">February</option>
          <option value="2">March</option>
          <option value="3">April</option>
          <option value="4">May</option>
          <option value="5">June</option>
          <option value="6">July</option>
          <option value="7">August</option>
          <option value="8">September</option>
          <option value="9">October</option>
          <option value="10">November</option>
          <option value="11">December</option>
        </select>
<label for="eventRows">Event Rows per Day:</label>
<select id="eventRows" required>
  <option value="1">1</option>
  <option value="2">2</option>
  <option value="3">3</option>
  <option value="4">4</option>
  <option value="5">5</option>
  <option value="6">6</option>
  <option value="7">7</option>
  <option value="8">8</option>
  <option value="9">9</option>
</select>

<label for="includeTracker" style="margin-left: 12px;">
  <input type="checkbox" id="includeTracker">
  Include Tracker Sheet?
</label>
        <button type="submit">Generate Calendar</button>
        <div id="calendarPreview"></div>
        <button id="downloadTestZipBtn" type="button" style="display:none;">
          <img src="images/Download Icon.svg" alt="Download Icon" style="width:24px;height:24px;vertical-align:middle;">
          Download ZIP
        </button>
      </form>
    `,
    roundrobin: `
      <h2>Round Robin Sports Scheduler</h2>
      <p>Coming soon: Generate balanced sports schedules and export to Excel.</p>
    `
  };

  // Helper function to show a page
  function showPage(page) {
    navItems.forEach(item => item.classList.remove("active"));
    if (page === "home") navHome.classList.add("active");
    if (page === "calendar") navCalendar.classList.add("active");
    if (page === "roundrobin") navRoundRobin.classList.add("active");
    mainContent.innerHTML = pages[page];

    // Attach download button handler ONLY after rendering calendar page
    if (page === "calendar") {
      const downloadBtn = document.getElementById("downloadTestZipBtn");
      if (downloadBtn) {
        downloadBtn.onclick = function() {
          console.log("Download button clicked (calendar page render)");
          // Only allow download if calendar preview exists
          const calendarPreview = document.getElementById("calendarPreview");
          if (!calendarPreview || !calendarPreview.innerHTML.trim()) {
            alert("Please generate a calendar first.");
            return;
          }
          // Get year, month, eventRows, includeTracker from form
          const year = parseInt(document.getElementById("year").value, 10);
          const month = parseInt(document.getElementById("month").value, 10);
          const eventRows = parseInt(document.getElementById("eventRows").value, 10);
          const includeTracker = document.getElementById("includeTracker").checked;

          // Build all Excel XML files
          const files = [
            { name: "[Content_Types].xml", content: getContentTypesXML(includeTracker) },
            { name: "_rels/.rels", content: getRelsXML() },
            { name: "xl/workbook.xml", content: getWorkbookXML(includeTracker) },
            { name: "xl/worksheets/sheet1.xml", content: getSheet1InstructionsXML() },
            // Add drawing reference to calendar sheet XML
            { name: "xl/worksheets/sheet2.xml", content: buildCalendarSheetWithExcelBuilder(year, month, eventRows, true) },
            { name: "xl/styles.xml", content: getStylesXML(eventRows) },
            // Add drawing XML for floating legend
            { name: "xl/drawings/drawing1.xml", content: getLegendDrawingXML(eventRows) },
            // Add rels for sheet2.xml to reference drawing
            { name: "xl/worksheets/_rels/sheet2.xml.rels", content: getSheet2RelsXML() },
            // Add empty rels for drawing1.xml
            { name: "xl/drawings/_rels/drawing1.xml.rels", content: getDrawing1RelsXML() }
          ];
          if (includeTracker) {
            files.push({ name: "xl/worksheets/sheet3.xml", content: getTrackerSheetXML(eventRows) });
          }
          files.push({ name: "xl/_rels/workbook.xml.rels", content: getWorkbookRelsXML(includeTracker) });

          try {
            const zipBytes = createZip(files);
            const blob = new Blob([zipBytes], {type:"application/zip"});
            const a = document.createElement("a");
            a.href = URL.createObjectURL(blob);
            a.download = "calendar.xlsx";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(a.href);
          } catch (e) {
            alert("Failed to generate download: " + e.message);
          }
        };
      }
    }
  }

  // Event listeners for nav
  navHome.addEventListener("click", () => showPage("home"));
  navCalendar.addEventListener("click", () => showPage("calendar"));
  navRoundRobin.addEventListener("click", () => showPage("roundrobin"));

  // Listen for calendar form submission (dynamic content!)
  mainContent.addEventListener("submit", function(event) {
    if (event.target && event.target.id === "calendarForm") {
      event.preventDefault();

      // Get year and month from the form
      const year = parseInt(document.getElementById("year").value, 10);
      const month = parseInt(document.getElementById("month").value, 10);

      // Get Event Rows & Tracker Option
      const eventRows = parseInt(document.getElementById("eventRows").value, 10);
      const includeTracker = document.getElementById("includeTracker").checked;

      // Generate calendar HTML preview (not Excel, just on-page)
      const calendarHTML = generateCalendar(year, month);
      document.getElementById("calendarPreview").innerHTML = calendarHTML;

      // Show the download ZIP button
      const downloadBtn = document.getElementById("downloadTestZipBtn");
      if (downloadBtn) {
        downloadBtn.style.display = "inline-block";
        // Re-attach the download handler after DOM update
        downloadBtn.onclick = function() {
          console.log("Download button clicked (after preview render)");
          // Only allow download if calendar preview exists
          const calendarPreview = document.getElementById("calendarPreview");
          if (!calendarPreview || !calendarPreview.innerHTML.trim()) {
            alert("Please generate a calendar first.");
            return;
          }
          // Get year, month, eventRows, includeTracker from form
          const year = parseInt(document.getElementById("year").value, 10);
          const month = parseInt(document.getElementById("month").value, 10);
          const eventRows = parseInt(document.getElementById("eventRows").value, 10);
          const includeTracker = document.getElementById("includeTracker").checked;

          // Build all Excel XML files
          const files = [
            { name: "[Content_Types].xml", content: getContentTypesXML(includeTracker) },
            { name: "_rels/.rels", content: getRelsXML() },
            { name: "xl/workbook.xml", content: getWorkbookXML(includeTracker) },
            { name: "xl/worksheets/sheet1.xml", content: getSheet1InstructionsXML() },
            // Add drawing reference to calendar sheet XML
            { name: "xl/worksheets/sheet2.xml", content: buildCalendarSheetWithExcelBuilder(year, month, eventRows, true) },
            { name: "xl/styles.xml", content: getStylesXML(eventRows) },
            // Add drawing XML for floating legend
            { name: "xl/drawings/drawing1.xml", content: getLegendDrawingXML(eventRows) },
            // Add rels for sheet2.xml to reference drawing
            { name: "xl/worksheets/_rels/sheet2.xml.rels", content: getSheet2RelsXML() },
            // Add empty rels for drawing1.xml
            { name: "xl/drawings/_rels/drawing1.xml.rels", content: getDrawing1RelsXML() }
          ];
          if (includeTracker) {
            files.push({ name: "xl/worksheets/sheet3.xml", content: getTrackerSheetXML(eventRows) });
          }
          files.push({ name: "xl/_rels/workbook.xml.rels", content: getWorkbookRelsXML(includeTracker) });

          try {
            const zipBytes = createZip(files);
            const blob = new Blob([zipBytes], {type:"application/zip"});
            const a = document.createElement("a");
            a.href = URL.createObjectURL(blob);
            a.download = "calendar.xlsx";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(a.href);
          } catch (e) {
            alert("Failed to generate download: " + e.message);
          }
        };
      }

      // Scroll to preview and download button for better UX
      setTimeout(() => {
        const preview = document.getElementById("calendarPreview");
        if (preview) preview.scrollIntoView({ behavior: "smooth", block: "center" });
      }, 100);
    }
  });

  // Function to create a simple calendar table for the preview on the page
  function generateCalendar(year, month) {
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const startDay = new Date(year, month, 1).getDay(); // 0=Sunday
    const monthNames = [
      "January","February","March","April","May","June",
      "July","August","September","October","November","December"
    ];
    let html = `<h3>${monthNames[month]} ${year}</h3><table border="1" cellpadding="4"><tr>
      <th>Sun</th><th>Mon</th><th>Tue</th><th>Wed</th><th>Thu</th><th>Fri</th><th>Sat</th>
    </tr><tr>`;

    // Fill empty cells until first day
    for (let i = 0; i < startDay; i++) html += "<td></td>";

    // Fill the days of the month
    for (let day = 1; day <= daysInMonth; day++) {
      html += `<td>${day}</td>`;
      if ((startDay + day) % 7 === 0 && day !== daysInMonth) html += "</tr><tr>";
    }

    html += "</tr></table>";
    return html;
  }
});

// =========================
// Floating Legend Drawing XML
// =========================
// Generates DrawingML for the floating legend panel, header, and value pills.
// User-tunable parameters for precise layout:
//   - headerCol1, headerCol2: columns for header anchor (e.g. 9, 12 for K:M)
//   - headerLeftEMU, headerRightEMU: left/right offsets in EMUs for header
//   - pillCol1, pillCol2: columns for value pill anchors (e.g. 9, 12 for K:M)
//   - pillGroupLeftEMU: left offset in EMUs for value pills
//   - pillWidthEMU: right offset in EMUs for value pills
//   - panelCol1, panelCol2: columns for background panel
//   - All other sizing (height, etc.) is still controlled by constants
function getLegendDrawingXML(eventRows, opts = {}) {
  // --- EMU conversion: 1 inch = 914400 EMUs ---
  const emu = v => Math.round(v * 914400);
  // --- Layout constants (inches) ---
  const headerHeightIn = 0.94;
  const pillHeightIn = 1.0;
  const anchorRowStart = 1; // Excel row 2 (0-based)

  // --- User-overridable anchor/offsets for fine-tuning ---
  const panelCol1 = opts.panelCol1 ?? 9, panelCol2 = opts.panelCol2 ?? 13;   // Panel: J to M (wider)
  const headerCol1 = opts.headerCol1 ?? 9; // K
  const headerCol2 = opts.headerCol2 ?? 12; // M
  const headerLeftEMU = opts.headerLeftEMU ?? emu(0.45); // left edge 0.25" into K
  const headerRightEMU = opts.headerRightEMU ?? emu(0.10); // right edge 0.25" into M
  const pillCol1 = opts.pillCol1 ?? 9; // K
  const pillCol2 = opts.pillCol2 ?? 12; // M
  const pillGroupLeftEMU = opts.pillGroupLeftEMU ?? emu(0.45); // left edge 0.25" into K
  const pillWidthEMU = opts.pillWidthEMU ?? emu(0.10);     // right edge 0.25" into M

  // --- Colors ---
  const headerColor = 'B6D7A8';
  const panelColor = 'FFFFFF';
  const panelBorder = '666666';
  const pillColors = [
    'DC143C', '228B22', '1E90FF', 'FFA500', '800080',
    'FFFF00', '00CED1', '8B4513', '4682B4'
  ];

  // --- Calculate row positions for legend shapes ---
  const headerRow = anchorRowStart;
  let rows = [];
  rows.push({ type: 'header', row: headerRow });
  let currentRow = headerRow + 1;
  rows.push({ type: 'blank', row: currentRow });
  for (let i = 0; i < eventRows; i++) {
    currentRow++;
    rows.push({ type: 'pill', row: currentRow, pillIdx: i });
    currentRow++;
    rows.push({ type: 'blank', row: currentRow });
  }
  const panelFromRow = headerRow;
  const panelToRow = currentRow + 1;

  // --- Begin DrawingML XML ---
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`;

  // --- Panel shape (background, covers header+pills area) ---
  xml += `\n  <twoCellAnchor>\n    <from><col>${panelCol1}</col><colOff>0</colOff><row>${panelFromRow}</row><rowOff>0</rowOff></from>\n    <to><col>${panelCol2}</col><colOff>0</colOff><row>${panelToRow}</row><rowOff>0</rowOff></to>\n    <sp>\n      <nvSpPr><cNvPr id="1" name="LegendPanel"/><cNvSpPr/></nvSpPr>\n      <spPr>\n        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>\n        <a:solidFill><a:srgbClr val="${panelColor}"/></a:solidFill>\n        <a:ln w="38100"><a:solidFill><a:srgbClr val="${panelBorder}"/></a:solidFill></a:ln>\n      </spPr>\n    </sp>\n    <clientData/>\n  </twoCellAnchor>`;

  // --- Header shape (centered in its row) ---
  const headerRowOff = Math.round((emu(1) - emu(headerHeightIn)) / 2);
  xml += `\n  <twoCellAnchor>\n    <from><col>${headerCol1}</col><colOff>${headerLeftEMU}</colOff><row>${headerRow}</row><rowOff>${headerRowOff}</rowOff></from>\n    <to><col>${headerCol2}</col><colOff>${headerRightEMU}</colOff><row>${headerRow}</row><rowOff>${headerRowOff + emu(headerHeightIn)}</rowOff></to>\n    <sp>\n      <nvSpPr><cNvPr id="2" name="LegendHeader"/><cNvSpPr/></nvSpPr>\n      <spPr>\n        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>\n        <a:solidFill><a:srgbClr val="${headerColor}"/></a:solidFill>\n        <a:ln w="0"><a:noFill/></a:ln>\n      </spPr>\n      <txBody>\n        <a:bodyPr anchor="ctr"/>\n        <a:lstStyle/>\n        <a:p>\n          <a:pPr algn="ctr"/>\n          <a:r><a:rPr lang="en-US" sz="1600" b="1"/><a:t>Legend</a:t></a:r>\n        </a:p>\n      </txBody>\n    </sp>\n    <clientData/>\n  </twoCellAnchor>`;

  // --- Value pills (all same width, centered as a group) ---
  let pillId = 3;
  for (const r of rows) {
    if (r.type === 'pill') {
      const pillRowOff = Math.round((emu(1) - emu(pillHeightIn)) / 2);
      xml += `\n  <twoCellAnchor>\n    <from><col>${pillCol1}</col><colOff>${pillGroupLeftEMU}</colOff><row>${r.row}</row><rowOff>${pillRowOff}</rowOff></from>\n    <to><col>${pillCol2}</col><colOff>${pillWidthEMU}</colOff><row>${r.row}</row><rowOff>${pillRowOff + emu(pillHeightIn)}</rowOff></to>\n    <sp>\n      <nvSpPr><cNvPr id="${pillId}" name="LegendPill${r.pillIdx + 1}"/><cNvSpPr/></nvSpPr>\n      <spPr>\n        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>\n        <a:solidFill><a:srgbClr val="${pillColors[r.pillIdx]}"/></a:solidFill>\n        <a:ln w="0"><a:noFill/></a:ln>\n      </spPr>\n      <txBody>\n        <a:bodyPr anchor="ctr"/>\n        <a:lstStyle/>\n        <a:p>\n          <a:pPr algn="ctr"/>\n          <a:r><a:rPr lang="en-US" sz="1300"/><a:t>Enter Value Here</a:t></a:r>\n        </a:p>\n      </txBody>\n    </sp>\n    <clientData/>\n  </twoCellAnchor>`;
      pillId++;
    }
  }
  xml += '\n</wsDr>';
  return xml;
}
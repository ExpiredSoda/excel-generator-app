// generators/stylesXml.js
// Generates styles.xml for Excel

export function getStylesXML(eventRows, customColors = null) {
  const defaultPalette = [
    "FFDC143C", "FF228B22", "FF1E90FF", "FFFFA500", "FF800080",
    "FFFFFF00", "FF00CED1", "FF8B4513", "FF4682B4"
  ];
  // Ensure palette is always ARGB (FF + hex, no #)
  const palette = (customColors || defaultPalette).map(c => {
    if (typeof c === 'string') {
      if (c.startsWith('#')) return 'FF' + c.slice(1).toUpperCase();
      if (c.length === 6) return 'FF' + c.toUpperCase();
      if (c.length === 8) return c.toUpperCase();
    }
    return c;
  });
  
  console.log('getStylesXML called with:');
  console.log('- eventRows:', eventRows);
  console.log('- customColors:', customColors);
  console.log('- final palette:', palette);
  
  const fills = [
    `<fill><patternFill patternType="none"/></fill>`,
    `<fill><patternFill patternType="gray125"/></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFB6D7A8"/><bgColor indexed="64"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor indexed="64"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF9C"/><bgColor indexed="64"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>`
  ];
  
  for (let i = 0; i < eventRows; i++) {
    fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${palette[i]}"/><bgColor indexed="64"/></patternFill></fill>`);
    console.log(`Fill ${i + 7}: color ${palette[i]}`);
  }

  const fonts = [
    '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
    '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
    '<font><b/><sz val="16"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
    '<font><b/><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
    '<font><sz val="13"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
    '<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
  ];

  const borders = [
    '<border/>',
    '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>',
    '<border><bottom style="thin"/></border>',
    '<border><left style="thick"/><right style="thick"/><top style="thick"/><bottom style="thick"/></border>',
    '<border><left style="medium"/><right style="medium"/><top style="medium"/><bottom style="medium"/></border>'
  ];

  const cellXfs = [
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>',
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0"/>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>',
    '<xf numFmtId="0" fontId="2" fillId="2" borderId="2" xfId="0"><alignment horizontal="center" vertical="center"/></xf>',
    '<xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0"/>',
    '<xf numFmtId="0" fontId="5" fillId="5" borderId="4" xfId="0"><alignment horizontal="center" vertical="center"/></xf>',
    '<xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="1" fillId="6" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>'
  ];
  
  // Add legend color styles with center alignment - these use the custom colors
  for (let i = 0; i < eventRows; i++) {
    cellXfs.push(`<xf numFmtId="0" fontId="0" fillId="${i + 7}" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>`);
    console.log(`CellXfs style ${10 + i}: fillId=${i + 7} (color: ${palette[i]})`);
  }

  // Add DXF styles for conditional formatting using the exact same palette
  const dxfs = [];
  for (let i = 0; i < eventRows; i++) {
    dxfs.push(`<dxf><fill><patternFill patternType="solid"><bgColor rgb="${palette[i]}"/></patternFill></fill></dxf>`);
    console.log(`DXF ${i}: bgColor=${palette[i]}`);
  }

  return `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="${fonts.length}">${fonts.join('')}</fonts>
  <fills count="${fills.length}">${fills.join('')}</fills>
  <borders count="${borders.length}">${borders.join('')}</borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${cellXfs.length}">${cellXfs.join('')}</cellXfs>
  <dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>
</styleSheet>`;
}

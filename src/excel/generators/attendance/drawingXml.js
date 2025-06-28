// generators/attendance/drawingXml.js
// Generates drawing XML for chart positioning

/**
 * Generate drawing XML for chart positioning
 * @param {object} options - Chart positioning options
 * @param {number} options.fromCol - Starting column (0-based)
 * @param {number} options.fromRow - Starting row (0-based)
 * @param {number} options.toCol - Ending column (0-based)
 * @param {number} options.toRow - Ending row (0-based)
 * @param {string} options.chartName - Chart name
 * @returns {string} Drawing XML
 */
export function buildDrawingXml(options = {}) {
  const {
    fromCol = 8,      // Column I (0-based) - matches working chart
    fromRow = 1,      // Row 2 (0-based) - matches working chart
    toCol = 17,       // Column R (0-based) - matches working chart
    toRow = 17,       // Row 18 (0-based) - matches working chart
    chartName = 'Legend Usage Chart'
  } = options;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>502584</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>53787</xdr:rowOff></xdr:from><xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>62529</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>48453</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="${chartName}"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{1B8180E1-FF28-818A-B476-B77940BBFC94}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
} 
// excel/core/excelSheet.js
// Excel worksheet representation and XML generation

/**
 * Represents a complete Excel worksheet
 */
export class ExcelSheet {
  constructor(name) {
    this.name = name;
    this.rows = [];
    this.merges = [];
    this.cols = [];
    this.conditionalFormatting = [];
  }

  /**
   * Add a row to the sheet
   * @param {ExcelRow} row - Row to add
   */
  addRow(row) {
    this.rows.push(row);
  }

  /**
   * Add a merge range to the sheet
   * @param {string} ref - Cell range reference (e.g., "A1:C1")
   */
  addMerge(ref) {
    this.merges.push(ref);
  }

  /**
   * Set column definitions for the sheet
   * @param {Array} colDefs - Array of column definitions
   */
  setCols(colDefs) {
    this.cols = colDefs;
  }

  /**
   * Add conditional formatting to the sheet
   * @param {ConditionalFormattingRule} cf - Conditional formatting rule
   */
  addConditionalFormatting(cf) {
    this.conditionalFormatting.push(cf);
  }  /**
   * Generate complete XML representation of the worksheet
   */
  toXML() {
    // Sheet format properties for auto-sizing - use customHeight="0" to enable auto-sizing
    let sheetFormatPr = '<sheetFormatPr defaultRowHeight="15" defaultColWidth="8.43" baseColWidth="10" customHeight="0"/>';
    
    let colsXML = '';
    if (this.cols.length > 0) {
      colsXML = `<cols>${this.cols.map(c => {
        let xml = `<col min="${c.min}" max="${c.max}" width="${c.width}"`;
        if (c.bestFit) xml += ' bestFit="1"';
        if (c.customWidth) xml += ' customWidth="1"';
        xml += '/>';
        return xml;
      }).join('')}</cols>`;
    }
    
    let mergesXML = '';
    if (this.merges.length > 0) {
      mergesXML = `<mergeCells count="${this.merges.length}">${this.merges.map(ref => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`;
    }
    
    let conditionalFormattingXML = '';
    if (this.conditionalFormatting.length > 0) {
      // Group conditional formatting rules by range
      const rangeGroups = {};
      this.conditionalFormatting.forEach(cf => {
        if (!rangeGroups[cf.sqref]) rangeGroups[cf.sqref] = [];
        rangeGroups[cf.sqref].push(cf);
      });
      
      conditionalFormattingXML = Object.keys(rangeGroups).map(sqref => 
        `<conditionalFormatting sqref="${sqref}">${rangeGroups[sqref].map(cf => cf.toXML()).join('')}</conditionalFormatting>`
      ).join('');
    }
      // Add xmlns:r for relationships (needed for <drawing r:id=...>)
    return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${sheetFormatPr}
  ${colsXML}
  <sheetData>
    ${this.rows.map(r => r.toXML()).join('\n    ')}
  </sheetData>
  ${mergesXML}
  ${conditionalFormattingXML}
</worksheet>`;
  }
}

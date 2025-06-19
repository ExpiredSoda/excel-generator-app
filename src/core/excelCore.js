// core/excelCore.js
// Core Excel classes and XML escape helper

export function escapeXml(unsafe) {
  if (typeof unsafe !== 'string') {
    unsafe = String(unsafe);
  }
  return unsafe.replace(/[<>&"']/g, function (c) {
    switch (c) {
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '&': return '&amp;';
      case '"': return '&quot;';
      case "'": return '&apos;';
    }
  });
}

export class ExcelCell {
  constructor(col, row, value = '', opts = {}) {
    this.col = col;
    this.row = row;
    this.value = value;
    this.type = opts.type || 'inlineStr';
    this.style = opts.style || 0;
    this.mergeAcross = opts.mergeAcross || 0;
  }
  get ref() {
    return `${this.col}${this.row}`;
  }
  toXML() {
    let attrs = `r="${this.ref}"`;
    if (this.style) attrs += ` s="${this.style}"`;
    let valNode = '';
    if (this.value !== '' && this.value !== null && typeof this.value !== 'undefined') {
      if (this.type === 'n') {
        attrs += ' t="n"';
        valNode = `<v>${this.value}</v>`;
      } else {
        attrs += ` t="${this.type}"`;
        const escapedValue = escapeXml(this.value);
        valNode = `<is><t>${escapedValue}</t></is>`;
      }
    } else {
      if (this.type && this.type !== 'inlineStr') {
        attrs += ` t="${this.type}"`;
      }
    }
    return `<c ${attrs}>${valNode}</c>`;
  }
}

export class ExcelRow {
  constructor(r) {
    this.r = r;
    this.cells = [];
  }
  addCell(cell) {
    this.cells.push(cell);
  }
  toXML() {
    return `<row r="${this.r}">${this.cells.map(c => c.toXML()).join('')}</row>`;
  }
}

export class ExcelSheet {
  constructor(name) {
    this.name = name;
    this.rows = [];
    this.merges = [];
    this.cols = [];
    this.conditionalFormatting = [];
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
  addConditionalFormatting(cf) {
    this.conditionalFormatting.push(cf);
  }  toXML() {
    let colsXML = '';
    if (this.cols.length > 0) {
      colsXML = `<cols>${this.cols.map(c => `<col min="${c.min}" max="${c.max}" width="${c.width}"/>`).join('')}</cols>`;
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
  ${colsXML}
  <sheetData>
    ${this.rows.map(r => r.toXML()).join('\n    ')}
  </sheetData>
  ${mergesXML}
  ${conditionalFormattingXML}
</worksheet>`;
  }
}

export class ConditionalFormattingRule {
  constructor(sqref, formula, fillColor, priority = 1, useExpression = false, dxfId = null) {
    this.sqref = sqref; // Cell range like "A3:G20"
    this.formula = formula; // Excel formula like "$I$2"
    this.fillColor = fillColor; // RGB color like "FFDC143C"
    this.priority = priority; // Priority for rule evaluation
    this.useExpression = useExpression; // Whether to use expression type vs cellIs
    this.dxfId = dxfId; // Reference to DXF style in styles.xml
  }
  toXML() {
    // Reference DXF by ID if provided, otherwise include inline DXF
    let dxfAttr = '';
    let dxfXML = '';
    
    if (this.dxfId !== null) {
      dxfAttr = ` dxfId="${this.dxfId}"`;
    } else {
      dxfXML = `<dxf><fill><patternFill patternType="solid"><fgColor rgb="${this.fillColor}"/></patternFill></fill></dxf>`;
    }
    
    if (this.useExpression) {
      // Expression type for complex formulas
      return `<cfRule type="expression" priority="${this.priority}"${dxfAttr}><formula>${escapeXml(this.formula)}</formula>${dxfXML}</cfRule>`;
    } else {
      // CellIs type for simple comparisons
      return `<cfRule type="cellIs" operator="equal" priority="${this.priority}"${dxfAttr}><formula>${escapeXml(this.formula)}</formula>${dxfXML}</cfRule>`;
    }
  }
}

export class ExcelBuilder {
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
  getSheetXML(idx) {
    return this.sheets[idx].toXML();
  }
}

// All logic for escapeXml, ExcelCell, ExcelRow, ExcelSheet, ConditionalFormattingRule, and ExcelBuilder matches script.js.
// No changes needed.

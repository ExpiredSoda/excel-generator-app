// excel/core/conditionalFormatting.js
// Excel conditional formatting rules

import { escapeXml } from './xmlUtils.js';

/**
 * Represents a conditional formatting rule for Excel cells
 */
export class ConditionalFormattingRule {
  constructor(sqref, formula, fillColor, priority = 1, useExpression = false, dxfId = null) {
    this.sqref = sqref; // Cell range like "A3:G20"
    this.formula = formula; // Excel formula like "$I$2"
    this.fillColor = fillColor; // RGB color like "FFDC143C"
    this.priority = priority; // Priority for rule evaluation
    this.useExpression = useExpression; // Whether to use expression type vs cellIs
    this.dxfId = dxfId; // Reference to DXF style in styles.xml
  }

  /**
   * Generate XML representation of the conditional formatting rule
   */
  toXML() {
    // Reference DXF by ID if provided, otherwise include inline DXF
    let dxfAttr = '';
    let dxfXML = '';
    
    if (this.dxfId !== null) {
      dxfAttr = ` dxfId="${this.dxfId}"`;
    } else {
      dxfXML = `<dxf><fill><patternFill patternType="solid"><bgColor rgb="${this.fillColor}"/></patternFill></fill></dxf>`;
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

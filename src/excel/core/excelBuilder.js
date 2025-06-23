// excel/core/excelBuilder.js
// Excel workbook builder pattern implementation

console.log('✓ ExcelBuilder: Module loaded');

/**
 * Builder pattern for creating Excel workbooks
 */
export class ExcelBuilder {
  constructor() {
    this.sheets = [];
    this.styles = null;
  }

  /**
   * Add a sheet to the workbook
   * @param {ExcelSheet} sheet - Sheet to add
   */
  addSheet(sheet) {
    this.sheets.push(sheet);
  }

  /**
   * Set styles XML for the workbook
   * @param {string} stylesXML - Complete styles XML content
   */
  setStyles(stylesXML) {
    this.styles = stylesXML;
  }

  /**
   * Get XML for a specific sheet by index
   * @param {number} idx - Sheet index
   * @returns {string} Sheet XML
   */
  getSheetXML(idx) {
    return this.sheets[idx].toXML();
  }
}

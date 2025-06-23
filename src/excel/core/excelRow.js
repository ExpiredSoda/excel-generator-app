// excel/core/excelRow.js
// Excel row representation and XML generation

/**
 * Represents a single Excel row containing multiple cells
 */
export class ExcelRow {
  constructor(r) {
    this.r = r;
    this.cells = [];
  }

  /**
   * Add a cell to this row
   * @param {ExcelCell} cell - Cell to add
   */
  addCell(cell) {
    this.cells.push(cell);
  }

  /**
   * Generate XML representation of the row
   */
  toXML() {
    return `<row r="${this.r}">${this.cells.map(c => c.toXML()).join('')}</row>`;
  }
}

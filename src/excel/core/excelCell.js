// excel/core/excelCell.js
// Excel cell representation and XML generation

import { escapeXml } from './xmlUtils.js';

console.log('✓ ExcelCell: Module loaded');

/**
 * Represents a single Excel cell
 */
export class ExcelCell {
  constructor(col, row, value = '', opts = {}) {
    this.col = col;
    this.row = row;
    this.value = value;
    this.type = opts.type || 'inlineStr';
    this.style = opts.style || 0;
    this.mergeAcross = opts.mergeAcross || 0;
  }

  /**
   * Get cell reference (e.g., "A1")
   */
  get ref() {
    return `${this.col}${this.row}`;
  }

  /**
   * Generate XML representation of the cell
   */
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

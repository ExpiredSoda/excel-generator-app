// xmlHelpers.js - XML utility functions

/**
 * Escape XML special characters
 * @param {*} unsafe - Value to escape
 * @returns {string} - Escaped XML string
 */
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
      default: return c;
    }
  });
}

/**
 * Convert color to ARGB format
 * @param {string} color - Color in various formats (#RGB, #RRGGBB, RRGGBB, AARRGGBB)
 * @returns {string} - ARGB color string
 */
export function toARGB(color) {
  if (typeof color !== 'string') return 'FF000000';
  
  // Remove # if present
  color = color.replace('#', '').toUpperCase();
  
  // Handle different formats
  if (color.length === 3) {
    // RGB -> RRGGBB
    color = color[0] + color[0] + color[1] + color[1] + color[2] + color[2];
  }
  
  if (color.length === 6) {
    // RRGGBB -> AARRGGBB (add FF for full opacity)
    color = 'FF' + color;
  }
  
  if (color.length === 8) {
    return color;
  }
  
  // Fallback to black
  return 'FF000000';
}

/**
 * Validate Excel cell reference
 * @param {string} ref - Cell reference like "A1"
 * @returns {boolean} - Whether the reference is valid
 */
export function isValidCellRef(ref) {
  return /^[A-Z]+[1-9]\d*$/.test(ref);
}

/**
 * Validate Excel range reference
 * @param {string} range - Range reference like "A1:B10"
 * @returns {boolean} - Whether the range is valid
 */
export function isValidRange(range) {
  const parts = range.split(':');
  return parts.length === 2 && parts.every(isValidCellRef);
}

/**
 * Convert column number to letter (1 = A, 2 = B, etc.)
 * @param {number} num - Column number (1-based)
 * @returns {string} - Column letter(s)
 */
export function numberToColumn(num) {
  let result = '';
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}

/**
 * Convert column letter to number (A = 1, B = 2, etc.)
 * @param {string} letter - Column letter(s)
 * @returns {number} - Column number (1-based)
 */
export function columnToNumber(letter) {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result;
}
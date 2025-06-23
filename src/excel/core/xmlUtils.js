// excel/core/xmlUtils.js
// XML utility functions for Excel generation

/**
 * Escapes special XML characters in strings
 * @param {any} unsafe - Value to escape
 * @returns {string} XML-safe string
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
    }
  });
}

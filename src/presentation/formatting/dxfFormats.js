// presentation/formatting/dxfFormats.js
// DXF (Differential Formatting) for conditional formatting

console.log('✓ DxfFormats: Module loaded');

/**
 * Generate DXF elements for conditional formatting
 * @param {Array} customColors - Array of color strings for conditional formatting
 * @returns {Array} Array of DXF XML elements
 */
export function createDxfElements(customColors = []) {
  return customColors.map(color => 
    `<dxf><fill><patternFill patternType="solid"><bgColor rgb="${color}"/></patternFill></fill></dxf>`
  );
}

/**
 * Generate complete DXF XML section
 * @param {Array} customColors - Array of color strings
 * @returns {string} Complete DXF XML
 */
export function generateDxfXML(customColors = []) {
  const dxfElements = createDxfElements(customColors);
  
  return dxfElements.length > 0 ? 
    `<dxfs count="${dxfElements.length}">${dxfElements.join('')}</dxfs>` : 
    '<dxfs count="0"/>';
}

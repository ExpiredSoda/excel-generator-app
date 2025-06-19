// utils/validation.js
// Legend input validation using sanitizeLegendInput
import { sanitizeLegendInput } from './sanitize.js';

export function validateLegendInput(input) {
  const sanitized = sanitizeLegendInput(input);
  const suspiciousPatterns = [
    /eval\s*\(/i,
    /function\s*\(/i,
    /alert\s*\(/i,
    /document\./i,
    /window\./i,
    /\$\{.*\}/i,
    /\[\s*\]/i,
    /\.\s*constructor/i
  ];
  for (const pattern of suspiciousPatterns) {
    if (pattern.test(sanitized)) {
      return 'Enter Value Here';
    }
  }
  return sanitized;
}

// Validation utilities for color selection and form inputs

/**
 * Validate color selection to prevent duplicates
 * @param {number} selectedIndex - Index of color picker being changed
 * @param {string} newColor - New color value
 * @returns {boolean} - Whether the color selection is valid
 */
export function validateColorSelection(selectedIndex, newColor) {
  const allColorPickers = document.querySelectorAll('.legend-color-picker');
  
  // Check if this color is already used by another row
  for (let i = 0; i < allColorPickers.length; i++) {
    if (i !== selectedIndex && allColorPickers[i].value === newColor) {
      // Color already in use, show warning and revert
      alert(`This color is already used by another legend item. Please choose a different color.`);
      return false;
    }
  }
  
  return true;
}

/**
 * Convert RGB color to hex format
 * @param {string} rgb - RGB color string
 * @returns {string|null} - Hex color string or null if invalid
 */
export function rgbToHex(rgb) {
  if (rgb.startsWith('#')) return rgb;
  
  const result = rgb.match(/\d+/g);
  if (!result || result.length < 3) return null;
  
  return "#" + result.slice(0, 3).map(x => {
    const hex = parseInt(x).toString(16);
    return hex.length === 1 ? "0" + hex : hex;
  }).join("");
}

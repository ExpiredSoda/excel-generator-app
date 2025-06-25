// generators/attendance/stylesXml.js
// Generates styles.xml for shift tracker using universal style system
import { getUniversalStylesXML } from '../../../presentation/index.js';

// Debug logs removed

export function getShiftTrackerStylesXML(customColors = []) {
  // Debug logs removed
  // Use universal style system, passing customColors for dynamic fills/styles
  return getUniversalStylesXML(customColors);
}
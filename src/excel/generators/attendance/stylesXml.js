// generators/attendance/stylesXml.js
// Generates styles.xml for shift tracker using universal style system
import { getUniversalStylesXML } from '../../../presentation/index.js';

// Debug logs removed

export function getShiftTrackerStylesXML() {
  // Debug logs removed
  // Use universal style system - no custom colors needed for attendance tracker
  return getUniversalStylesXML([]);
}
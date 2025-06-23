// generators/attendance/stylesXml.js
// Generates styles.xml for shift tracker using universal style system
import { getUniversalStylesXML } from '../../../presentation/index.js';

// Debug: Track successful imports and module loading
console.log('✓ AttendanceStyles imports loaded:', {
  getUniversalStylesXML: typeof getUniversalStylesXML
});

export function getShiftTrackerStylesXML() {
  console.log('🎨 AttendanceStyles: Generating shift tracker styles...');
  // Use universal style system - no custom colors needed for attendance tracker
  return getUniversalStylesXML([]);
}
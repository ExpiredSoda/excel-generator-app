// presentation/styles/fills.js
// Fill/background definitions for Excel styling

console.log('✓ Fills: Module loaded');

/**
 * Standard fill patterns for backgrounds
 */
export const BASE_FILLS = [
  '<fill><patternFill patternType="none"/></fill>', // 0 - No fill
  '<fill><patternFill patternType="gray125"/></fill>', // 1 - Gray125
  '<fill><patternFill patternType="solid"><fgColor rgb="FF20B388"/><bgColor indexed="64"/></patternFill></fill>', // 2 - Primary green
  '<fill><patternFill patternType="solid"><fgColor rgb="FFF8F9FA"/><bgColor indexed="64"/></patternFill></fill>', // 3 - Light gray
  '<fill><patternFill patternType="solid"><fgColor rgb="FFF0F8FF"/><bgColor indexed="64"/></patternFill></fill>', // 4 - Light blue
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEAA7"/><bgColor indexed="64"/></patternFill></fill>', // 5 - Light yellow
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>', // 6 - White
  '<fill><patternFill patternType="solid"><fgColor rgb="FFE6E6FA"/><bgColor indexed="64"/></patternFill></fill>', // 7 - Light lavender (calendar)
  '<fill><patternFill patternType="solid"><fgColor rgb="FFECF0F1"/><bgColor indexed="64"/></patternFill></fill>', // 8 - Very light gray (section background)
  '<fill><patternFill patternType="solid"><fgColor rgb="FFE8F5E8"/><bgColor indexed="64"/></patternFill></fill>', // 9 - Very light green (section background)
  '<fill><patternFill patternType="solid"><fgColor rgb="FFF0F8FF"/><bgColor indexed="64"/></patternFill></fill>', // 10 - Very light blue (section background)
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFEF9E7"/><bgColor indexed="64"/></patternFill></fill>', // 11 - Very light yellow (section background)
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFDE8E8"/><bgColor indexed="64"/></patternFill></fill>' // 12 - Light red (warning)
];

/**
 * Fill ID constants for easy reference
 */
export const FILL_IDS = {
  NONE: 0,
  GRAY125: 1,
  PRIMARY_GREEN: 2,
  LIGHT_GRAY: 3,
  LIGHT_BLUE: 4,
  LIGHT_YELLOW: 5,
  WHITE: 6,
  LAVENDER: 7,
  SECTION_GRAY: 8,
  SECTION_GREEN: 9,
  SECTION_BLUE: 10,
  SECTION_YELLOW: 11,
  LIGHT_RED: 12
};

/**
 * Generate custom fill patterns for dynamic colors
 * @param {Array} customColors - Array of color strings (e.g., ['FFDC143C'])
 * @returns {Array} Array of fill XML strings
 */
export function createCustomFills(customColors = []) {
  return customColors.map(color => 
    `<fill><patternFill patternType="solid"><fgColor rgb="${color}"/><bgColor indexed="64"/></patternFill></fill>`
  );
}

/**
 * Get all fills including custom ones
 * @param {Array} customColors - Array of custom color strings
 * @returns {Array} Complete array of fill definitions
 */
export function getAllFills(customColors = []) {
  const customFills = createCustomFills(customColors);
  return [...BASE_FILLS, ...customFills];
}

// presentation/styles/borders.js
// Border definitions for Excel styling

console.log('✓ Borders: Module loaded');

/**
 * Standard border patterns
 */
export const BORDERS = [
  '<border/>', // 0 - No border
  '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>', // 1 - All thin
  '<border><bottom style="medium"/></border>', // 2 - Bottom medium
  '<border><left style="medium"/><right style="medium"/><top style="medium"/><bottom style="medium"/></border>' // 3 - All medium
];

/**
 * Border ID constants for easy reference
 */
export const BORDER_IDS = {
  NONE: 0,
  ALL_THIN: 1,
  BOTTOM_MEDIUM: 2,
  ALL_MEDIUM: 3
};

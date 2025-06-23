// presentation/styles/fonts.js
// Font definitions for Excel styling

console.log('✓ Fonts: Module loaded');

/**
 * Standard font definitions used across all Excel generators
 */
export const FONTS = [
  '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>', // 0 - Default text
  '<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>', // 1 - White header text (bold)
  '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>', // 2 - Regular data text
  '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>', // 3 - Bold text (headers, names)
  '<font><sz val="10"/><color rgb="FF666666"/><name val="Calibri"/><family val="2"/><i/></font>', // 4 - Small italic gray (instructions)
  '<font><b/><sz val="16"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>', // 5 - Large title
  '<font><b/><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>', // 6 - Medium section headers
  '<font><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>', // 7 - Small white text
  '<font><b/><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>', // 8 - Bold 12pt for tracker headers
  '<font><sz val="9"/><color rgb="FF666666"/><name val="Calibri"/><family val="2"/><i/></font>', // 9 - Small italic gray (footer)
  '<font><sz val="11"/><color rgb="FF666666"/><name val="Calibri"/><family val="2"/></font>' // 10 - Regular gray (callout)
];

/**
 * Font ID constants for easy reference
 */
export const FONT_IDS = {
  DEFAULT: 0,
  WHITE_HEADER: 1,
  REGULAR_DATA: 2,
  BOLD_TEXT: 3,
  INSTRUCTIONS: 4,
  LARGE_TITLE: 5,
  SECTION_HEADER: 6,
  SMALL_WHITE: 7,
  TRACKER_HEADER: 8,
  FOOTER_GRAY: 9,
  CALLOUT_GRAY: 10
};

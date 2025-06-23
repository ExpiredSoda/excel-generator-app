// generators/calendar/stylesXml.js
// Generates styles.xml for Excel calendar using the universal style system
import { generateStylesXML } from '../../../presentation/styles/stylesXml.js';

// Debug: Track successful imports
console.log('✓ CalendarStylesXML imports loaded:', {
  generateStylesXML: typeof generateStylesXML
});

export function getCalendarStylesXML(eventRows, customColors = null, finalPalette = null) {
  console.log('🎨 CalendarStylesXML: Generating calendar styles...', {
    eventRows,
    customColorsProvided: !!customColors,
    finalPaletteProvided: !!finalPalette
  });
  
  const defaultPalette = [
    "FFDC143C", "FF228B22", "FF1E90FF", "FFFFA500", "FF800080",
    "FFFFFF00", "FF00CED1", "FF8B4513", "FF4682B4"
  ];
  
  // Use finalPalette if provided, otherwise process customColors or use default
  let palette;
  if (finalPalette) {
    palette = finalPalette;
  } else {
    // Ensure palette is always ARGB (FF + hex, no #)
    palette = (customColors || defaultPalette).map(c => {
      if (typeof c === 'string') {
        if (c.startsWith('#')) return 'FF' + c.slice(1).toUpperCase();
        if (c.length === 6) return 'FF' + c.toUpperCase();
        if (c.length === 8) return c.toUpperCase();
      }
      return c;
    });
  }

  // Only use the colors we need for the calendar (based on eventRows)
  const legendColors = palette.slice(0, eventRows);

  console.log('🎨 CalendarStylesXML: Using universal style system with colors:', {
    legendColors,
    count: legendColors.length
  });

  // Use the universal style system and let it handle all the complexity
  return generateStylesXML(legendColors);
}

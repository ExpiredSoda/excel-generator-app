// Main entry point for the Excel Generator App

import { setupNavigation, showPage } from './web/ui/navigation.js';

document.addEventListener("DOMContentLoaded", function() {
  // Initialize navigation
  setupNavigation();
  // Show home page by default
  showPage('home');

  // Make setupCalendarBuilder available globally for navigation
  window.setupCalendarPage = async function() {
    const { setupCalendarBuilder } = await import('./web/ui/calendarBuilder.js');
    const { buildCalendarSheetWithExcelBuilder } = await import('./excel/generators/calendar/calendarBuilderSheet.js');
    const { getCalendarStylesXML } = await import('./excel/generators/calendar/stylesXml.js');
    const { getContentTypesXML, getRelsXML } = await import('./excel/generators/calendar/contentTypesXml.js');
    const { getSheet1InstructionsXML } = await import('./excel/generators/calendar/instructionsSheet.js');
    const { getWorkbookXML, getWorkbookRelsXML } = await import('./excel/generators/calendar/workbookXml.js');
    const { createZip } = await import('./excel/utils/zipUtils.js');
    setupCalendarBuilder({
      buildCalendarSheetWithExcelBuilder,
      getCalendarStylesXML,
      getContentTypesXML,
      getRelsXML,
      getWorkbookXML,
      getWorkbookRelsXML,
      getSheet1InstructionsXML,
      createZip
    });
  };
});

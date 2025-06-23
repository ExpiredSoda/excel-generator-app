// Main entry point for the Excel Generator App

import { setupNavigation, showPage } from './web/ui/navigation.js';

// Debug: Track successful imports
console.log('✓ Main imports loaded:', {
  setupNavigation: typeof setupNavigation,
  showPage: typeof showPage
});

document.addEventListener("DOMContentLoaded", function() {
  console.log('🚀 Main: DOM loaded, initializing app...');
  
  // Initialize navigation
  setupNavigation();
  
  // Show home page by default
  showPage('home');
  // Make setupCalendarBuilder available globally for navigation
  window.setupCalendarPage = function() {
    // Import calendarBuilder and all the required calendar generation functions
    import('./web/ui/calendarBuilder.js').then(({ setupCalendarBuilder }) => {      import('./excel/generators/calendar/calendarBuilderSheet.js').then(({ buildCalendarSheetWithExcelBuilder }) => {
        import('./excel/generators/calendar/stylesXml.js').then(({ getCalendarStylesXML }) => {
          import('./excel/generators/calendar/contentTypesXml.js').then(({ getContentTypesXML, getRelsXML }) => {
            import('./excel/generators/calendar/instructionsSheet.js').then(({ getSheet1InstructionsXML }) => {
              import('./excel/generators/calendar/workbookXml.js').then(({ getWorkbookXML, getWorkbookRelsXML }) => {
                import('./excel/utils/zipUtils.js').then(({ createZip }) => {
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
                });
              });
            });
          });
        });
      });
    });
  };
  console.log('🎯 Main: App initialization complete - ready for refactoring!');
});

// Main entry point for the Excel Generator App

import { setupNavigation, showPage } from './ui/navigation.js';
import { setupCalendarPageHandlers } from './ui/eventHandlers.js';
import { buildCalendarSheetWithExcelBuilder } from './generators/calendarSheet.js';
import { getStylesXML } from './generators/stylesXml.js';
import { getContentTypesXML, getRelsXML, getSheet1InstructionsXML } from './generators/contentTypesXml.js';
import { getWorkbookXML, getWorkbookRelsXML } from './generators/workbookXml.js';
import { createZip } from './utils/zipWriter.js';

document.addEventListener("DOMContentLoaded", function() {
  console.log('DOM loaded, initializing app...');
  
  // Initialize navigation
  setupNavigation();
  
  // Show home page by default
  showPage('home');
  
  // Setup calendar page functionality globally so navigation can access it
  window.setupCalendarPage = function() {
    console.log('Setting up calendar page handlers...');
    setupCalendarPageHandlers({
      buildCalendarSheetWithExcelBuilder,
      getStylesXML,
      getContentTypesXML,
      getRelsXML,
      getWorkbookXML,
      getWorkbookRelsXML,
      getSheet1InstructionsXML,
      createZip
    });
  };
  
  console.log('App initialization complete');
});

// main.js imports all modular logic and sets up navigation and calendar page handlers on DOMContentLoaded.
// This matches the initialization logic in script.js, but in a modular way.
// No changes needed.

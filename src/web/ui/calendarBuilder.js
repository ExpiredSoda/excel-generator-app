// Calendar Builder Event Handlers
// Handles calendar form interactions, legend management, and Excel generation
import { sanitizeLegendInput } from '../../shared/utils/sanitize.js';
import { validateLegendInput } from '../../shared/utils/validation.js';
import { renderCalendarPreview } from '../utils/previewCalendar.js';

export function setupCalendarBuilder({ buildCalendarSheetWithExcelBuilder, getCalendarStylesXML, getContentTypesXML, getRelsXML, getWorkbookXML, getWorkbookRelsXML, getSheet1InstructionsXML, createZip }) {
  const dependencies = {
    buildCalendarSheetWithExcelBuilder,
    getCalendarStylesXML,
    getContentTypesXML,
    getRelsXML,
    getWorkbookXML,
    getWorkbookRelsXML,
    getSheet1InstructionsXML,
    createZip
  };

  const missingDeps = Object.entries(dependencies).filter(([name, fn]) => typeof fn !== 'function');
  if (missingDeps.length > 0) {
    throw new Error(`Missing required dependencies: ${missingDeps.map(([name]) => name).join(', ')}`);
  }
  const form = document.getElementById('calendarForm');
  const legendFieldsContainer = document.getElementById('legendFieldsContainer');
  const eventRowsSelect = document.getElementById('eventRowsSelect');
  const downloadBtn = document.getElementById('downloadBtn');
  let lastZip = null;

  const domElements = { form, legendFieldsContainer, eventRowsSelect, downloadBtn };
  const missingElements = Object.entries(domElements).filter(([name, el]) => !el);
  if (missingElements.length > 0) {
    // Optionally, handle missing elements gracefully here
  }

  function generateLegendFields(eventRows) {
    const palette = [
      "#DC143C", "#228B22", "#1E90FF", "#FFA500", "#800080",
      "#FFFF00", "#00CED1", "#8B4513", "#4682B4"
    ];
    const defaultNames = [
      "Meeting", "Workout", "Appointment", "Holiday", "Personal",
      "Work", "Travel", "Study", "Event"
    ];
    let html = `<div class="legend-fields"><h4><img src="images/Gear Icon.svg" alt="Settings" style="width:20px;height:20px;vertical-align:middle;margin-right:8px;">Customize Your Legend Values:</h4>`;
    for (let i = 0; i < eventRows; i++) {
      html += `<div class="legend-field-group"><div class="legend-color-picker-container"><input type="color" class="legend-color-picker" data-index="${i}" value="${palette[i % palette.length]}"><div class="legend-color-indicator" style="background-color: ${palette[i % palette.length]};"></div></div><input type="text" class="legend-input" placeholder="${defaultNames[i] || 'Enter Value'}" required></div>`;
    }
    html += `</div>`;
    return html;
  }

  function updateLegendFields() {
    const eventRows = parseInt(eventRowsSelect.value, 10);
    legendFieldsContainer.innerHTML = generateLegendFields(eventRows);
    document.querySelectorAll('.legend-color-picker').forEach(picker => {
      picker.addEventListener('input', handleColorChange);
    });
    document.querySelectorAll('.legend-input').forEach(input => {
      input.addEventListener('input', e => {
        input.value = validateLegendInput(input.value);
      });
    });
  }

  function getLegendColors() {
    const legendColors = [];
    document.querySelectorAll('.legend-color-picker').forEach(picker => {
      legendColors.push(picker.value);
    });
    return legendColors;
  }

  function getLegendValues() {
    const legendValues = [];
    document.querySelectorAll('.legend-input').forEach(input => {
      legendValues.push(sanitizeLegendInput(input.value));
    });
    return legendValues;
  }

  function handleColorChange(event) {
    const picker = event.target;
    const index = parseInt(picker.dataset.index, 10);
    const newColor = picker.value;
    const allColorPickers = document.querySelectorAll('.legend-color-picker');
    const allIndicators = document.querySelectorAll('.legend-color-indicator');
    for (let i = 0; i < allColorPickers.length; i++) {
      if (i !== index && allColorPickers[i].value === newColor) {
        alert('This color is already in use. Please choose a different one.');
        picker.value = '#000000';
        return;
      }
    }
    if (allIndicators[index]) {
      allIndicators[index].style.backgroundColor = newColor;
    }
  }

  eventRowsSelect.addEventListener('change', updateLegendFields);
  updateLegendFields();

  form.addEventListener('submit', async function(e) {
    e.preventDefault();
    const year = parseInt(document.getElementById('yearInput').value, 10);
    const month = parseInt(document.getElementById('monthSelect').value, 10);
    const eventRows = parseInt(eventRowsSelect.value, 10);
    const legendValues = getLegendValues();
    const legendColors = getLegendColors();
    const includeTracker = document.getElementById('includeTracker')?.checked || false;
    // Render HTML preview
    const previewHtml = renderCalendarPreview({ year, month, eventRows, legendValues, legendColors });
    // Move preview just below the form
    let previewDiv = document.getElementById('calendarPreview');
    if (!previewDiv) {
      previewDiv = document.createElement('div');
      previewDiv.id = 'calendarPreview';
      form.parentNode.insertBefore(previewDiv, form.nextSibling);
    }
    previewDiv.innerHTML = previewHtml;
    previewDiv.style.marginTop = '32px';    // Show modern toast notification
    showToast('Calendar generated! Click Download Excel to save your file.');
      // Convert hex colors to ARGB format for Excel
    const legendColorsExcel = legendColors.map(hex => {
      const cleanHex = hex.replace('#', '').toUpperCase();
      return 'FF' + cleanHex; // Add full opacity (FF) prefix
    });

    const calendarSheet = buildCalendarSheetWithExcelBuilder(year, month, eventRows, false, legendValues, legendColorsExcel);
    const stylesXml = getCalendarStylesXML(eventRows, legendColors, legendColorsExcel);
    const contentTypesXml = getContentTypesXML(includeTracker);
    const relsXml = getRelsXML();
    const workbookXml = getWorkbookXML(includeTracker);
    const workbookRelsXml = getWorkbookRelsXML(includeTracker);
    const instructionsXml = getSheet1InstructionsXML();

    const files = [
      { name: 'xl/worksheets/sheet1.xml', content: instructionsXml },
      { name: 'xl/worksheets/sheet2.xml', content: calendarSheet },
      { name: 'xl/styles.xml', content: stylesXml },
      { name: 'xl/workbook.xml', content: workbookXml },
      { name: 'xl/_rels/workbook.xml.rels', content: workbookRelsXml },
      { name: '[Content_Types].xml', content: contentTypesXml },
      { name: '_rels/.rels', content: relsXml }
    ];    if (includeTracker) {
      const { getTrackerSheetXML } = await import('../../excel/generators/calendar/calendarTrackerSheet.js');
      files.push({ name: 'xl/worksheets/sheet3.xml', content: getTrackerSheetXML(eventRows, legendValues) });
    }
    lastZip = createZip(files);
    downloadBtn.style.display = 'inline-flex';
    // Remove old output text
    document.getElementById('output').innerHTML = '';
  });
  downloadBtn.addEventListener('click', function() {
    if (!lastZip) return;
    const blob = new Blob([lastZip], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'SmartCalendar.xlsx';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);
  });

  // Modern toast notification
  function showToast(message) {
    let toast = document.getElementById('calendarToast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'calendarToast';
      document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.className = 'calendar-toast show';
    setTimeout(() => {
      toast.className = 'calendar-toast';
    }, 3000);
  }
}
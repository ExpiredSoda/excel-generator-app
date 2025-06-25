// Calendar Builder Event Handlers
// Handles calendar form interactions, legend management, and Excel generation
import { sanitizeLegendInput } from '../../shared/utils/sanitize.js';
import { validateLegendInput } from '../../shared/utils/validation.js';
import { renderCalendarPreview } from '../utils/previewCalendar.js';

// Toast system for calendar builder
let isShowingToast = false;
let toastQueue = [];

export function setupCalendarBuilder() {
  const mainContent = document.querySelector('.main-content');
  if (!mainContent) return;

  // Wait for user to click Start Building
  const startBtn = document.getElementById('startCalendarBuilderBtn');
  
  if (startBtn) {
    startBtn.addEventListener('click', async () => {
      const template = document.getElementById('calendarPageTemplate');
      if (template) {
        mainContent.innerHTML = template.innerHTML;
        // Dynamically import dependencies as in previous logic
        const { buildCalendarSheetWithExcelBuilder } = await import('../../excel/generators/calendar/calendarBuilderSheet.js');
        const { getCalendarStylesXML } = await import('../../excel/generators/calendar/stylesXml.js');
        const { getContentTypesXML, getRelsXML } = await import('../../excel/generators/calendar/contentTypesXml.js');
        const { getSheet1InstructionsXML } = await import('../../excel/generators/calendar/instructionsSheet.js');
        const { getWorkbookXML, getWorkbookRelsXML } = await import('../../excel/generators/calendar/workbookXml.js');
        const { createZip } = await import('../../excel/utils/zipUtils.js');
        // Now run the original setup logic, but only after the form is injected
        setupCalendarForm({
          buildCalendarSheetWithExcelBuilder,
          getCalendarStylesXML,
          getContentTypesXML,
          getRelsXML,
          getWorkbookXML,
          getWorkbookRelsXML,
          getSheet1InstructionsXML,
          createZip
        });
      }
    });
  } else {
    // Try again after a short delay
    setTimeout(() => {
      const delayedBtn = document.getElementById('startCalendarBuilderBtn');
      if (delayedBtn) {
        delayedBtn.addEventListener('click', async () => {
          const template = document.getElementById('calendarPageTemplate');
          if (template) {
            mainContent.innerHTML = template.innerHTML;
            const { buildCalendarSheetWithExcelBuilder } = await import('../../excel/generators/calendar/calendarBuilderSheet.js');
            const { getCalendarStylesXML } = await import('../../excel/generators/calendar/stylesXml.js');
            const { getContentTypesXML, getRelsXML } = await import('../../excel/generators/calendar/contentTypesXml.js');
            const { getSheet1InstructionsXML } = await import('../../excel/generators/calendar/instructionsSheet.js');
            const { getWorkbookXML, getWorkbookRelsXML } = await import('../../excel/generators/calendar/workbookXml.js');
            const { createZip } = await import('../../excel/utils/zipUtils.js');
            setupCalendarForm({
              buildCalendarSheetWithExcelBuilder,
              getCalendarStylesXML,
              getContentTypesXML,
              getRelsXML,
              getWorkbookXML,
              getWorkbookRelsXML,
              getSheet1InstructionsXML,
              createZip
            });
          }
        });
      }
    }, 100);
  }
}

function setupCalendarForm({ buildCalendarSheetWithExcelBuilder, getCalendarStylesXML, getContentTypesXML, getRelsXML, getWorkbookXML, getWorkbookRelsXML, getSheet1InstructionsXML, createZip }) {
  const mainContent = document.querySelector('.main-content');
  if (!mainContent) return;

  // Do NOT inject any intro or explanation HTML here. Only set up the form logic.

  const form = document.getElementById('calendarForm');
  if (!form) return;
  
  const legendFieldsContainer = document.getElementById('legendFieldsContainer');
  const eventRowsSelect = document.getElementById('eventRowsSelect');
  const downloadBtn = document.getElementById('downloadBtn');
  let lastZip = null;

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
      html += `<div class="legend-field-group">
        <span class="legend-color-dot color-dot-picker" style="background:${palette[i % palette.length]}; position: relative; cursor: pointer;">
          <input type="color" class="legend-color-picker" data-index="${i}" value="${palette[i % palette.length]}" style="opacity:0;position:absolute;left:0;top:0;width:100%;height:100%;cursor:pointer;">
        </span>
        <input type="text" class="legend-input" placeholder="${defaultNames[i] || 'Enter Value'}" required>
      </div>`;
    }
    html += `</div>`;
    return html;
  }

  function updateLegendFields() {
    const eventRows = parseInt(eventRowsSelect.value, 10);
    legendFieldsContainer.innerHTML = generateLegendFields(eventRows);
    
    document.querySelectorAll('.legend-color-dot.color-dot-picker').forEach((dot, i) => {
      const picker = dot.querySelector('.legend-color-picker');
      dot.addEventListener('click', () => picker.click());
      picker.addEventListener('input', () => {
        dot.style.background = picker.value;
      });
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

  eventRowsSelect?.addEventListener('change', updateLegendFields);
  if (eventRowsSelect) updateLegendFields();

  form?.addEventListener('submit', async function(e) {
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
    previewDiv.style.marginTop = '32px';

    // Show modern toast notification using the new system
    showToast('success', 'Calendar generated! Click Download Excel to save your file.');

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
    ];

    if (includeTracker) {
      const { getTrackerSheetXML } = await import('../../excel/generators/calendar/calendarTrackerSheet.js');
      files.push({ name: 'xl/worksheets/sheet3.xml', content: getTrackerSheetXML(eventRows, legendValues) });
    }
    lastZip = createZip(files);
    downloadBtn.style.display = 'inline-flex';
    // Remove old output text
    document.getElementById('output').innerHTML = '';
  });

  downloadBtn?.addEventListener('click', function() {
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

  // Modern toast notification system (matches attendance tracker)
  function showToast(type, message) {
    // Add to queue if a toast is currently showing
    if (isShowingToast) {
      toastQueue.push({ type, message });
      return;
    }

    isShowingToast = true;
    const toast = document.createElement('div');
    toast.className = `modern-toast modern-toast-${type}`;
    toast.innerHTML = `
      <div class="toast-content">
        <span class="toast-icon">${getToastIcon(type)}</span>
        <span class="toast-message">${message}</span>
        <button class="toast-close" onclick="this.parentElement.parentElement.remove()">×</button>
      </div>
    `;
    document.body.appendChild(toast);
    setTimeout(() => toast.classList.add('show'), 10);
    
    const hideToast = () => {
      if (toast.parentElement) {
        toast.classList.remove('show');
        setTimeout(() => {
          toast.remove();
          isShowingToast = false;
          // Show next toast in queue if any
          if (toastQueue.length > 0) {
            const next = toastQueue.shift();
            setTimeout(() => showToast(next.type, next.message), 100);
          }
        }, 300);
      }
    };

    // Auto-hide after 5 seconds
    setTimeout(hideToast, 5000);
    
    // Allow manual close
    const closeBtn = toast.querySelector('.toast-close');
    closeBtn.onclick = hideToast;
  }

  function getToastIcon(type) {
    const icons = {
      'success': '✅',
      'error': '❌',
      'warning': '⚠️',
      'info': 'ℹ️'
    };
    return icons[type] || 'ℹ️';
  }
}
// previewCalendar.js
// Generates an HTML preview of the calendar for the UI

export function renderCalendarPreview({ year, month, eventRows, legendValues, legendColors }) {
  const monthNames = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  const daysOfWeek = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const startDay = new Date(year, month, 1).getDay();
  const actualLegendValues = legendValues || [
    "Meeting", "Workout", "Appointment", "Holiday", "Personal",
    "Work", "Travel", "Study", "Event"
  ].slice(0, eventRows);
  const palette = legendColors || [
    "#DC143C", "#228B22", "#1E90FF", "#FFA500", "#800080",
    "#FFFF00", "#00CED1", "#8B4513", "#4682B4"
  ];

  let html = `<div class="calendar-preview"><h3 class="calendar-preview-title">${monthNames[month]} ${year}</h3><table class="calendar-table"><thead><tr>`;
  for (let d = 0; d < 7; d++) {
    html += `<th>${daysOfWeek[d]}</th>`;
  }
  html += `</tr></thead><tbody>`;
  let day = 1;
  let started = false;
  for (let week = 0; week < 6 && day <= daysInMonth; week++) {
    html += `<tr>`;
    for (let dow = 0; dow < 7; dow++) {
      if (!started && dow === startDay) started = true;
      if (started && day <= daysInMonth) {
        html += `<td><div class="calendar-date">${day}</div>`;
        for (let er = 0; er < eventRows; er++) {
          html += `<div class="calendar-event-row"></div>`;
        }
        html += `</td>`;
        day++;
      } else {
        html += `<td></td>`;
      }
    }
    html += `</tr>`;
  }
  html += `</tbody></table>`;
  // Legend bar with round color dots
  html += `<div class="calendar-legend" style="margin-top:18px;display:flex;align-items:center;flex-wrap:wrap;gap:12px;"><strong style="margin-right:10px;">Legend:</strong>`;
  for (let i = 0; i < actualLegendValues.length; i++) {
    // Abbreviate if needed
    let shortLabel = actualLegendValues[i];
    if (shortLabel.length > 16) shortLabel = shortLabel.slice(0, 13) + '...';
    html += `<span style="display:inline-flex;align-items:center;gap:8px;margin-right:8px;">
      <span style="display:inline-block;width:18px;height:18px;border-radius:50%;background:${palette[i]};vertical-align:middle;margin-right:4px;"></span>
      <span style="font-weight:600;font-size:1rem;">${shortLabel}</span>
    </span>`;
  }
  html += `</div></div>`;
  return html;
}

// New function for date selection calendar
export function generateCalendarPreview(year, month, selectedDates = []) {
  const monthNames = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  const daysOfWeek = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const startDay = new Date(year, month, 1).getDay();

  let html = `<div class="calendar-preview">
    <h3 class="calendar-preview-title">${monthNames[month]} ${year}</h3>
    <table class="calendar-table">
      <thead>
        <tr>`;
  
  for (let d = 0; d < 7; d++) {
    html += `<th>${daysOfWeek[d]}</th>`;
  }
  
  html += `</tr></thead><tbody>`;
  
  let day = 1;
  let started = false;
  
  for (let week = 0; week < 6 && day <= daysInMonth; week++) {
    html += `<tr>`;
    for (let dow = 0; dow < 7; dow++) {
      if (!started && dow === startDay) started = true;
      
      if (started && day <= daysInMonth) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const isSelected = selectedDates.includes(dateStr);
        const isToday = dateStr === new Date().toISOString().split('T')[0];
        
        let cellClass = 'calendar-date';
        let cellStyle = '';
        
        if (isSelected) {
          cellClass += ' selected';
          cellStyle = 'background-color: #20b388; color: white; border-radius: 4px; padding: 4px; cursor: pointer;';
        } else if (isToday) {
          cellStyle = 'background-color: #f0f9f6; border: 2px solid #20b388; border-radius: 4px; padding: 4px; cursor: pointer;';
        } else {
          cellStyle = 'cursor: pointer; padding: 4px;';
        }
        
        html += `<td>
          <div class="${cellClass}" data-date="${dateStr}" style="${cellStyle}">${day}</div>
        </td>`;
        day++;
      } else {
        html += `<td></td>`;
      }
    }
    html += `</tr>`;
  }
  
  html += `</tbody></table>`;
  
  // Add instructions for custom selection
  if (selectedDates.length > 0) {
    html += `<div style="text-align: center; margin-top: 16px; font-size: 14px; color: #666;">
      <strong>${selectedDates.length}</strong> date${selectedDates.length !== 1 ? 's' : ''} selected
    </div>`;
  }
  
  html += `</div>`;
  
  return html;
}

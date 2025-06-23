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

  let html = `<section class="calendar-preview-section"><div class="calendar-preview"><h3 class="calendar-preview-title">${monthNames[month]} ${year}</h3><table class="calendar-table"><thead><tr>`;
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
  // Legend
  html += `<div class="calendar-legend"><strong>Legend:</strong>`;
  for (let i = 0; i < actualLegendValues.length; i++) {
    html += `<span class="calendar-legend-item" style="background:${palette[i]};">${actualLegendValues[i]}</span> `;
  }
  html += `</div></div></section>`;
  return html;
}

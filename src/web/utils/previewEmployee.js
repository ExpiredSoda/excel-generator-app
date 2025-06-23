// utils/previewEmployee.js
// Generates HTML preview of employee cards for the attendance tracker UI

export function renderEmployeePreview(employees) {
  if (!employees || employees.length === 0) {
    return `
      <div class="preview-section">
        <h4>Employee Preview</h4>
        <div class="empty-state">
          <p>No employees added yet. Add employees to see the preview.</p>
        </div>
      </div>
    `;
  }

  let html = `
    <div class="preview-section">
      <h4>Employee Preview</h4>
      <div class="employee-cards">
  `;

  employees.forEach((employee, index) => {
    const shiftHours = calculateShiftHours(employee.shifts);
    
    html += `
      <div class="employee-card" style="border-left: 4px solid ${employee.color}">
        <div class="employee-header">
          <h5 class="employee-name">${escapeHtml(employee.name)}</h5>
          <span class="employee-id">#${escapeHtml(employee.id || 'N/A')}</span>
        </div>
        <div class="employee-details">
          <div class="detail-row">
            <span class="detail-label">Title:</span>
            <span class="detail-value">${escapeHtml(employee.title)}</span>
          </div>
          <div class="detail-row">
            <span class="detail-label">Shift:</span>
            <span class="detail-value">${formatTimeForDisplay(employee.shifts.start)} - ${formatTimeForDisplay(employee.shifts.end)}</span>
          </div>
          <div class="detail-row">
            <span class="detail-label">Daily Hours:</span>
            <span class="detail-value">${shiftHours} hours</span>
          </div>
          ${employee.email ? `
          <div class="detail-row">
            <span class="detail-label">Email:</span>
            <span class="detail-value">${escapeHtml(employee.email)}</span>
          </div>
          ` : ''}
          ${employee.phone ? `
          <div class="detail-row">
            <span class="detail-label">Phone:</span>
            <span class="detail-value">${escapeHtml(employee.phone)}</span>
          </div>
          ` : ''}
        </div>
        <div class="employee-breaks">
          <h6>Break Schedule:</h6>
          <div class="breaks-grid">
            ${employee.shifts.firstBreak ? `<span class="break-time">1st: ${formatTimeForDisplay(employee.shifts.firstBreak)}</span>` : ''}
            ${employee.shifts.lunch ? `<span class="break-time">Lunch: ${formatTimeForDisplay(employee.shifts.lunch)}</span>` : ''}
            ${employee.shifts.secondBreak ? `<span class="break-time">2nd: ${formatTimeForDisplay(employee.shifts.secondBreak)}</span>` : ''}
          </div>
        </div>
        <div class="employee-actions">
          <button class="btn edit-btn btn-sm" onclick="attendanceTracker.editEmployee(${index})">
            <i class="fas fa-edit"></i> Edit
          </button>
          <button class="btn delete-btn btn-sm" onclick="attendanceTracker.removeEmployee(${index})">
            <i class="fas fa-trash"></i> Remove
          </button>
        </div>
      </div>
    `;
  });

  html += `
      </div>
    </div>
  `;

  return html;
}

export function renderEmployeeStats(employees) {
  if (!employees || employees.length === 0) {
    return `
      <div class="stats-section">
        <h4>Team Statistics</h4>
        <div class="empty-state">
          <p>Add employees to see team statistics.</p>
        </div>
      </div>
    `;
  }

  const totalEmployees = employees.length;
  const avgShiftHours = employees.reduce((sum, emp) => sum + parseFloat(calculateShiftHours(emp.shifts)), 0) / totalEmployees;
  const uniqueTitles = [...new Set(employees.map(emp => emp.title))].length;
  const totalWeeklyHours = employees.reduce((sum, emp) => sum + (parseFloat(calculateShiftHours(emp.shifts)) * 5), 0); // Assuming 5-day work week

  return `
    <div class="stats-section">
      <h4>Team Statistics</h4>
      <div class="stats-grid">
        <div class="stat-card">
          <div class="stat-number">${totalEmployees}</div>
          <div class="stat-label">Total Employees</div>
        </div>
        <div class="stat-card">
          <div class="stat-number">${avgShiftHours.toFixed(1)}</div>
          <div class="stat-label">Avg Daily Hours</div>
        </div>
        <div class="stat-card">
          <div class="stat-number">${uniqueTitles}</div>
          <div class="stat-label">Job Titles</div>
        </div>
        <div class="stat-card">
          <div class="stat-number">${totalWeeklyHours.toFixed(0)}</div>
          <div class="stat-label">Weekly Hours</div>
        </div>
      </div>
    </div>
  `;
}

// Helper functions
function escapeHtml(text) {
  if (typeof text !== 'string') return '';
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function formatTimeForDisplay(time) {
  if (!time) return '';
  
  const [hours, minutes] = time.split(':');
  const hour = parseInt(hours);
  const ampm = hour >= 12 ? 'PM' : 'AM';
  const displayHour = hour % 12 || 12;
  
  return `${displayHour}:${minutes} ${ampm}`;
}

function calculateShiftHours(shifts) {
  if (!shifts.start || !shifts.end) return '0';
  
  const startMinutes = timeToMinutes(shifts.start);
  const endMinutes = timeToMinutes(shifts.end);
  
  let totalMinutes = endMinutes - startMinutes;
  if (totalMinutes < 0) totalMinutes += 24 * 60; // Handle overnight shifts
  
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  
  return minutes > 0 ? `${hours}.${Math.round(minutes/60*10)}` : `${hours}`;
}

function timeToMinutes(timeString) {
  const [hours, minutes] = timeString.split(':').map(Number);
  return hours * 60 + minutes;
}
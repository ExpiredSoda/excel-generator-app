// Employee Shift Tracker Module
// Handles employee management, validation, and shift tracking functionality

import { sanitizeEmployeeName, sanitizeEmployeeText, sanitizeEmail, sanitizePhoneNumber } from '../../shared/utils/sanitize.js';
import { validateEmployeeName, validateEmployeeEmail, validateEmployeeTitle, validateEmployeePhone, validateShiftTime, validateBreakTime } from '../../shared/utils/validation.js';
import { renderEmployeePreview, renderEmployeeStats } from '../utils/previewEmployee.js';

export class AttendanceTracker {
  constructor() {
    this.employees = [];
    this.maxEmployees = 60;
    this.editingIndex = -1;
    this.hasUnsavedChanges = false;
    this.loadEmployees();
    this.setupEventListeners();
    this.setupNavigationWarning();
  }

  // --- Helper Methods ---
  updateButtonText(selector, text) {
    const btn = document.querySelector(selector);
    if (btn) btn.textContent = text;
  }

  resetPresetButtons() {
    document.querySelectorAll('.preset-btn').forEach(btn => btn.classList.remove('active'));
    const customBtn = document.querySelector('.preset-btn.custom');
    if (customBtn) customBtn.classList.add('active');
  }

  updateColorIndicator(color) {
    const colorIndicator = document.querySelector('.color-indicator');
    if (colorIndicator) colorIndicator.style.backgroundColor = color;
  }

  // --- Main Methods ---
  sanitizeEmployeeData(data) {
    const nameValidation = validateEmployeeName(data.name);
    const titleValidation = validateEmployeeTitle(data.title);
    const emailValidation = validateEmployeeEmail(data.email);
    const phoneValidation = validateEmployeePhone(data.phone);
    return {
      name: nameValidation.sanitized || data.name,
      id: sanitizeEmployeeText(data.id, 20),
      email: emailValidation.sanitized || '',
      title: titleValidation.sanitized || data.title,
      phone: phoneValidation.sanitized || '',
      color: data.color,
      shifts: { ...data.shifts },
      dateAdded: new Date().toISOString()
    };
  }

  setupEventListeners() {
    const employeeForm = document.getElementById('employeeForm');
    if (employeeForm) {
      employeeForm.addEventListener('submit', (e) => this.handleAddEmployee(e));
    }
    const clearFormBtn = document.getElementById('clearFormBtn');
    if (clearFormBtn) {
      clearFormBtn.addEventListener('click', () => this.clearForm());
    }
    this.setupShiftPresets();
    const csvUploadArea = document.getElementById('csvUploadArea');
    const csvFileInput = document.getElementById('csvFileInput');
    if (csvUploadArea && csvFileInput) {
      csvUploadArea.addEventListener('click', () => csvFileInput.click());
      csvUploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
      csvUploadArea.addEventListener('drop', (e) => this.handleFileDrop(e));
      csvFileInput.addEventListener('change', (e) => this.handleFileSelect(e));
    }
    const exportListBtn = document.getElementById('exportListBtn');
    const clearAllBtn = document.getElementById('clearAllBtn');
    if (exportListBtn) {
      exportListBtn.classList.add('download-btn');
      exportListBtn.addEventListener('click', () => this.exportEmployeeList());
    }
    if (clearAllBtn) {
      clearAllBtn.addEventListener('click', () => this.clearAllEmployees());
    }
    const generateTrackerBtn = document.getElementById('generateTrackerBtn');
    if (generateTrackerBtn) {
      generateTrackerBtn.addEventListener('click', () => this.generateShiftTracker());
    }
    this.setupRealTimeValidation();
  }

  setupShiftPresets() {
    const presetButtons = document.querySelectorAll('.preset-btn');
    presetButtons.forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        const preset = btn.dataset.preset;
        this.applyShiftPreset(preset);
        presetButtons.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
      });
    });
  }

  applyShiftPreset(preset) {
    const presets = {
      'first-shift': { start: '08:00', firstBreak: '10:00', lunch: '12:00', secondBreak: '14:30', end: '16:30' },
      'second-shift': { start: '09:00', firstBreak: '11:00', lunch: '13:00', secondBreak: '15:30', end: '17:30' },
      'third-shift': { start: '11:30', firstBreak: '13:30', lunch: '15:30', secondBreak: '17:30', end: '20:00' }
    };
    if (presets[preset]) {
      const times = presets[preset];
      document.getElementById('shiftStart').value = times.start;
      document.getElementById('firstBreak').value = times.firstBreak;
      document.getElementById('lunchBreak').value = times.lunch;
      document.getElementById('secondBreak').value = times.secondBreak;
      document.getElementById('shiftEnd').value = times.end;
      this.showToast('info', `Applied ${preset.replace('-', ' ')} template`);
    }
  }

  setupNavigationWarning() {
    window.addEventListener('beforeunload', (e) => {
      if (this.employees.length > 0) {
        const message = 'You have employee data that will be lost. Are you sure you want to leave?';
        e.preventDefault();
        e.returnValue = message;
        return message;
      }
    });
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
      if (item.id !== 'nav-attendance') {
        item.addEventListener('click', (e) => {
          if (this.employees.length > 0) {
            e.preventDefault();
            this.showConfirmToast(
              'You have employee data that will be lost. Continue?',
              () => {
                this.employees = [];
                this.clearSavedData();
                const originalEvent = new MouseEvent('click', { view: window, bubbles: true, cancelable: true });
                item.removeEventListener('click', arguments.callee);
                item.dispatchEvent(originalEvent);
              }
            );
          }
        });
      }
    });
    if (this.employees.length === 0) {
      setTimeout(() => {
        this.showToast('info', 'Note: Employee data is stored for this browser session only and will be cleared when you close this tab.');
      }, 1000);
    }
  }

  setupRealTimeValidation() {
    const fields = [
      'employeeName', 'employeeEmail', 'employeeTitle', 
      'shiftStart', 'shiftEnd', 'firstBreak', 'lunchBreak', 'secondBreak'
    ];
    fields.forEach(fieldId => {
      const field = document.getElementById(fieldId);
      if (field) {
        field.addEventListener('blur', () => this.validateField(fieldId));
        field.addEventListener('input', () => this.clearFieldError(fieldId));
      }
    });
    this.setupColorPicker();
  }

  setupColorPicker() {
    const colorPicker = document.getElementById('employeeColor');
    const colorIndicator = document.querySelector('.color-indicator');
    if (colorPicker && colorIndicator) {
      colorPicker.addEventListener('input', (e) => { colorIndicator.style.backgroundColor = e.target.value; });
      colorPicker.addEventListener('change', (e) => { colorIndicator.style.backgroundColor = e.target.value; });
      colorIndicator.addEventListener('click', () => { colorPicker.click(); });
      colorIndicator.style.backgroundColor = colorPicker.value;
    }
  }

  async handleAddEmployee(event) {
    event.preventDefault();
    if (this.employees.length >= this.maxEmployees) {
      this.showToast('error', `Maximum of ${this.maxEmployees} employees allowed.`);
      return;
    }
    const formData = this.getFormData();
    if (!this.validateAllFields(formData)) {
      this.showToast('error', 'Please fix the validation errors before adding the employee.');
      return;
    }
    if (this.editingIndex === -1 && this.isDuplicateEmployee(formData)) {
      this.showToast('error', 'An employee with this name or ID already exists.');
      return;
    }
    const sanitizedEmployee = this.sanitizeEmployeeData(formData);
    if (this.editingIndex === -1) {
      this.employees.push(sanitizedEmployee);
      this.showToast('success', `Employee ${sanitizedEmployee.name} added successfully!`);
    } else {
      this.employees[this.editingIndex] = sanitizedEmployee;
      this.showToast('success', `Employee ${sanitizedEmployee.name} updated successfully!`);
      this.editingIndex = -1;
      this.updateButtonText('.add-employee-btn', 'Add Employee');
    }
    this.saveEmployees();
    this.renderEmployeeList();
    this.clearForm();
    this.updateGenerateButton();
  }

  getFormData() {
    return {
      name: document.getElementById('employeeName')?.value || '',
      id: document.getElementById('employeeId')?.value || '',
      email: document.getElementById('employeeEmail')?.value || '',
      title: document.getElementById('employeeTitle')?.value || '',
      phone: document.getElementById('employeePhone')?.value || '',
      color: document.getElementById('employeeColor')?.value || '#4CAF50',
      shifts: {
        start: document.getElementById('shiftStart')?.value || '',
        firstBreak: document.getElementById('firstBreak')?.value || '',
        lunch: document.getElementById('lunchBreak')?.value || '',
        secondBreak: document.getElementById('secondBreak')?.value || '',
        end: document.getElementById('shiftEnd')?.value || ''
      }
    };
  }

  validateAllFields(data) {
    let isValid = true;
    const nameValidation = validateEmployeeName(data.name);
    if (!nameValidation.isValid) { this.showFieldError('employeeName', nameValidation.error); isValid = false; }
    const titleValidation = validateEmployeeTitle(data.title);
    if (!titleValidation.isValid) { this.showFieldError('employeeTitle', titleValidation.error); isValid = false; }
    const emailValidation = validateEmployeeEmail(data.email);
    if (!emailValidation.isValid) { this.showFieldError('employeeEmail', emailValidation.error); isValid = false; }
    const phoneValidation = validateEmployeePhone(data.phone);
    if (!phoneValidation.isValid) { this.showFieldError('employeePhone', phoneValidation.error); isValid = false; }
    const shiftValidation = validateShiftTime(data.shifts.start, data.shifts.end);
    if (!shiftValidation.isValid) { this.showFieldError('shiftEnd', shiftValidation.error); isValid = false; }
    if (data.shifts.firstBreak) {
      const firstBreakValidation = validateBreakTime(data.shifts.firstBreak, data.shifts.start, data.shifts.end, 'First break');
      if (!firstBreakValidation.isValid) { this.showFieldError('firstBreak', firstBreakValidation.error); isValid = false; }
    }
    if (data.shifts.lunch) {
      const lunchValidation = validateBreakTime(data.shifts.lunch, data.shifts.start, data.shifts.end, 'Lunch break');
      if (!lunchValidation.isValid) { this.showFieldError('lunchBreak', lunchValidation.error); isValid = false; }
    }
    if (data.shifts.secondBreak) {
      const secondBreakValidation = validateBreakTime(data.shifts.secondBreak, data.shifts.start, data.shifts.end, 'Second break');
      if (!secondBreakValidation.isValid) { this.showFieldError('secondBreak', secondBreakValidation.error); isValid = false; }
    }
    return isValid;
  }

  validateShiftTimes(shifts) {
    const startTime = this.timeToMinutes(shifts.start);
    const endTime = this.timeToMinutes(shifts.end);
    if (endTime <= startTime) { this.showFieldError('shiftEnd', 'End time must be after start time'); return false; }
    const breaks = [
      { time: shifts.firstBreak, field: 'firstBreak', name: '1st break' },
      { time: shifts.lunch, field: 'lunchBreak', name: 'lunch break' },
      { time: shifts.secondBreak, field: 'secondBreak', name: '2nd break' }
    ];
    for (const breakTime of breaks) {
      if (breakTime.time) {
        const breakMinutes = this.timeToMinutes(breakTime.time);
        if (breakMinutes <= startTime || breakMinutes >= endTime) {
          this.showFieldError(breakTime.field, `${breakTime.name} must be between shift start and end times`);
          return false;
        }
      }
    }
    return true;
  }

  timeToMinutes(timeString) {
    const [hours, minutes] = timeString.split(':').map(Number);
    return hours * 60 + minutes;
  }

  isDuplicateEmployee(newEmployee) {
    return this.employees.some(emp => 
      emp.name.toLowerCase() === newEmployee.name.toLowerCase() ||
      (newEmployee.id && emp.id && emp.id.toLowerCase() === newEmployee.id.toLowerCase()) ||
      (newEmployee.email && emp.email && emp.email.toLowerCase() === newEmployee.email.toLowerCase())
    );
  }

  showFieldError(fieldId, message) {
    const field = document.getElementById(fieldId);
    const formGroup = field?.closest('.form-group');
    if (formGroup) {
      formGroup.classList.add('error');
      let errorElement = formGroup.querySelector('.error-message');
      if (!errorElement) {
        errorElement = document.createElement('div');
        errorElement.className = 'error-message';
        formGroup.appendChild(errorElement);
      }
      errorElement.textContent = message;
      errorElement.classList.add('show');
    }
  }

  clearFieldError(fieldId) {
    const field = document.getElementById(fieldId);
    const formGroup = field?.closest('.form-group');
    if (formGroup) {
      formGroup.classList.remove('error');
      const errorElement = formGroup.querySelector('.error-message');
      if (errorElement) {
        errorElement.classList.remove('show');
      }
    }
  }

  validateField(fieldId) {
    const field = document.getElementById(fieldId);
    if (!field) return;
    const value = field.value.trim();
    switch (fieldId) {
      case 'employeeName':
        const nameValidation = validateEmployeeName(value);
        if (!nameValidation.isValid) {
          this.showFieldError(fieldId, nameValidation.error);
        }
        break;
      case 'employeeEmail':
        const emailValidation = validateEmployeeEmail(value);
        if (!emailValidation.isValid) {
          this.showFieldError(fieldId, emailValidation.error);
        }
        break;
      case 'employeeTitle':
        const titleValidation = validateEmployeeTitle(value);
        if (!titleValidation.isValid) {
          this.showFieldError(fieldId, titleValidation.error);
        }
        break;
      case 'employeePhone':
        const phoneValidation = validateEmployeePhone(value);
        if (!phoneValidation.isValid) {
          this.showFieldError(fieldId, phoneValidation.error);
        }
        break;
    }
  }

  updateEmployeePreview() {
    const previewContainer = document.querySelector('.employee-preview');
    if (!previewContainer) return;
    const previewHTML = renderEmployeePreview(this.employees);
    const statsHTML = renderEmployeeStats(this.employees);
    previewContainer.innerHTML = previewHTML + statsHTML;
  }

  renderEmployeeList() {
    const employeeList = document.getElementById('employeeList');
    const employeeCount = document.getElementById('employeeCount');
    if (!employeeList) return;
    if (this.employees.length === 0) {
      employeeList.innerHTML = `
        <div class="empty-state">
          <p>No employees added yet. Add your first employee above!</p>
        </div>
      `;
    } else {
      employeeList.innerHTML = this.employees.map((emp, index) => `
        <div class="employee-card">
          <div class="employee-info">
            <div class="employee-color-dot" style="background-color: ${emp.color}"></div>
            <div class="employee-details">
              <h4>${emp.name}${emp.id ? ` (${emp.id})` : ''}</h4>
              <p>${emp.title} • ${this.formatShiftTime(emp.shifts.start)} - ${this.formatShiftTime(emp.shifts.end)}</p>
              ${emp.email ? `<p>📧 ${emp.email}</p>` : ''}
            </div>
          </div>
          <div class="employee-actions">
            <button class="edit-btn" onclick="attendanceTracker.editEmployee(${index})">Edit</button>
            <button class="delete-btn" onclick="attendanceTracker.deleteEmployee(${index})">Delete</button>
          </div>
        </div>
      `).join('');
    }
    if (employeeCount) {
      employeeCount.textContent = this.employees.length;
    }
    this.updateEmployeePreview();
  }

  formatShiftTime(time) {
    if (!time) return '';
    const [hours, minutes] = time.split(':');
    const hour = parseInt(hours);
    const ampm = hour >= 12 ? 'PM' : 'AM';
    const displayHour = hour % 12 || 12;
    return `${displayHour}:${minutes} ${ampm}`;
  }

  deleteEmployee(index) {
    const employee = this.employees[index];
    this.showToast(
      `${employee.name} deleted successfully`,
      'success'
    );
    this.employees.splice(index, 1);
    this.saveEmployees();
    this.renderEmployeeList();
    this.updateGenerateButton();
  }

  editEmployee(index) {
    const employee = this.employees[index];
    this.editingIndex = index;
    document.getElementById('employeeName').value = employee.name;
    document.getElementById('employeeId').value = employee.id || '';
    document.getElementById('employeeEmail').value = employee.email || '';
    document.getElementById('employeeTitle').value = employee.title;
    document.getElementById('employeePhone').value = employee.phone || '';
    document.getElementById('employeeColor').value = employee.color;
    this.updateColorIndicator(employee.color);
    document.getElementById('shiftStart').value = employee.shifts.start;
    document.getElementById('firstBreak').value = employee.shifts.firstBreak || '';
    document.getElementById('lunchBreak').value = employee.shifts.lunch || '';
    document.getElementById('secondBreak').value = employee.shifts.secondBreak || '';
    document.getElementById('shiftEnd').value = employee.shifts.end;
    this.updateButtonText('.add-employee-btn', 'Update Employee');
    this.resetPresetButtons();
    document.querySelector('.employee-form-section').scrollIntoView({ behavior: 'smooth' });
    this.showToast('info', `Editing ${employee.name} - make changes and click Update Employee`);
  }

  clearForm() {
    const form = document.getElementById('employeeForm');
    if (form) {
      form.reset();
      const colorPicker = document.getElementById('employeeColor');
      if (colorPicker) colorPicker.value = '#4CAF50';
      this.updateColorIndicator('#4CAF50');
      form.querySelectorAll('.form-group').forEach(group => {
        group.classList.remove('error');
        const errorMsg = group.querySelector('.error-message');
        if (errorMsg) errorMsg.classList.remove('show');
      });
      this.editingIndex = -1;
      this.updateButtonText('.add-employee-btn', 'Add Employee');
      this.resetPresetButtons();
    }
  }

  clearAllEmployees() {
    if (arguments[0] !== 'silent') {
      this.showToast(
        `All employees and saved data cleared`,
        'success'
      );
    }
    this.employees = [];
    this.clearSavedData();
    this.renderEmployeeList();
    this.updateGenerateButton();
  }

  clearSavedData() {
    try {
      sessionStorage.removeItem('employeeShiftData');
    } catch (error) {
      // Optionally handle error
    }
  }

  updateGenerateButton() {
    const generateBtn = document.getElementById('generateTrackerBtn');
    if (generateBtn) {
      generateBtn.disabled = this.employees.length === 0;
    }
  }

  showAlert(type, message) {
    this.showToast(type, message);
  }

  showToast(type, message) {
    const toast = document.createElement('div');
    toast.className = `modern-toast modern-toast-${type}`;
    toast.innerHTML = `
      <div class="toast-content">
        <span class="toast-icon">${this.getToastIcon(type)}</span>
        <span class="toast-message">${message}</span>
        <button class="toast-close" onclick="this.parentElement.parentElement.remove()">×</button>
      </div>
    `;
    document.body.appendChild(toast);
    setTimeout(() => toast.classList.add('show'), 10);
    setTimeout(() => {
      if (toast.parentElement) {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
      }
    }, 5000);
  }

  getToastIcon(type) {
    const icons = {
      'success': '✅',
      'error': '❌',
      'warning': '⚠️',
      'info': 'ℹ️'
    };
    return icons[type] || 'ℹ️';
  }

  saveEmployees() {
    try {
      sessionStorage.setItem('employeeShiftData', JSON.stringify(this.employees));
    } catch (error) {
      this.showToast('error', 'Failed to save employee data');
    }
  }

  loadEmployees() {
    try {
      const saved = sessionStorage.getItem('employeeShiftData');
      this.employees = saved ? JSON.parse(saved) : [];
    } catch (error) {
      this.employees = [];
    }
  }

  exportEmployeeList() {
    if (this.employees.length === 0) {
      this.showAlert('warning', 'No employees to export');
      return;
    }
    const headers = ['Name', 'ID', 'Email', 'Title', 'Phone', 'Shift Start', 'First Break', 'Lunch', 'Second Break', 'Shift End', 'Color'];
    const csvContent = [
      headers.join(','),
      ...this.employees.map(emp => [
        `"${emp.name}"`,
        `"${emp.id}"`,
        `"${emp.email}"`,
        `"${emp.title}"`,
        `"${emp.phone}"`,
        emp.shifts.start,
        emp.shifts.firstBreak,
        emp.shifts.lunch,
        emp.shifts.secondBreak,
        emp.shifts.end,
        emp.color
      ].join(','))
    ].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `employee_shift_list_${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    this.showAlert('success', 'Employee list exported successfully');
  }

  handleDragOver(event) {
    event.preventDefault();
    event.currentTarget.classList.add('dragover');
  }

  handleFileDrop(event) {
    event.preventDefault();
    event.currentTarget.classList.remove('dragover');
    const files = event.dataTransfer.files;
    if (files.length > 0) {
      this.processUploadedFile(files[0]);
    }
  }

  handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
      this.processUploadedFile(file);
    }
  }

  async processUploadedFile(file) {
    this.showAlert('info', 'Bulk upload feature will be implemented in the next update');
  }

  async generateShiftTracker() {
    if (this.employees.length === 0) {
      this.showToast('warning', 'No employees to generate tracker for');
      return;
    }
    try {
      const attendanceModule = await import('../../excel/generators/attendance/attendanceTrackerSheet.js');
      const referenceModule = await import('../../excel/generators/attendance/referenceSheet.js');
      const instructionsModule = await import('../../excel/generators/attendance/instructionsSheet.js');
      const stylesModule = await import('../../excel/generators/attendance/stylesXml.js');
      const workbookModule = await import('../../excel/generators/attendance/workbookXml.js');
      const contentTypesModule = await import('../../excel/generators/attendance/contentTypesXml.js');
      const zipModule = await import('../../excel/utils/zipUtils.js');
      const shiftTrackerSheet = attendanceModule.buildShiftTrackerSheet(this.employees);
      const referenceSheet = referenceModule.buildReferenceSheet(this.employees);
      const instructionsSheet = instructionsModule.buildInstructionsSheet();
      const stylesXML = stylesModule.getShiftTrackerStylesXML();
      const workbookXML = workbookModule.getShiftTrackerWorkbookXML();
      const workbookRelsXML = workbookModule.getShiftTrackerWorkbookRelsXML();
      const contentTypesXML = contentTypesModule.getShiftTrackerContentTypesXML();
      const relsXML = contentTypesModule.getShiftTrackerRelsXML();
      const files = [
        { name: 'xl/worksheets/sheet1.xml', content: shiftTrackerSheet },
        { name: 'xl/worksheets/sheet2.xml', content: referenceSheet },
        { name: 'xl/worksheets/sheet3.xml', content: instructionsSheet },
        { name: 'xl/styles.xml', content: stylesXML },
        { name: 'xl/workbook.xml', content: workbookXML },
        { name: 'xl/_rels/workbook.xml.rels', content: workbookRelsXML },
        { name: '[Content_Types].xml', content: contentTypesXML },
        { name: '_rels/.rels', content: relsXML }
      ];
      const zipBuffer = zipModule.createZip(files);
      this.downloadExcelFile(zipBuffer);
      this.showToast('success', 'Shift tracker Excel file generated successfully!');
    } catch (error) {
      this.showToast('error', 'Failed to generate Excel file. Please try again.');
    }
  }

  downloadExcelFile(zipBuffer) {
    const blob = new Blob([zipBuffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `shift_tracker_${new Date().toISOString().split('T')[0]}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  init() {
    this.renderEmployeeList();
    this.updateGenerateButton();
    window.clearEmployeeData = () => {
      this.employees = [];
      this.clearSavedData();
      this.renderEmployeeList();
      this.updateGenerateButton();
      location.reload();
    };
    if (this.employees.length > 0) {
      console.log('💡 Developer tip: Run clearEmployeeData() in console to clear all saved employee data');
    }
  }
}

let attendanceTracker;

export function setupAttendanceTracker() {
  attendanceTracker = new AttendanceTracker();
  attendanceTracker.init();
  window.attendanceTracker = attendanceTracker;
}
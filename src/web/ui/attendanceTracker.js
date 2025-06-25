// Employee Shift Tracker Module
// Handles employee management, validation, and shift tracking functionality

import { sanitizeEmployeeName, sanitizeEmployeeText, sanitizeEmail, sanitizePhoneNumber } from '../../shared/utils/sanitize.js';
import { validateEmployeeName, validateEmployeeEmail, validateEmployeeTitle, validateEmployeePhone, validateShiftTime, validateBreakTime } from '../../shared/utils/validation.js';
import { renderEmployeePreview, renderEmployeeStats } from '../utils/previewEmployee.js';
import { buildLegendSheet } from '../../excel/generators/attendance/legendSheet.js';
import { renderCalendarPreview as importedRenderCalendarPreview } from '../utils/previewCalendar.js';
import { getMonthName, getYear } from '../../shared/utils/timeUtils.js';

export class AttendanceTracker {
  constructor() {
    this.employees = [];
    this.maxEmployees = 100;
    this.editingIndex = -1;
    this.hasUnsavedChanges = false;
    this.lastImportIssues = null;
    this.toastQueue = []; // Add toast queue
    this.isShowingToast = false; // Track if a toast is currently showing
    this.hasShownDataStorageToast = false; // Track if data storage toast has been shown
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
    
    // Add CSV template download functionality
    const templateBtn = document.getElementById('downloadTemplateBtn');
    if (templateBtn) {
      templateBtn.addEventListener('click', () => this.downloadCSVTemplate());
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
      employeeList.innerHTML = this.employees.map((emp, index) => {
        // Determine if employee needs attention (missing critical info)
        const needsAttention = emp._needsAttention || 
          (!emp.title || emp.title === 'Employee') || 
          (!emp.email || !this.isValidEmail(emp.email));
        
        const cardClass = needsAttention ? 'employee-card needs-attention' : 'employee-card';
        
        return `
        <div class="${cardClass}">
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
      `}).join('');
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
    // Add to queue if a toast is currently showing
    if (this.isShowingToast) {
      this.toastQueue.push({ type, message });
      return;
    }

    this.isShowingToast = true;
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
    
    const hideToast = () => {
      if (toast.parentElement) {
        toast.classList.remove('show');
        setTimeout(() => {
          toast.remove();
          this.isShowingToast = false;
          // Show next toast in queue if any
          if (this.toastQueue.length > 0) {
            const next = this.toastQueue.shift();
            setTimeout(() => this.showToast(next.type, next.message), 100);
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

  downloadCSVTemplate() {
    const csvContent = this.generateCSVTemplate();
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', 'employee_import_template.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    this.showToast('success', 'CSV template downloaded! Fill it out and upload to import employees.');
  }

  generateCSVTemplate() {
    const headers = [
      'Name', 'ID', 'Email', 'Title', 'Phone', 
      'Shift Start', 'First Break', 'Lunch Break', 'Second Break', 'Shift End'
    ];
    
    let csv = headers.join(',') + '\n';
    
    // Add sample data rows with instructions
    const sampleData = [
      ['John Doe', 'EMP001', 'john.doe@company.com', 'Developer', '555-1234', '09:00', '10:30', '12:00', '15:00', '17:00'],
      ['Jane Smith', 'EMP002', 'jane.smith@company.com', 'Manager', '555-5678', '08:30', '10:00', '12:30', '15:30', '17:30'],
      ['Bob Johnson', 'EMP003', 'bob.johnson@company.com', 'Analyst', '555-9999', '09:30', '11:00', '13:00', '16:00', '18:00']
    ];
    
    sampleData.forEach(row => {
      csv += row.map(cell => `"${cell}"`).join(',') + '\n';
    });
    
    // Add instructions as comments
    csv += '\n# Instructions:\n';
    csv += '# 1. Replace the sample data above with your actual employee information\n';
    csv += '# 2. Name and Title are required fields\n';
    csv += '# 3. Times should be in 24-hour format (HH:MM) or 12-hour format (H:MM AM/PM)\n';
    csv += '# 4. Break times are optional - leave blank if not applicable\n';
    csv += '# 5. Save this file and upload it using the "Choose File" button\n';
    csv += '# 6. Delete these instruction lines before uploading\n';
    
    return csv;
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
    if (!file) return;
    
    const fileName = file.name.toLowerCase();
    const isCSV = fileName.endsWith('.csv');
    const isExcel = fileName.endsWith('.xlsx') || fileName.endsWith('.xls');
    
    if (!isCSV && !isExcel) {
      this.showToast('error', 'Please upload a CSV or Excel file (.csv, .xlsx, .xls)');
      return;
    }
    
    this.showToast('info', 'Processing file...');
    
    try {
      let employeeData = [];
      
      if (isCSV) {
        employeeData = await this.parseCSVFile(file);
      } else if (isExcel) {
        employeeData = await this.parseExcelFile(file);
      }
      
      if (employeeData.length === 0) {
        this.showToast('warning', 'No valid employee data found in the file');
        return;
      }
      
      // Validate and import employees
      const importResults = await this.importEmployees(employeeData);
      
      if (importResults.successful > 0) {
        this.showToast('success', `Successfully imported ${importResults.successful} employees`);
        if (importResults.failed > 0) {
          this.showToast('warning', `${importResults.failed} employees failed validation and were skipped`);
        }
      } else {
        this.showToast('error', 'No employees could be imported. Please check the file format and data.');
      }
      
    } catch (error) {
      console.error('File processing error:', error);
      this.showToast('error', 'Failed to process file. Please check the file format and try again.');
    }
  }

  async parseCSVFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const csv = e.target.result;
          const lines = csv.split('\n').map(line => line.trim()).filter(line => line);
          
          if (lines.length < 2) {
            resolve([]);
            return;
          }
          
          // Parse header row to find column indices
          const headers = this.parseCSVLine(lines[0]);
          console.log('DEBUG: CSV headers found:', headers);
          const columnMap = this.mapCSVColumns(headers);
          console.log('DEBUG: Column mapping:', columnMap);
          
          const employees = [];
          for (let i = 1; i < lines.length; i++) {
            const values = this.parseCSVLine(lines[i]);
            if (values.length > 0) {
              const employee = this.extractEmployeeFromCSV(values, columnMap);
              if (employee && employee.name) {
                employees.push(employee);
              }
            }
          }
          
          console.log(`DEBUG: Extracted ${employees.length} employee records from CSV`);
          resolve(employees);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = () => reject(new Error('Failed to read CSV file'));
      reader.readAsText(file);
    });
  }

  async parseExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          // Use a lightweight Excel parsing approach
          // For now, suggest users export Excel to CSV for import
          this.showToast('info', 'For Excel files, please save as CSV and re-upload. Full Excel support coming soon!');
          resolve([]);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = () => reject(new Error('Failed to read Excel file'));
      reader.readAsArrayBuffer(file);
    });
  }

  parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      
      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i++; // Skip next quote
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === ',' && !inQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    
    result.push(current.trim());
    return result;
  }

  mapCSVColumns(headers) {
    const map = {};
    headers.forEach((header, index) => {
      const lower = header.toLowerCase().trim();
      
      // Map common column names to our fields - enhanced for real HR data
      if (lower.includes('employee name') || lower.includes('name') || lower === 'employee' || lower === 'full name') {
        map.name = index;
      } else if (lower.includes('employee id') || lower.includes('emp id') || lower.includes('id') || lower === 'emp_id' || lower === 'employee_id') {
        map.id = index;
      } else if (lower.includes('email address') || lower.includes('email') || lower === 'e-mail' || lower.includes('e_mail')) {
        map.email = index;
      } else if (lower.includes('job title') || lower.includes('title') || lower.includes('position') || lower.includes('job')) {
        map.title = index;
      } else if (lower.includes('phone number') || lower.includes('phone') || lower.includes('tel') || lower.includes('mobile')) {
        map.phone = index;
      } else if ((lower.includes('shift start') || lower.includes('start time')) && !lower.includes('break')) {
        map.shiftStart = index;
      } else if ((lower.includes('shift end') || lower.includes('end time')) && !lower.includes('break')) {
        map.shiftEnd = index;
      } else if (lower.includes('first break') || (lower.includes('first') && lower.includes('break'))) {
        map.firstBreak = index;
      } else if (lower.includes('lunch break') || lower.includes('lunch time') || (lower.includes('lunch') && lower.includes('break'))) {
        map.lunch = index;
      } else if (lower.includes('second break') || (lower.includes('second') && lower.includes('break'))) {
        map.secondBreak = index;
      }
    });
    
    return map;
  }

  extractEmployeeFromCSV(values, columnMap) {
    const employee = {
      name: '',
      id: '',
      email: '',
      title: '',
      phone: '',
      color: '#4CAF50',
      shifts: {
        start: '09:00',
        firstBreak: '',
        lunch: '',
        secondBreak: '',
        end: '17:00'
      }
    };

    // Extract values based on column mapping
    if (columnMap.name !== undefined) employee.name = values[columnMap.name] || '';
    if (columnMap.id !== undefined) employee.id = values[columnMap.id] || '';
    if (columnMap.email !== undefined) employee.email = values[columnMap.email] || '';
    if (columnMap.title !== undefined) employee.title = values[columnMap.title] || '';
    if (columnMap.phone !== undefined) employee.phone = values[columnMap.phone] || '';
    if (columnMap.shiftStart !== undefined) employee.shifts.start = this.normalizeTimeFormat(values[columnMap.shiftStart]) || '09:00';
    if (columnMap.shiftEnd !== undefined) employee.shifts.end = this.normalizeTimeFormat(values[columnMap.shiftEnd]) || '17:00';
    if (columnMap.firstBreak !== undefined) employee.shifts.firstBreak = this.normalizeTimeFormat(values[columnMap.firstBreak]) || '';
    if (columnMap.lunch !== undefined) employee.shifts.lunch = this.normalizeTimeFormat(values[columnMap.lunch]) || '';
    if (columnMap.secondBreak !== undefined) employee.shifts.secondBreak = this.normalizeTimeFormat(values[columnMap.secondBreak]) || '';

    return employee;
  }

  normalizeTimeFormat(timeStr) {
    if (!timeStr) return '';
    
    // Remove whitespace and convert to string
    timeStr = String(timeStr).trim();
    
    // Handle various time formats
    // 9:00 AM, 9:00am, 09:00, 9:00, etc.
    const timeMatch = timeStr.match(/(\d{1,2}):?(\d{0,2})\s*(am|pm)?/i);
    
    if (!timeMatch) return '';
    
    let hours = parseInt(timeMatch[1]);
    let minutes = timeMatch[2] ? parseInt(timeMatch[2]) : 0;
    const ampm = timeMatch[3] ? timeMatch[3].toLowerCase() : null;
    
    // Convert to 24-hour format
    if (ampm === 'pm' && hours !== 12) {
      hours += 12;
    } else if (ampm === 'am' && hours === 12) {
      hours = 0;
    }
    
    // Validate time
    if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
      return '';
    }
    
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  }

  async importEmployees(employeeData) {
    let successful = 0;
    let failed = 0;
    let partialImports = 0;
    const failureReasons = [];
    const completelyRejectedEmployees = []; // Only track completely rejected employees for report
    
    console.log(`DEBUG: Attempting to import ${employeeData.length} employees`);
    
    // Check if we have any parseable data at all
    const parseableEmployees = employeeData.filter(emp => emp.name && emp.name.trim().length > 0);
    
    if (parseableEmployees.length === 0) {
      this.showToast('error', 'No valid employee data found. CSV must contain at least employee names to import.');
      return { successful: 0, failed: employeeData.length, partialImports: 0 };
    }
    
    for (let i = 0; i < employeeData.length; i++) {
      const employeeInfo = employeeData[i];
      const rowNumber = i + 2; // CSV row number (accounting for header)
      let rowIssues = [];
      
      try {
        // Check if we're at capacity
        if (this.employees.length >= this.maxEmployees) {
          this.showToast('warning', `Maximum capacity of ${this.maxEmployees} employees reached. Remaining imports skipped.`);
          break;
        }
        
        // Minimum requirement: just need a name
        if (!employeeInfo.name || !employeeInfo.name.trim()) {
          failed++;
          failureReasons.push(`Row ${rowNumber}: Missing employee name (required)`);
          rowIssues.push('Missing employee name (required)');
          completelyRejectedEmployees.push({ row: rowNumber, employee: employeeInfo, issues: rowIssues });
          console.log(`DEBUG: Failed - no name in row ${rowNumber}`);
          continue;
        }
        
        // Check for duplicates
        if (this.isDuplicateEmployee(employeeInfo)) {
          failed++;
          failureReasons.push(`Row ${rowNumber} (${employeeInfo.name}): Duplicate employee`);
          rowIssues.push('Duplicate employee name');
          completelyRejectedEmployees.push({ row: rowNumber, employee: employeeInfo, issues: rowIssues });
          console.log(`DEBUG: Duplicate employee: ${employeeInfo.name}`);
          continue;
        }
        
        // Sanitize data and track any issues
        const sanitizedEmployee = this.sanitizeEmployeeDataWithIssues(employeeInfo, rowIssues);
        
        // Check if this is a partial import (missing title or other important data)
        let isPartialImport = false;
        if (!sanitizedEmployee.title || !sanitizedEmployee.title.trim()) {
          isPartialImport = true;
        }
        if (!sanitizedEmployee.email || !this.isValidEmail(sanitizedEmployee.email)) {
          isPartialImport = true;
        }
        // Note: Phone number is optional and doesn't count as partial import
        
        // Add default values for missing fields
        if (!sanitizedEmployee.title) sanitizedEmployee.title = 'Employee';
        if (!sanitizedEmployee.email) sanitizedEmployee.email = '';
        if (!sanitizedEmployee.phone) sanitizedEmployee.phone = '';
        
        // Mark as needing attention for visual highlighting
        if (isPartialImport) {
          sanitizedEmployee._needsAttention = true;
          partialImports++;
        }
        
        // Add to employees list
        this.employees.push(sanitizedEmployee);
        successful++;
        console.log(`DEBUG: Successfully imported: ${sanitizedEmployee.name}${isPartialImport ? ' (needs attention)' : ''}`);
        
      } catch (error) {
        console.error('Error importing employee:', employeeInfo, error);
        failureReasons.push(`Row ${rowNumber} (${employeeInfo.name || 'Unknown'}): ${error.message}`);
        rowIssues.push(`Import error: ${error.message}`);
        completelyRejectedEmployees.push({ row: rowNumber, employee: employeeInfo, issues: rowIssues });
        failed++;
      }
    }
    
    // Show detailed results
    if (successful > 0) {
      let message = `Successfully imported ${successful} employee${successful > 1 ? 's' : ''}`;
      if (partialImports > 0) {
        message += ` (${partialImports} highlighted in red need attention - use Edit to complete)`;
      }
      this.showToast('success', message);
    }
    
    if (failed > 0) {
      this.showToast('warning', `${failed} employee${failed > 1 ? 's' : ''} were completely rejected. Click 'Download Import Report' to see details.`);
    }
    
    // Store only completely rejected employees for download report
    if (completelyRejectedEmployees.length > 0) {
      this.lastImportIssues = completelyRejectedEmployees;
      this.showImportReportButton();
    }
    
    // Show detailed failure reasons in console for debugging
    if (failureReasons.length > 0) {
      console.log('DEBUG: Import failure reasons:', failureReasons);
    }
    
    // Update UI
    this.saveEmployees();
    this.renderEmployeeList();
    this.updateGenerateButton();
    
    return { successful, failed, partialImports };
  }

  // Enhanced sanitization that tracks issues
  sanitizeEmployeeDataWithIssues(data, issues = []) {
    const sanitized = this.sanitizeEmployeeData(data);
    
    // Track time format issues
    if (data.shifts) {
      if (data.shifts.start && !this.normalizeTimeFormat(data.shifts.start)) {
        issues.push(`Invalid start time format: "${data.shifts.start}"`);
      }
      if (data.shifts.end && !this.normalizeTimeFormat(data.shifts.end)) {
        issues.push(`Invalid end time format: "${data.shifts.end}"`);
      }
      if (data.shifts.firstBreak && !this.normalizeTimeFormat(data.shifts.firstBreak)) {
        issues.push(`Invalid first break time: "${data.shifts.firstBreak}"`);
      }
      if (data.shifts.lunch && !this.normalizeTimeFormat(data.shifts.lunch)) {
        issues.push(`Invalid lunch time: "${data.shifts.lunch}"`);
      }
      if (data.shifts.secondBreak && !this.normalizeTimeFormat(data.shifts.secondBreak)) {
        issues.push(`Invalid second break time: "${data.shifts.secondBreak}"`);
      }
    }
    
    return sanitized;
  }

  // Helper to validate email format
  isValidEmail(email) {
    if (!email) return false;
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email.trim());
  }

  // Show the import report download button
  showImportReportButton() {
    // Remove any existing report button
    const existingBtn = document.getElementById('downloadImportReportBtn');
    if (existingBtn) {
      existingBtn.remove();
    }

    // Create the download report button with descriptive text and centered styling
    const reportBtn = document.createElement('button');
    reportBtn.id = 'downloadImportReportBtn';
    reportBtn.className = 'btn btn-secondary';
    reportBtn.innerHTML = `
      <div style="display: flex; flex-direction: column; align-items: center; gap: 4px;">
        <div style="display: flex; align-items: center; gap: 8px;">
          <span>📋</span>
          <span>Download Import Report</span>
        </div>
        <div style="font-size: 11px; opacity: 0.9; font-weight: normal;">
          See what was completely rejected and why
        </div>
      </div>
    `;
    reportBtn.style.cssText = `
      margin: 10px auto;
      display: block;
      padding: 12px 20px;
      background: #6c757d;
      border: none;
      color: white;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
      transition: all 0.3s ease;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      min-height: 60px;
      text-align: center;
    `;
    
    reportBtn.addEventListener('click', () => this.downloadImportReport());
    
    // Add hover effects
    reportBtn.addEventListener('mouseenter', () => {
      reportBtn.style.background = '#5a6268';
      reportBtn.style.transform = 'translateY(-1px)';
      reportBtn.style.boxShadow = '0 4px 8px rgba(0,0,0,0.15)';
    });
    
    reportBtn.addEventListener('mouseleave', () => {
      reportBtn.style.background = '#6c757d';
      reportBtn.style.transform = 'translateY(0)';
      reportBtn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
    });

    // Insert after the upload area with centering
    const uploadArea = document.querySelector('.upload-area');
    if (uploadArea && uploadArea.parentNode) {
      uploadArea.parentNode.insertBefore(reportBtn, uploadArea.nextSibling);
    } else {
      // Fallback: add to employee list container
      const employeeList = document.getElementById('employeeList');
      if (employeeList && employeeList.parentNode) {
        employeeList.parentNode.insertBefore(reportBtn, employeeList);
      }
    }

    // Auto-remove after 30 seconds with fade effect
    setTimeout(() => {
      if (reportBtn.parentElement) {
        reportBtn.style.opacity = '0';
        reportBtn.style.transform = 'translateY(-10px)';
        setTimeout(() => reportBtn.remove(), 300);
      }
    }, 30000);
  }

  // Download annotated CSV with import issues - only completely rejected employees
  downloadImportReport() {
    if (!this.lastImportIssues || this.lastImportIssues.length === 0) {
      this.showToast('info', 'No completely rejected employees to report.');
      return;
    }
    
    // Create CSV content with issues highlighted
    let csvContent = 'Row,Employee Name,Job Title,Email,Rejection Reason\n';
    
    this.lastImportIssues.forEach(issue => {
      const employee = issue.employee;
      const issuesText = issue.issues.join('; ');
      
      // Escape CSV values
      const name = this.escapeCsvValue(employee.name || '');
      const title = this.escapeCsvValue(employee.title || '');
      const email = this.escapeCsvValue(employee.email || '');
      const issues = this.escapeCsvValue(issuesText);
      
      csvContent += `${issue.row},${name},${title},${email},${issues}\n`;
    });
    
    // Download the file
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `rejected_employees_${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    this.showToast('success', 'Import report downloaded successfully!');
  }

  // Helper to escape CSV values
  escapeCsvValue(value) {
    if (!value) return '';
    const stringValue = String(value);
    if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
      return `"${stringValue.replace(/"/g, '""')}"`;
    }
    return stringValue;
  }

  getLegendValues() {
    return Array.from(document.querySelectorAll('.legend-input')).map(input => input.value.trim());
  }

  getLegendColors() {
    return Array.from(document.querySelectorAll('.legend-color-picker')).map(picker => picker.value);
  }

  async generateShiftTracker() {
    try {
      // Import Excel generation modules
      const { buildShiftTrackerSheet } = await import('../../excel/generators/attendance/attendanceTrackerSheet.js');
      const { buildReferenceSheet } = await import('../../excel/generators/attendance/referenceSheet.js');
      const { buildInstructionsSheet } = await import('../../excel/generators/attendance/instructionsSheet.js');
      const { getShiftTrackerStylesXML } = await import('../../excel/generators/attendance/stylesXml.js');
      const { getShiftTrackerWorkbookXML, getShiftTrackerWorkbookRelsXML } = await import('../../excel/generators/attendance/workbookXml.js');
      const { getShiftTrackerContentTypesXML, getShiftTrackerRelsXML } = await import('../../excel/generators/attendance/contentTypesXml.js');
      const { createZip } = await import('../../excel/utils/zipUtils.js');
      const { buildLegendSheet } = await import('../../excel/generators/attendance/legendSheet.js');

      // Get date configuration
      const yearInput = document.getElementById('yearInput');
      const year = yearInput ? parseInt(yearInput.value, 10) : (window.attendanceDateSelection?.year || new Date().getFullYear());
      const monthSelect = document.getElementById('monthSelect');
      const month = monthSelect ? parseInt(monthSelect.value, 10) : (window.attendanceDateSelection?.month || new Date().getMonth());
      const eventRowsSelect = document.getElementById('eventRowsSelect');
      const eventRows = eventRowsSelect ? parseInt(eventRowsSelect.value, 10) : 1;
      
      // Build array of selected dates
      let selectedDates = window.attendanceDateSelection?.selectedDates || [];
      if (!selectedDates || selectedDates.length === 0) {
        selectedDates = [];
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        for (let d = 1; d <= daysInMonth; d++) {
          selectedDates.push(`${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`);
        }
      }
      
      // Get legend configuration
      let legendObjs = [];
      if (window.attendanceDateSelection?.legendValues) {
        legendObjs = window.attendanceDateSelection.legendValues;
      } else {
        const legendValues = this.getLegendValues();
        const legendColors = this.getLegendColors();
        legendObjs = legendValues.map((label, i) => ({ label, color: legendColors[i] }));
      }
      
      // Build Excel color format
      const legendColorsExcel = legendObjs.map(obj => {
        const cleanHex = obj.color.replace('#', '').toUpperCase();
        return 'FF' + cleanHex;
      });

      // Calculate sheet name
      let shiftTrackerSheetName = "Shift Tracker";
      if (selectedDates && selectedDates.length > 0) {
        const month = getMonthName(selectedDates[0]);
        const year = getYear(selectedDates[0]);
        shiftTrackerSheetName = `Shift Tracker ${month} ${year}`;
      }

      // Generate Excel sheets
      const instructionsSheet = buildInstructionsSheet();
      const shiftTrackerSheet = buildShiftTrackerSheet(this.employees, selectedDates, legendObjs);
      const referenceSheet = buildReferenceSheet(this.employees, legendObjs, shiftTrackerSheetName);
      const legendSheet = buildLegendSheet(legendObjs);
      const stylesXML = getShiftTrackerStylesXML(legendColorsExcel);
      const workbookXML = getShiftTrackerWorkbookXML(shiftTrackerSheetName);
      const workbookRelsXML = getShiftTrackerWorkbookRelsXML();
      const contentTypesXML = getShiftTrackerContentTypesXML();
      const relsXML = getShiftTrackerRelsXML();

      // Create files array
      const files = [
        { name: 'xl/worksheets/sheet1.xml', content: instructionsSheet },
        { name: 'xl/worksheets/sheet2.xml', content: shiftTrackerSheet },
        { name: 'xl/worksheets/sheet3.xml', content: referenceSheet },
        { name: 'xl/worksheets/sheet4.xml', content: legendSheet },
        { name: 'xl/styles.xml', content: stylesXML },
        { name: 'xl/workbook.xml', content: workbookXML },
        { name: 'xl/_rels/workbook.xml.rels', content: workbookRelsXML },
        { name: '[Content_Types].xml', content: contentTypesXML },
        { name: '_rels/.rels', content: relsXML }
      ];

      // Create and download Excel file
      try {
        const zipBuffer = createZip(files);
        this.downloadExcelFile(zipBuffer);
        this.showToast('success', 'Shift tracker Excel file generated successfully!');
      } catch (zipError) {
        console.error('❌ ZIP creation failed:', zipError);
        throw new Error(`ZIP creation failed: ${zipError.message}`);
      }
    } catch (error) {
      console.error('💥 Excel generation failed:', error.message);
      console.error('Error stack:', error.stack);
      
      // Show error message to user
      const isDev = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
      const errorMessage = isDev 
        ? `Failed to generate Excel file: ${error.message}` 
        : 'Failed to generate Excel file. Please check the console for details and try refreshing the page.';
      
      this.showToast('error', errorMessage);
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

  // Show the data storage toast once when user enters the employee section
  showDataStorageToastOnce() {
    if (!this.hasShownDataStorageToast) {
      this.hasShownDataStorageToast = true;
      setTimeout(() => {
        this.showToast('info', 'Note: Employee data is stored for this browser session only and will be cleared when you close this tab.');
      }, 1000);
    }
  }

  // Static method for calendar preview rendering (for compatibility)
  static renderCalendarPreview({ year, month, eventRows, legendValues, legendColors }) {
    // Simple fallback that returns empty string since preview is not essential during Excel generation
    return '<div class="calendar-preview-placeholder">Calendar preview skipped during Excel generation</div>';
  }
}

let attendanceTracker;

export function setupAttendanceTracker() {
  attendanceTracker = new AttendanceTracker();
  attendanceTracker.renderEmployeeList();
  attendanceTracker.updateGenerateButton();
  
  // Setup the clearEmployeeData function for debugging
  window.clearEmployeeData = () => {
    attendanceTracker.employees = [];
    attendanceTracker.clearSavedData();
    attendanceTracker.renderEmployeeList();
    attendanceTracker.updateGenerateButton();
    location.reload();
  };
  
  window.attendanceTracker = attendanceTracker;
  
  // Show the data storage toast once when entering the employee section
  attendanceTracker.showDataStorageToastOnce();
}

export function setupAttendanceBuilderPage() {
  const mainContent = document.querySelector('.main-content');
  if (!mainContent) return;

  // Wait for user to click Start Building
  const startBtn = document.getElementById('startAttendanceBuilderBtn');
  
  if (startBtn) {
    startBtn.addEventListener('click', () => {
      // Show the date selection page first
      const template = document.getElementById('attendanceDateSelectionTemplate');
      if (template) {
        mainContent.innerHTML = template.innerHTML;
        setTimeout(() => {
          if (!window.attendanceTracker) {
            window.attendanceTracker = new AttendanceTracker();
          }
          setupAttendanceDateSelection(window.attendanceTracker);
        }, 0);
      }
    });
  } else {
    // Try again after a short delay
    setTimeout(() => {
      const delayedBtn = document.getElementById('startAttendanceBuilderBtn');
      if (delayedBtn) {
        delayedBtn.addEventListener('click', () => {
          const template = document.getElementById('attendanceDateSelectionTemplate');
          if (template) {
            mainContent.innerHTML = template.innerHTML;
            setTimeout(() => {
              if (!window.attendanceTracker) {
                window.attendanceTracker = new AttendanceTracker();
              }
              setupAttendanceDateSelection(window.attendanceTracker);
            }, 0);
          }
        });
      }
    }, 100);
  }
}

export function setupAttendanceDateSelection(tracker) {
  const mainContent = document.querySelector('.main-content');
  if (!mainContent) return;

  // Initialize date selection state
  let selectedMonth = 5; // June (0-indexed)
  let selectedYear = 2025;
  let selectedDateRange = null;
  let selectedDates = [];

  // Legend state
  let legendCount = 8;
  let customLegendCount = 0;
  const maxCustomLegends = 10;

  // Get DOM elements
  const monthSelect = document.getElementById('attendanceMonthSelect');
  const yearSelect = document.getElementById('attendanceYearSelect');
  const dateRangeBtns = document.querySelectorAll('.date-range-btn');
  const dateRangeInfo = document.getElementById('dateRangeInfo');
  const calendarPreview = document.getElementById('attendanceCalendarPreview');
  const continueBtn = document.getElementById('continueToFormBtn');
  const legendFields = document.getElementById('legendFields');
  const addLegendBtn = document.getElementById('addLegendBtn');
  const legendColorPreview = document.getElementById('legendColorPreview');

  // Set initial values
  if (monthSelect) monthSelect.value = selectedMonth;
  if (yearSelect) yearSelect.value = selectedYear;

  // Month/Year change handlers
  if (monthSelect) {
    monthSelect.addEventListener('change', (e) => {
      selectedMonth = parseInt(e.target.value);
      updateCalendarPreview();
      validateDateSelectionToast();
    });
  }

  if (yearSelect) {
    yearSelect.addEventListener('change', (e) => {
      selectedYear = parseInt(e.target.value);
      updateCalendarPreview();
      validateDateSelectionToast();
    });
  }

  // Date range button handlers
  dateRangeBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const rangeType = btn.dataset.range;
      
      // Update button states
      dateRangeBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      
      selectedDateRange = rangeType;
      updateDateSelection();
    });
  });

  // Add custom legend field (max 10 custom legends)
  if (addLegendBtn && legendFields) {
    addLegendBtn.addEventListener('click', (e) => {
      e.preventDefault();
      if (customLegendCount >= maxCustomLegends) {
        tracker.showToast('warning', `Maximum of ${maxCustomLegends} custom legends reached.`);
        return;
      }
      legendCount++;
      customLegendCount++;
      const defaultColor = getDefaultLegendColor(legendCount);
      const row = document.createElement('div');
      row.className = 'legend-field-row';
      row.innerHTML = `<span class="legend-color-dot color-dot-picker" style="background:${defaultColor}; position: relative; cursor: pointer;">
        <input type="color" class="legend-color-picker" value="${defaultColor}" style="opacity:0;position:absolute;left:0;top:0;width:100%;height:100%;cursor:pointer;">
      </span><input type="text" class="legend-input" value="" maxlength="40" placeholder="Custom Legend ${legendCount}">
      <button type="button" class="legend-remove-btn" aria-label="Remove legend">×</button>`;
      legendFields.appendChild(row);
      row.querySelector('input.legend-input').focus();
      setupLegendRowEvents(row);
      // Add remove button event
      const removeBtn = row.querySelector('.legend-remove-btn');
      if (removeBtn) {
        removeBtn.addEventListener('click', () => {
          row.remove();
          updateLegendColorPreview();
          tracker.showToast('info', 'Legend removed.');
        });
      }
      updateLegendColorPreview();
    });
  }

  // When generating initial legend fields, add remove buttons and listeners
  legendFields.querySelectorAll('.legend-field-row').forEach(row => {
    if (!row.querySelector('.legend-remove-btn')) {
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'legend-remove-btn';
      removeBtn.setAttribute('aria-label', 'Remove legend');
      removeBtn.textContent = '×';
      removeBtn.addEventListener('click', () => {
        row.remove();
        updateLegendColorPreview();
        tracker.showToast('info', 'Legend removed.');
      });
      row.appendChild(removeBtn);
    }
    setupLegendRowEvents(row);
  });

  // Continue button handler
  if (continueBtn) {
    continueBtn.addEventListener('click', () => {
      if (selectedDateRange && selectedDates.length > 0) {
        // Collect legend values and colors
        const legendInputs = legendFields.querySelectorAll('.legend-input');
        const colorPickers = legendFields.querySelectorAll('.legend-color-picker');
        const legendValues = [];
        for (let i = 0; i < legendInputs.length; i++) {
          const label = legendInputs[i].value.trim();
          const color = colorPickers[i].value;
          if (label.length > 0) {
            legendValues.push({ label, color });
          }
        }
        window.attendanceDateSelection = {
          month: selectedMonth,
          year: selectedYear,
          dateRangeType: selectedDateRange,
          selectedDates: selectedDates,
          legendValues: legendValues
        };
        const template = document.getElementById('attendancePageTemplate');
        if (template) {
          mainContent.innerHTML = template.innerHTML;
          setTimeout(() => {
            setupAttendanceTracker();
          }, 0);
        }
      }
    });
  }

  function setupLegendRowEvents(row) {
    const colorDot = row.querySelector('.legend-color-dot.color-dot-picker');
    const colorPicker = row.querySelector('.legend-color-picker');
    const input = row.querySelector('.legend-input');
    if (colorDot && colorPicker) {
      colorDot.addEventListener('click', () => {
        colorPicker.click();
      });
      colorPicker.addEventListener('input', () => {
        colorDot.style.background = colorPicker.value;
        updateLegendColorPreview();
      });
    }
    if (input) {
      input.addEventListener('input', updateLegendColorPreview);
    }
  }

  function updateLegendColorPreview() {
    if (!legendColorPreview) return;
    const legendInputs = legendFields.querySelectorAll('.legend-input');
    const colorPickers = legendFields.querySelectorAll('.legend-color-picker');
    let html = '<div class="calendar-legend"><strong>Legend:</strong> ';
    for (let i = 0; i < legendInputs.length; i++) {
      const label = legendInputs[i].value.trim();
      const color = colorPickers[i].value;
      if (label.length > 0) {
        // Abbreviate label for preview if too long or matches a known abbreviation
        let shortLabel = label;
        const abbrevMap = {
          'PTO (Personal Time Off)': 'PTO',
          'VTO (Voluntary Time Off)': 'VTO',
          'FMLA/ADA': 'FMLA',
          'Approved Leave of Absence': 'Leave',
          'Excused Absence': 'Excused',
          'Unexcused Absence': 'Unexcused',
          'Late Arrival': 'Late',
          'Early Departure': 'Early'
        };
        if (abbrevMap[label]) {
          shortLabel = abbrevMap[label];
        } else if (label.length > 16) {
          shortLabel = label.slice(0, 13) + '...';
        }
        html += `<span class="calendar-legend-item" style="background:${color};">${shortLabel}</span> `;
      }
    }
    html += '</div>';
    legendColorPreview.innerHTML = html;
  }

  function getDefaultLegendColor(index) {
    // Cycle through a palette for custom legends
    const palette = [
      '#20b388', '#228B22', '#1E90FF', '#FFA500', '#800080', '#FF6666', '#FFD700', '#4682B4', '#DC143C', '#00CED1', '#8B4513', '#FF69B4'
    ];
    return palette[index % palette.length];
  }

  function updateDateSelection() {
    if (!selectedDateRange) {
      dateRangeInfo.textContent = '';
      continueBtn.disabled = true;
      removePersistentToast();
      return;
    }
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December'];
    const monthName = monthNames[selectedMonth];
    switch (selectedDateRange) {
      case 'full':
        selectedDates = getAllDatesInMonth(selectedYear, selectedMonth);
        dateRangeInfo.textContent = `Full month: ${monthName} ${selectedYear} (${selectedDates.length} days)`;
        break;
      case 'workdays':
        selectedDates = getWorkDaysInMonth(selectedYear, selectedMonth);
        dateRangeInfo.textContent = `Work days only: ${monthName} ${selectedYear} (${selectedDates.length} days)`;
        break;
      case 'custom':
        selectedDates = [];
        dateRangeInfo.textContent = 'Click dates on the calendar to select them';
        break;
    }
    updateCalendarPreview();
    continueBtn.disabled = selectedDates.length === 0;
    validateDateSelectionToast();
  }

  function updateCalendarPreview() {
    if (!calendarPreview) return;
    import('../utils/previewCalendar.js').then(module => {
      const calendarHTML = module.generateCalendarPreview(selectedYear, selectedMonth, selectedDates);
      calendarPreview.innerHTML = calendarHTML;
      if (selectedDateRange === 'custom') {
        const calendarDays = calendarPreview.querySelectorAll('.calendar-date');
        calendarDays.forEach(day => {
          day.addEventListener('click', () => {
            const date = day.dataset.date;
            if (date) {
              const dateIndex = selectedDates.indexOf(date);
              if (dateIndex > -1) {
                selectedDates.splice(dateIndex, 1);
                day.classList.remove('selected');
              } else {
                selectedDates.push(date);
                day.classList.add('selected');
              }
              dateRangeInfo.textContent = `Custom selection: ${selectedDates.length} dates selected`;
              continueBtn.disabled = selectedDates.length === 0;
              validateDateSelectionToast();
            }
          });
        });
      }
    });
  }

  // Helper functions
  function getAllDatesInMonth(year, month) {
    const dates = [];
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    for (let day = 1; day <= daysInMonth; day++) {
      dates.push(`${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`);
    }
    
    return dates;
  }

  function getWorkDaysInMonth(year, month) {
    const dates = [];
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    for (let day = 1; day <= daysInMonth; day++) {
      const date = new Date(year, month, day);
      const dayOfWeek = date.getDay();
      
      // Monday = 1, Tuesday = 2, ..., Friday = 5
      if (dayOfWeek >= 1 && dayOfWeek <= 5) {
        dates.push(`${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`);
      }
    }
    
    return dates;
  }

  // Auto-disappearing toast logic (changed from persistent)
  function showPersistentToast(type, message) {
    // Remove any existing warning/info toast
    document.querySelectorAll('.modern-toast.persistent-toast').forEach(t => t.remove());
    const toast = document.createElement('div');
    toast.className = `modern-toast modern-toast-${type} persistent-toast show`;
    toast.innerHTML = `
      <div class="toast-content">
        <span class="toast-message">${message}</span>
        <button class="toast-close">&times;</button>
      </div>
    `;
    document.body.appendChild(toast);
    const closeBtn = toast.querySelector('.toast-close');
    closeBtn.addEventListener('click', () => {
      toast.classList.remove('show');
      setTimeout(() => toast.remove(), 300);
    });
    
    // Auto-disappear after 4 seconds (instead of staying persistent)
    setTimeout(() => {
      if (toast.parentElement) {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
      }
    }, 4000);
  }
  function removePersistentToast() {
    document.querySelectorAll('.modern-toast.persistent-toast').forEach(t => t.remove());
  }

  // Date validation and toast logic (show once per change)
  let lastToastMessage = null; // Track last message to avoid repeats
  
  function validateDateSelectionToast() {
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    
    let currentMessage = null;
    
    // 1. Year check
    if (selectedYear < currentYear) {
      currentMessage = 'You are picking a date range in the past.';
    } else if (selectedYear > currentYear) {
      // Future year, no toast
      currentMessage = null;
    } else {
      // 2. Month check (current year)
      if (selectedMonth < currentMonth) {
        currentMessage = 'You are picking a date range in the past.';
      } else if (selectedMonth > currentMonth) {
        // Future month, no toast
        currentMessage = null;
      } else {
        // 3. Day check (current month)
        if (selectedDateRange === 'full' || selectedDateRange === 'workdays' || selectedDateRange === 'custom') {
          if (!selectedDates || selectedDates.length === 0) {
            currentMessage = null;
          } else {
            let allPast = true;
            let allFuture = true;
            for (const dateStr of selectedDates) {
              const dateObj = new Date(dateStr);
              if (dateObj < today) {
                allFuture = false;
              } else if (dateObj > today) {
                allPast = false;
              } else {
                allPast = false;
                allFuture = false;
              }
            }
            if (allPast) {
              currentMessage = 'You are picking a date range in the past.';
            } else if (allFuture) {
              // No toast
              currentMessage = null;
            } else {
              currentMessage = 'You have selected a range that includes both past and future dates.';
            }
          }
        }
      }
    }
    
    // Only show toast if message changed (prevents spam)
    if (currentMessage !== lastToastMessage) {
      removePersistentToast();
      if (currentMessage) {
        const toastType = currentMessage.includes('past') ? 'warning' : 'info';
        showPersistentToast(toastType, currentMessage);
      }
      lastToastMessage = currentMessage;
    }
  }

  // Initialize the calendar preview
  updateCalendarPreview();
  updateLegendColorPreview();
}
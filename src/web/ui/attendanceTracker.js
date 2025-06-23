// Employee Shift Tracker Module
// Handles employee management, validation, and shift tracking functionality

import { sanitizeEmployeeName, sanitizeEmployeeText, sanitizeEmail, sanitizePhoneNumber } from '../../shared/utils/sanitize.js';
import { validateEmployeeName, validateEmployeeEmail, validateEmployeeTitle, validateEmployeePhone, validateShiftTime, validateBreakTime } from '../../shared/utils/validation.js';
import { renderEmployeePreview, renderEmployeeStats } from '../utils/previewEmployee.js';

// Debug: Track successful imports
console.log('✓ AttendanceTracker imports loaded:', {
  sanitizeEmployeeName: typeof sanitizeEmployeeName,
  sanitizeEmployeeText: typeof sanitizeEmployeeText,
  sanitizeEmail: typeof sanitizeEmail,
  sanitizePhoneNumber: typeof sanitizePhoneNumber,
  validateEmployeeName: typeof validateEmployeeName,
  validateEmployeeEmail: typeof validateEmployeeEmail,
  validateEmployeeTitle: typeof validateEmployeeTitle,
  validateEmployeePhone: typeof validateEmployeePhone,
  validateShiftTime: typeof validateShiftTime,
  validateBreakTime: typeof validateBreakTime,
  renderEmployeePreview: typeof renderEmployeePreview,
  renderEmployeeStats: typeof renderEmployeeStats
});

export class AttendanceTracker {
  constructor() {
    console.log('🏢 AttendanceTracker: Initializing...');
    this.employees = [];
    this.maxEmployees = 60;
    this.editingIndex = -1;
    this.hasUnsavedChanges = false;
    this.loadEmployees();
    this.setupEventListeners();
    this.setupNavigationWarning();
    console.log('✅ AttendanceTracker: Initialization complete');
  }/**
   * Sanitize employee data for security using shared utilities
   */  
  sanitizeEmployeeData(data) {
    // Use validation functions that also sanitize
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
      shifts: {
        start: data.shifts.start,
        firstBreak: data.shifts.firstBreak,
        lunch: data.shifts.lunch,
        secondBreak: data.shifts.secondBreak,
        end: data.shifts.end
      },
      dateAdded: new Date().toISOString()
    };
  }
  /**
   * Initialize all event listeners for the attendance tracker
   */
  setupEventListeners() {
    // Form submission
    const employeeForm = document.getElementById('employeeForm');
    if (employeeForm) {
      employeeForm.addEventListener('submit', (e) => this.handleAddEmployee(e));
    }

    // Clear form button
    const clearFormBtn = document.getElementById('clearFormBtn');
    if (clearFormBtn) {
      clearFormBtn.addEventListener('click', () => this.clearForm());
    }

    // Shift preset buttons
    this.setupShiftPresets();

    // CSV upload
    const csvUploadArea = document.getElementById('csvUploadArea');
    const csvFileInput = document.getElementById('csvFileInput');
    
    if (csvUploadArea && csvFileInput) {
      csvUploadArea.addEventListener('click', () => csvFileInput.click());
      csvUploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
      csvUploadArea.addEventListener('drop', (e) => this.handleFileDrop(e));
      csvFileInput.addEventListener('change', (e) => this.handleFileSelect(e));
    }

    // List actions
    const exportListBtn = document.getElementById('exportListBtn');
    const clearAllBtn = document.getElementById('clearAllBtn');
    
    if (exportListBtn) {
      exportListBtn.addEventListener('click', () => this.exportEmployeeList());
    }
    
    if (clearAllBtn) {
      clearAllBtn.addEventListener('click', () => this.clearAllEmployees());
    }

    // Generate tracker
    const generateTrackerBtn = document.getElementById('generateTrackerBtn');
    if (generateTrackerBtn) {
      generateTrackerBtn.addEventListener('click', () => this.generateShiftTracker());
    }

    // Real-time validation
    this.setupRealTimeValidation();
  }

  /**
   * Setup shift preset buttons
   */
  setupShiftPresets() {
    const presetButtons = document.querySelectorAll('.preset-btn');
    presetButtons.forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        const preset = btn.dataset.preset;
        this.applyShiftPreset(preset);
        
        // Update active state
        presetButtons.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
      });
    });
  }
  /**
   * Apply shift preset times
   */
  applyShiftPreset(preset) {
    const presets = {
      'first-shift': {
        start: '08:00',
        firstBreak: '10:00',
        lunch: '12:00', 
        secondBreak: '14:30',
        end: '16:30'
      },
      'second-shift': {
        start: '09:00',
        firstBreak: '11:00',
        lunch: '13:00',
        secondBreak: '15:30',
        end: '17:30'
      },
      'third-shift': {
        start: '11:30',
        firstBreak: '13:30',
        lunch: '15:30',
        secondBreak: '17:30',
        end: '20:00'
      }
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

  /**
   * Setup navigation warning for unsaved changes
   */
  setupNavigationWarning() {
    // Show warning on page reload
    window.addEventListener('beforeunload', (e) => {
      if (this.employees.length > 0) {
        const message = 'You have employee data that will be lost. Are you sure you want to leave?';
        e.preventDefault();
        e.returnValue = message;
        return message;
      }
    });    // Show warning when navigating away
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
      if (item.id !== 'nav-attendance') {
        item.addEventListener('click', (e) => {
          if (this.employees.length > 0) {
            e.preventDefault();
            this.showConfirmToast(
              'You have employee data that will be lost. Continue?',
              () => {
                // Clear data and navigate
                this.employees = [];
                this.clearSavedData();
                // Trigger the original navigation
                const originalEvent = new MouseEvent('click', {
                  view: window,
                  bubbles: true,
                  cancelable: true
                });
                // Remove this listener temporarily to avoid recursion
                item.removeEventListener('click', arguments.callee);
                item.dispatchEvent(originalEvent);
              }
            );
          }
        });
      }
    });    // Show data persistence warning on first load
    if (this.employees.length === 0) {
      setTimeout(() => {
        this.showToast('info', 'Note: Employee data is stored for this browser session only and will be cleared when you close this tab.');
      }, 1000);
    }
  }  /**
   * Setup real-time validation for form fields
   */
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

    // Handle color picker functionality
    this.setupColorPicker();
  }

  /**
   * Setup color picker functionality
   */
  setupColorPicker() {
    const colorPicker = document.getElementById('employeeColor');
    const colorIndicator = document.querySelector('.color-indicator');
    
    if (colorPicker && colorIndicator) {
      // Update indicator when color changes
      colorPicker.addEventListener('input', (e) => {
        colorIndicator.style.backgroundColor = e.target.value;
      });
      
      colorPicker.addEventListener('change', (e) => {
        colorIndicator.style.backgroundColor = e.target.value;
      });
      
      // Make clicking the indicator open the color picker
      colorIndicator.addEventListener('click', () => {
        colorPicker.click();
      });
      
      // Set initial color
      colorIndicator.style.backgroundColor = colorPicker.value;
    }
  }
  /**
   * Handle adding a new employee
   */
  async handleAddEmployee(event) {
    event.preventDefault();
    
    if (this.employees.length >= this.maxEmployees) {
      this.showToast('error', `Maximum of ${this.maxEmployees} employees allowed.`);
      return;
    }

    const formData = this.getFormData();
    
    // Validate all fields
    if (!this.validateAllFields(formData)) {
      this.showToast('error', 'Please fix the validation errors before adding the employee.');
      return;
    }

    // Check for duplicates (but allow editing same employee)
    if (this.editingIndex === -1 && this.isDuplicateEmployee(formData)) {
      this.showToast('error', 'An employee with this name or ID already exists.');
      return;
    }

    // Sanitize input data
    const sanitizedEmployee = this.sanitizeEmployeeData(formData);
    
    // Add or update employee
    if (this.editingIndex === -1) {
      this.employees.push(sanitizedEmployee);
      this.showToast('success', `Employee ${sanitizedEmployee.name} added successfully!`);
    } else {
      this.employees[this.editingIndex] = sanitizedEmployee;
      this.showToast('success', `Employee ${sanitizedEmployee.name} updated successfully!`);
      this.editingIndex = -1;
      
      // Update button text back
      const submitBtn = document.querySelector('.add-employee-btn');
      if (submitBtn) {
        submitBtn.textContent = 'Add Employee';
      }
    }
    
    this.saveEmployees();
    this.renderEmployeeList();
    this.clearForm();
    this.updateGenerateButton();
  }

  /**
   * Get form data from all input fields
   */
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
  /**
   * Validate all form fields using enhanced validation utilities
   */
  validateAllFields(data) {
    let isValid = true;
    const errors = [];

    // Validate employee name
    const nameValidation = validateEmployeeName(data.name);
    if (!nameValidation.isValid) {
      this.showFieldError('employeeName', nameValidation.error);
      isValid = false;
    }

    // Validate job title
    const titleValidation = validateEmployeeTitle(data.title);
    if (!titleValidation.isValid) {
      this.showFieldError('employeeTitle', titleValidation.error);
      isValid = false;
    }

    // Validate email (optional but must be valid if provided)
    const emailValidation = validateEmployeeEmail(data.email);
    if (!emailValidation.isValid) {
      this.showFieldError('employeeEmail', emailValidation.error);
      isValid = false;
    }

    // Validate phone (optional but must be valid if provided)
    const phoneValidation = validateEmployeePhone(data.phone);
    if (!phoneValidation.isValid) {
      this.showFieldError('employeePhone', phoneValidation.error);
      isValid = false;
    }

    // Validate shift times
    const shiftValidation = validateShiftTime(data.shifts.start, data.shifts.end);
    if (!shiftValidation.isValid) {
      this.showFieldError('shiftEnd', shiftValidation.error);
      isValid = false;
    }

    // Validate break times if provided
    if (data.shifts.firstBreak) {
      const firstBreakValidation = validateBreakTime(data.shifts.firstBreak, data.shifts.start, data.shifts.end, 'First break');
      if (!firstBreakValidation.isValid) {
        this.showFieldError('firstBreak', firstBreakValidation.error);
        isValid = false;
      }
    }

    if (data.shifts.lunch) {
      const lunchValidation = validateBreakTime(data.shifts.lunch, data.shifts.start, data.shifts.end, 'Lunch break');
      if (!lunchValidation.isValid) {
        this.showFieldError('lunchBreak', lunchValidation.error);
        isValid = false;
      }
    }

    if (data.shifts.secondBreak) {
      const secondBreakValidation = validateBreakTime(data.shifts.secondBreak, data.shifts.start, data.shifts.end, 'Second break');
      if (!secondBreakValidation.isValid) {
        this.showFieldError('secondBreak', secondBreakValidation.error);
        isValid = false;
      }
    }

    return isValid;
  }

  /**
   * Validate shift times logic
   */
  validateShiftTimes(shifts) {
    const startTime = this.timeToMinutes(shifts.start);
    const endTime = this.timeToMinutes(shifts.end);
    
    if (endTime <= startTime) {
      this.showFieldError('shiftEnd', 'End time must be after start time');
      return false;
    }

    // Validate break times are within shift
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

  /**
   * Convert time string to minutes for comparison
   */
  timeToMinutes(timeString) {
    const [hours, minutes] = timeString.split(':').map(Number);
    return hours * 60 + minutes;
  }
  /**
   * Check for duplicate employees
   */
  isDuplicateEmployee(newEmployee) {
    return this.employees.some(emp => 
      emp.name.toLowerCase() === newEmployee.name.toLowerCase() ||
      (newEmployee.id && emp.id && emp.id.toLowerCase() === newEmployee.id.toLowerCase()) ||
      (newEmployee.email && emp.email && emp.email.toLowerCase() === newEmployee.email.toLowerCase())
    );
  }

  /**
   * Show field-specific error message
   */
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

  /**
   * Clear field error state
   */
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
  /**
   * Validate individual field using enhanced validation utilities
   */
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
  }/**
   * Update employee preview using shared utility
   */
  updateEmployeePreview() {
    const previewContainer = document.querySelector('.employee-preview');
    if (!previewContainer) return;

    const previewHTML = renderEmployeePreview(this.employees);
    const statsHTML = renderEmployeeStats(this.employees);
    
    previewContainer.innerHTML = previewHTML + statsHTML;
  }
  /**
   * Render the employee list
   */
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

    // Update preview after list changes
    this.updateEmployeePreview();
  }

  /**
   * Format time for display
   */
  formatShiftTime(time) {
    if (!time) return '';
    
    const [hours, minutes] = time.split(':');
    const hour = parseInt(hours);
    const ampm = hour >= 12 ? 'PM' : 'AM';
    const displayHour = hour % 12 || 12;
    
    return `${displayHour}:${minutes} ${ampm}`;
  }  /**
   * Delete an employee
   */
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

  /**
   * Edit an employee
   */
  editEmployee(index) {
    const employee = this.employees[index];
    this.editingIndex = index;
    
    // Fill form with employee data
    document.getElementById('employeeName').value = employee.name;
    document.getElementById('employeeId').value = employee.id || '';
    document.getElementById('employeeEmail').value = employee.email || '';
    document.getElementById('employeeTitle').value = employee.title;
    document.getElementById('employeePhone').value = employee.phone || '';
    document.getElementById('employeeColor').value = employee.color;
    
    // Update color indicator
    const colorIndicator = document.querySelector('.color-indicator');
    if (colorIndicator) {
      colorIndicator.style.backgroundColor = employee.color;
    }
    
    // Fill shift times
    document.getElementById('shiftStart').value = employee.shifts.start;
    document.getElementById('firstBreak').value = employee.shifts.firstBreak || '';
    document.getElementById('lunchBreak').value = employee.shifts.lunch || '';
    document.getElementById('secondBreak').value = employee.shifts.secondBreak || '';
    document.getElementById('shiftEnd').value = employee.shifts.end;
    
    // Update button text
    const submitBtn = document.querySelector('.add-employee-btn');
    if (submitBtn) {
      submitBtn.textContent = 'Update Employee';
    }
    
    // Reset preset buttons to custom
    document.querySelectorAll('.preset-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelector('.preset-btn.custom').classList.add('active');
    
    // Scroll to form
    document.querySelector('.employee-form-section').scrollIntoView({ behavior: 'smooth' });
    
    this.showToast('info', `Editing ${employee.name} - make changes and click Update Employee`);
  }
  /**
   * Clear the form
   */
  clearForm() {
    const form = document.getElementById('employeeForm');
    if (form) {
      form.reset();
      // Reset color picker to default
      const colorPicker = document.getElementById('employeeColor');
      const colorIndicator = document.querySelector('.color-indicator');
      if (colorPicker && colorIndicator) {
        colorPicker.value = '#4CAF50';
        colorIndicator.style.backgroundColor = '#4CAF50';
      }
      // Clear all error states
      form.querySelectorAll('.form-group').forEach(group => {
        group.classList.remove('error');
        const errorMsg = group.querySelector('.error-message');
        if (errorMsg) {
          errorMsg.classList.remove('show');
        }
      });
      
      // Reset editing state
      this.editingIndex = -1;
      const submitBtn = document.querySelector('.add-employee-btn');
      if (submitBtn) {
        submitBtn.textContent = 'Add Employee';
      }
      
      // Reset to custom preset
      document.querySelectorAll('.preset-btn').forEach(btn => btn.classList.remove('active'));
      document.querySelector('.preset-btn.custom').classList.add('active');
    }
  }  /**
   * Clear all employees
   */
  clearAllEmployees() {
    // Only show toast if called by user, not by navigation
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
  /**
   * Clear all saved data from sessionStorage
   */
  clearSavedData() {
    try {
      sessionStorage.removeItem('employeeShiftData');

    } catch (error) {
      console.error('Failed to clear saved data:', error);
    }
  }

  /**
   * Update generate button state
   */
  updateGenerateButton() {
    const generateBtn = document.getElementById('generateTrackerBtn');
    if (generateBtn) {
      generateBtn.disabled = this.employees.length === 0;
    }
  }
  /**
   * Show alert message
   */
  showAlert(type, message) {
    this.showToast(type, message);
  }

  /**
   * Show modern toast notification at top of page
   */
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
    
    // Trigger animation
    setTimeout(() => toast.classList.add('show'), 10);
    
    // Auto remove after 5 seconds
    setTimeout(() => {
      if (toast.parentElement) {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
      }
    }, 5000);
  }

  /**
   * Get icon for toast type
   */
  getToastIcon(type) {
    const icons = {
      'success': '✅',
      'error': '❌',
      'warning': '⚠️',
      'info': 'ℹ️'
    };
    return icons[type] || 'ℹ️';
  }
  /**
   * Save employees to sessionStorage (clears when tab closes)
   */
  saveEmployees() {
    try {
      // Use sessionStorage instead of localStorage for automatic cleanup
      sessionStorage.setItem('employeeShiftData', JSON.stringify(this.employees));
    } catch (error) {
      console.error('Failed to save employee data:', error);
      this.showToast('error', 'Failed to save employee data');
    }
  }

  /**
   * Load employees from sessionStorage
   */
  loadEmployees() {
    try {
      const saved = sessionStorage.getItem('employeeShiftData');
      this.employees = saved ? JSON.parse(saved) : [];
    } catch (error) {
      console.error('Failed to load employee data:', error);
      this.employees = [];
    }
  }

  /**
   * Export employee list as CSV
   */
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

  /**
   * Handle file drag over
   */
  handleDragOver(event) {
    event.preventDefault();
    event.currentTarget.classList.add('dragover');
  }

  /**
   * Handle file drop
   */
  handleFileDrop(event) {
    event.preventDefault();
    event.currentTarget.classList.remove('dragover');
    
    const files = event.dataTransfer.files;
    if (files.length > 0) {
      this.processUploadedFile(files[0]);
    }
  }

  /**
   * Handle file selection
   */
  handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
      this.processUploadedFile(file);
    }
  }

  /**
   * Process uploaded CSV/Excel file  /**
   * Process uploaded CSV/Excel file
   */
  async processUploadedFile(file) {
    // This would implement CSV/Excel parsing
    // For now, show a placeholder message
    this.showAlert('info', 'Bulk upload feature will be implemented in the next update');
  }
  /**
   * Generate shift tracker Excel file
   */
  async generateShiftTracker() {
    if (this.employees.length === 0) {
      this.showToast('warning', 'No employees to generate tracker for');
      return;
    }    try {
      // Import the modular generators
      const attendanceModule = await import('../../excel/generators/attendance/attendanceTrackerSheet.js');
      const referenceModule = await import('../../excel/generators/attendance/referenceSheet.js');
      const instructionsModule = await import('../../excel/generators/attendance/instructionsSheet.js');
      const stylesModule = await import('../../excel/generators/attendance/stylesXml.js');
      const workbookModule = await import('../../excel/generators/attendance/workbookXml.js');
      const contentTypesModule = await import('../../excel/generators/attendance/contentTypesXml.js');
      const zipModule = await import('../../excel/utils/zipUtils.js');
      
      // Generate all sheet XML content
      const shiftTrackerSheet = attendanceModule.buildShiftTrackerSheet(this.employees);
      const referenceSheet = referenceModule.buildReferenceSheet(this.employees);
      const instructionsSheet = instructionsModule.buildInstructionsSheet();
      const stylesXML = stylesModule.getShiftTrackerStylesXML();
      const workbookXML = workbookModule.getShiftTrackerWorkbookXML();
      const workbookRelsXML = workbookModule.getShiftTrackerWorkbookRelsXML();
      const contentTypesXML = contentTypesModule.getShiftTrackerContentTypesXML();
      const relsXML = contentTypesModule.getShiftTrackerRelsXML();

      // Create the file structure
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

      // Generate ZIP file
      const zipBuffer = zipModule.createZip(files);
      
      // Download the file
      this.downloadExcelFile(zipBuffer);
      
      this.showToast('success', 'Shift tracker Excel file generated successfully!');
    } catch (error) {
      console.error('Error generating Excel file:', error);
      this.showToast('error', 'Failed to generate Excel file. Please try again.');
    }
  }

  /**
   * Download the Excel file
   */
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
  /**
   * Initialize the tracker when the page loads
   */
  init() {
    this.renderEmployeeList();
    this.updateGenerateButton();
    
    // Add console helper for developers to clear data
    window.clearEmployeeData = () => {
      this.employees = [];
      this.clearSavedData();
      this.renderEmployeeList();
      this.updateGenerateButton();

      location.reload();
    };
    
    // Show instructions in console
    if (this.employees.length > 0) {
      console.log('💡 Developer tip: Run clearEmployeeData() in console to clear all saved employee data');
    }
  }
}

// Create global instance
let attendanceTracker;

export function setupAttendanceTracker() {
  attendanceTracker = new AttendanceTracker();
  attendanceTracker.init();
  
  // Make it globally accessible for button onclick handlers
  window.attendanceTracker = attendanceTracker;
}
// utils/validation.js
// Input validation using sanitization utilities
import { sanitizeLegendInput, sanitizeEmployeeName, sanitizeEmployeeText, sanitizeEmail, sanitizePhoneNumber } from './sanitize.js';

// Debug: Track successful imports
console.log('✓ Validation imports loaded:', {
  sanitizeLegendInput: typeof sanitizeLegendInput,
  sanitizeEmployeeName: typeof sanitizeEmployeeName,
  sanitizeEmployeeText: typeof sanitizeEmployeeText,
  sanitizeEmail: typeof sanitizeEmail,
  sanitizePhoneNumber: typeof sanitizePhoneNumber
});

export function validateLegendInput(input) {
  const sanitized = sanitizeLegendInput(input);
  const suspiciousPatterns = [
    /eval\s*\(/i,
    /function\s*\(/i,
    /alert\s*\(/i,
    /document\./i,
    /window\./i,
    /\$\{.*\}/i,
    /\[\s*\]/i,
    /\.\s*constructor/i
  ];
  for (const pattern of suspiciousPatterns) {
    if (pattern.test(sanitized)) {
      return 'Enter Value Here';
    }
  }
  return sanitized;
}

export function validateEmployeeName(input) {
  if (!input || typeof input !== 'string') {
    return { isValid: false, error: 'Employee name is required', sanitized: '' };
  }
  
  const sanitized = sanitizeEmployeeName(input);
  
  if (!sanitized || sanitized.length < 2) {
    return { isValid: false, error: 'Name must be at least 2 characters', sanitized };
  }
  
  if (sanitized.length > 100) {
    return { isValid: false, error: 'Name is too long (max 100 characters)', sanitized };
  }
  
  // Check for suspicious patterns
  const suspiciousPatterns = [
    /eval\s*\(/i, /function\s*\(/i, /alert\s*\(/i, /document\./i, /window\./i
  ];
  
  for (const pattern of suspiciousPatterns) {
    if (pattern.test(sanitized)) {
      return { isValid: false, error: 'Invalid characters in name', sanitized: '' };
    }
  }
  
  return { isValid: true, error: null, sanitized };
}

export function validateEmployeeEmail(input) {
  if (!input || typeof input !== 'string') {
    return { isValid: true, error: null, sanitized: '' }; // Email is optional
  }
  
  const sanitized = sanitizeEmail(input);
  
  if (!sanitized) {
    return { isValid: true, error: null, sanitized: '' };
  }
  
  // Enhanced email regex
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  
  if (!emailRegex.test(sanitized)) {
    return { isValid: false, error: 'Please enter a valid email address', sanitized };
  }
  
  return { isValid: true, error: null, sanitized };
}

export function validateEmployeeTitle(input) {
  if (!input || typeof input !== 'string') {
    return { isValid: false, error: 'Job title is required', sanitized: '' };
  }
  
  const sanitized = sanitizeEmployeeText(input, 100);
  
  if (!sanitized || sanitized.length < 2) {
    return { isValid: false, error: 'Job title must be at least 2 characters', sanitized };
  }
  
  return { isValid: true, error: null, sanitized };
}

export function validateEmployeePhone(input) {
  if (!input || typeof input !== 'string') {
    return { isValid: true, error: null, sanitized: '' }; // Phone is optional
  }
  
  const sanitized = sanitizePhoneNumber(input);
  
  if (!sanitized) {
    return { isValid: true, error: null, sanitized: '' };
  }
  
  // Basic phone validation (allows various formats)
  const phoneRegex = /^[\+]?[0-9\s\-\(\)\.]{7,20}$/;
  
  if (!phoneRegex.test(sanitized)) {
    return { isValid: false, error: 'Please enter a valid phone number', sanitized };
  }
  
  return { isValid: true, error: null, sanitized };
}

export function validateShiftTime(startTime, endTime) {
  if (!startTime || !endTime) {
    return { isValid: false, error: 'Both start and end times are required' };
  }
  
  const startMinutes = timeToMinutes(startTime);
  const endMinutes = timeToMinutes(endTime);
  
  if (endMinutes <= startMinutes) {
    return { isValid: false, error: 'End time must be after start time' };
  }
  
  // Check for reasonable shift length (max 16 hours)
  const shiftLength = endMinutes - startMinutes;
  if (shiftLength > 16 * 60) {
    return { isValid: false, error: 'Shift cannot be longer than 16 hours' };
  }
  
  return { isValid: true, error: null };
}

export function validateBreakTime(breakTime, startTime, endTime, breakName = 'break') {
  if (!breakTime) {
    return { isValid: true, error: null }; // Breaks are optional
  }
  
  const breakMinutes = timeToMinutes(breakTime);
  const startMinutes = timeToMinutes(startTime);
  const endMinutes = timeToMinutes(endTime);
  
  if (breakMinutes <= startMinutes || breakMinutes >= endMinutes) {
    return { 
      isValid: false, 
      error: `${breakName} must be between shift start and end times` 
    };
  }
  
  return { isValid: true, error: null };
}

// Helper function to convert time to minutes
function timeToMinutes(timeString) {
  if (!timeString) return 0;
  const [hours, minutes] = timeString.split(':').map(Number);
  return hours * 60 + minutes;
}

// validateLegendInput in validation.js matches the logic in script.js. No changes needed.

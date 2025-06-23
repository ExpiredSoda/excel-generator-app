// utils/sanitize.js
// Input sanitization for legend values and employee data

// Debug: Track module loading
console.log('✓ Sanitize: Module loaded');

export function sanitizeLegendInput(input) {
  if (typeof input !== 'string') return '';
  let sanitized = input
    .replace(/<[^>]*>/g, '')
    .replace(/[<>"'&]/g, '')
    .replace(/javascript:/gi, '')
    .replace(/on\w+\s*= /gi, '')
    .replace(/\s+/g, ' ')
    .trim();
  sanitized = sanitized.substring(0, 50);
  return sanitized || 'Enter Value Here';
}

export function sanitizeEmployeeName(input) {
  if (typeof input !== 'string') return '';
  let sanitized = input
    .replace(/<[^>]*>/g, '')          // Remove HTML tags
    .replace(/[<>"'&]/g, '')          // Remove dangerous characters
    .replace(/javascript:/gi, '')     // Remove JavaScript protocols
    .replace(/on\w+\s*= /gi, '')     // Remove event handlers
    .replace(/[^\w\s\-'.,]/g, '')    // Only allow letters, numbers, spaces, hyphens, apostrophes, periods, commas
    .replace(/\s+/g, ' ')            // Normalize whitespace
    .trim();
  sanitized = sanitized.substring(0, 100); // Reasonable name length
  return sanitized;
}

export function sanitizeEmployeeText(input, maxLength = 200) {
  if (typeof input !== 'string') return '';
  let sanitized = input
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '') // Remove script tags
    .replace(/<[^>]*>/g, '')          // Remove HTML tags
    .replace(/[<>"'&]/g, '')          // Remove dangerous characters
    .replace(/javascript:/gi, '')     // Remove JavaScript protocols
    .replace(/on\w+\s*= /gi, '')     // Remove event handlers
    .replace(/\s+/g, ' ')            // Normalize whitespace
    .trim();
  sanitized = sanitized.substring(0, maxLength);
  return sanitized;
}

export function sanitizeEmail(input) {
  if (typeof input !== 'string') return '';
  let sanitized = input
    .replace(/<[^>]*>/g, '')          // Remove HTML tags
    .replace(/[<>"'&]/g, '')          // Remove dangerous characters
    .replace(/javascript:/gi, '')     // Remove JavaScript protocols
    .replace(/\s+/g, '')             // Remove all whitespace from email
    .toLowerCase()                    // Normalize to lowercase
    .trim();
  sanitized = sanitized.substring(0, 320); // RFC 5321 email length limit
  return sanitized;
}

export function sanitizePhoneNumber(input) {
  if (typeof input !== 'string') return '';
  let sanitized = input
    .replace(/<[^>]*>/g, '')          // Remove HTML tags
    .replace(/[^0-9\-\(\)\+\s\.]/g, '') // Only allow phone number characters
    .replace(/\s+/g, ' ')            // Normalize whitespace
    .trim();
  sanitized = sanitized.substring(0, 20); // Reasonable phone length
  return sanitized;
}

// sanitizeLegendInput in sanitize.js matches the logic in script.js. No changes needed.

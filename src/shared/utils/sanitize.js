// utils/sanitize.js
// Input sanitization for legend values and employee data

// --- Common Helper ---
function baseSanitize(input) {
  if (typeof input !== 'string') return '';
  return input
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '') // Remove script tags
    .replace(/<[^>]*>/g, '')          // Remove HTML tags
    .replace(/[<>"'&]/g, '')         // Remove dangerous characters
    .replace(/javascript:/gi, '')     // Remove JavaScript protocols
    .replace(/on\w+\s*= /gi, '')    // Remove event handlers
    .replace(/\s+/g, ' ')            // Normalize whitespace
    .trim();
}

export function sanitizeLegendInput(input) {
  let sanitized = baseSanitize(input);
  sanitized = sanitized.substring(0, 50);
  return sanitized || 'Enter Value Here';
}

export function sanitizeEmployeeName(input) {
  let sanitized = baseSanitize(input)
    .replace(/[^\w\s\-\'\.,]/g, ''); // Only allow letters, numbers, spaces, hyphens, apostrophes, periods, commas
  sanitized = sanitized.substring(0, 100);
  return sanitized;
}

export function sanitizeEmployeeText(input, maxLength = 200) {
  let sanitized = baseSanitize(input);
  sanitized = sanitized.substring(0, maxLength);
  return sanitized;
}

export function sanitizeEmail(input) {
  let sanitized = baseSanitize(input)
    .replace(/\s+/g, '')             // Remove all whitespace from email
    .toLowerCase();                   // Normalize to lowercase
  sanitized = sanitized.substring(0, 320); // RFC 5321 email length limit
  return sanitized;
}

export function sanitizePhoneNumber(input) {
  let sanitized = baseSanitize(input)
    .replace(/[^0-9\-\(\)\+\s\.]/g, '') // Only allow phone number characters
    .substring(0, 20);
  return sanitized;
}

// sanitizeLegendInput in sanitize.js matches the logic in script.js. No changes needed.

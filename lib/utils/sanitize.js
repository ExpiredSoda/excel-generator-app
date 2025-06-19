// utils/sanitize.js
// Input sanitization for legend values

/**
 * Sanitize legend input to remove dangerous content
 * @param {string} input - Raw input from user
 * @returns {string} - Sanitized input safe for Excel
 */
export function sanitizeLegendInput(input) {
  if (typeof input !== 'string') return 'Enter Value Here';
  
  // Remove any HTML tags, scripts, and dangerous characters
  let sanitized = input
    .replace(/<[^>]*>/g, '') // Remove HTML tags
    .replace(/[<>\"'&]/g, '') // Remove dangerous characters
    .replace(/javascript:/gi, '') // Remove javascript: protocol
    .replace(/on\w+\s*=/gi, '') // Remove event handlers
    .replace(/\s+/g, ' ') // Normalize whitespace
    .trim();
  
  // Limit length to prevent abuse
  sanitized = sanitized.substring(0, 50);
  
  // If empty after sanitization, return default
  return sanitized || 'Enter Value Here';
}

/**
 * Validate legend input for suspicious patterns
 * @param {string} input - Input to validate
 * @returns {string} - Validated and sanitized input
 */
export function validateLegendInput(input) {
  const sanitized = sanitizeLegendInput(input);
  
  // Check for suspicious patterns
  const suspiciousPatterns = [
    /eval\s*\(/i,
    /function\s*\(/i,
    /alert\s*\(/i,
    /document\./i,
    /window\./i,
    /\$\{.*\}/i, // Template literals
    /\[\s*\]/i, // Array access
    /\.\s*constructor/i
  ];
  
  for (const pattern of suspiciousPatterns) {
    if (pattern.test(sanitized)) {
      console.warn('Suspicious input detected:', input);
      return 'Enter Value Here'; // Return safe default
    }
  }
  
  return sanitized;
}

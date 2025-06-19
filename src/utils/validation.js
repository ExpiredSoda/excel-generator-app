// utils/validation.js
// Legend input validation using sanitizeLegendInput
import { sanitizeLegendInput } from './sanitize.js';

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
// validateLegendInput in validation.js matches the logic in script.js. No changes needed.

// utils/sanitize.js
// Input sanitization for legend values

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

// sanitizeLegendInput in sanitize.js matches the logic in script.js. No changes needed.

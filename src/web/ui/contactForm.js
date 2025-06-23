import { sanitizeEmployeeName, sanitizeEmployeeText } from '../../shared/utils/sanitize.js';
import { validateEmployeeEmail } from '../../shared/utils/validation.js';

export function setupContactForm() {
  const form = document.getElementById('contactForm');
  const status = document.getElementById('contactFormStatus');
  if (form) {
    form.onsubmit = function(e) {
      e.preventDefault();
      const name = sanitizeEmployeeName(document.getElementById('contactName').value.trim());
      const email = document.getElementById('contactEmail').value.trim();
      const message = sanitizeEmployeeText(document.getElementById('contactMessage').value.trim(), 1000);
      if (!name || !email || !message) {
        status.textContent = 'Please fill out all fields.';
        status.style.color = '#dc3545';
        return;
      }
      const emailResult = validateEmployeeEmail(email);
      if (!emailResult.isValid) {
        status.textContent = emailResult.error || 'Invalid email address.';
        status.style.color = '#dc3545';
        return;
      }
      const mailto = `mailto:danielplanos@freeexcelgenerator.com?subject=Contact%20from%20${encodeURIComponent(name)}&body=${encodeURIComponent(message + '\n\nFrom: ' + name + ' (' + email + ')')}`;
      window.location.href = mailto;
      status.textContent = 'Thank you! Your message will be sent via your email client. I usually respond in about 1-2 business days.';
      status.style.color = '#20b388';
    };
  }
} 
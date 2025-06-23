import { validateEmployeeEmail } from '../../shared/utils/validation.js';
import { sanitizeEmployeeName, sanitizeEmployeeText } from '../../shared/utils/sanitize.js';

// Sidebar navigation and dynamic content switching
export function setupNavigation() {
  const navItems = document.querySelectorAll('.nav-item');
  const mainContent = document.querySelector('.main-content');

  navItems.forEach(item => {
    item.addEventListener('click', (e) => {
      // Intercept navigation for unsaved attendance data
      const currentActive = document.querySelector('.nav-item.active');
      const isLeavingAttendance = currentActive && currentActive.id === 'nav-attendance' && item.id !== 'nav-attendance';
      if (isLeavingAttendance && window.attendanceTracker) {
        // Clear employee data and session storage when leaving Attendance Tracker (silent, no toast)
        window.attendanceTracker.clearAllEmployees('silent');
      }
      navItems.forEach(i => i.classList.remove('active'));
      item.classList.add('active');
      let page = item.id.replace('nav-', '');
      if (page === 'attendance') page = 'attendance';
      if (page === 'meeting') page = 'meeting';
      if (page === 'home') page = 'home';
      if (page === 'calendar') page = 'calendar';
      showPage(page, mainContent);
    });
  });
}

export function showPage(page, mainContent) {
  if (!mainContent) mainContent = document.querySelector('.main-content');
  if (page === 'home') {
    import('./homepage.js').then(module => {
      module.setupHomepage(mainContent);
    });
  } else if (page === 'calendar') {
    // Use the template for the calendar builder page
    const template = document.getElementById('calendarPageTemplate');
    if (template) {
      mainContent.innerHTML = template.innerHTML;
    } else {
      mainContent.innerHTML = '<h2>Calendar Builder template not found.</h2>';
      return;
    }
    setTimeout(async () => {
      // Dynamically import dependencies as in main.js
      const { setupCalendarBuilder } = await import('./calendarBuilder.js');
      const { buildCalendarSheetWithExcelBuilder } = await import('../../excel/generators/calendar/calendarBuilderSheet.js');
      const { getCalendarStylesXML } = await import('../../excel/generators/calendar/stylesXml.js');
      const { getContentTypesXML, getRelsXML } = await import('../../excel/generators/calendar/contentTypesXml.js');
      const { getSheet1InstructionsXML } = await import('../../excel/generators/calendar/instructionsSheet.js');
      const { getWorkbookXML, getWorkbookRelsXML } = await import('../../excel/generators/calendar/workbookXml.js');
      const { createZip } = await import('../../excel/utils/zipUtils.js');
      setupCalendarBuilder({
        buildCalendarSheetWithExcelBuilder,
        getCalendarStylesXML,
        getContentTypesXML,
        getRelsXML,
        getWorkbookXML,
        getWorkbookRelsXML,
        getSheet1InstructionsXML,
        createZip
      });
    }, 0);
  } else if (page === 'attendance') {
    mainContent.innerHTML = document.getElementById('attendancePageTemplate').innerHTML;
    setTimeout(() => {
      import('./attendanceTracker.js').then(module => {
        module.setupAttendanceTracker();
      });
    }, 0);
  } else if (page === 'meeting') {
    mainContent.innerHTML = `<h2>Meeting Tracker</h2><p>Coming soon! Create meeting tracking sheets here.</p>`;
  } else if (page === 'contact') {
    mainContent.innerHTML = `
      <div class="contact-container">
        <h2>Contact Me</h2>
        <form id="contactForm" class="contact-form">
          <div class="contact-form-group">
            <label for="contactName" class="contact-label">Name</label>
            <input type="text" id="contactName" name="name" required maxlength="60" class="contact-input">
          </div>
          <div class="contact-form-group">
            <label for="contactEmail" class="contact-label">Email</label>
            <input type="email" id="contactEmail" name="email" required maxlength="100" class="contact-input">
          </div>
          <div class="contact-form-group">
            <label for="contactMessage" class="contact-label">Message</label>
            <textarea id="contactMessage" name="message" required maxlength="1000" rows="6" class="contact-textarea"></textarea>
          </div>
          <button type="submit" class="btn generate-btn contact-submit-btn">Send Message</button>
        </form>
        <div id="contactFormStatus" class="contact-form-status"></div>
      </div>
    `;
    setTimeout(() => {
      import('./contactForm.js').then(module => {
        module.setupContactForm();
      });
    }, 0);
  } else {
    mainContent.innerHTML = '<h2>Page Not Found</h2>';
  }
}

// Add JS to handle the link to contact page
    setTimeout(() => {
      const suggestLink = document.getElementById('suggestLink');
      if (suggestLink) {
        suggestLink.addEventListener('click', function(e) {
          e.preventDefault();
          const navContact = document.getElementById('nav-contact');
          if (navContact) navContact.click();
        });
      }
    }, 100);

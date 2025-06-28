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
    const template = document.getElementById('calendarIntroTemplate');
    if (template) {
      mainContent.innerHTML = template.innerHTML;
    } else {
      mainContent.innerHTML = '<h2>Calendar Builder template not found.</h2>';
      return;
    }
    setTimeout(async () => {
      const { setupCalendarBuilder } = await import('./calendarBuilder.js');
      setupCalendarBuilder();
    }, 0);
  } else if (page === 'attendance') {
    mainContent.innerHTML = document.getElementById('attendanceIntroTemplate').innerHTML;
    setTimeout(() => {
      import('./attendanceTracker.js').then(module => {
        module.setupAttendanceBuilderPage();
      });
    }, 0);
  } else if (page === 'meeting') {
          mainContent.innerHTML = `
        <section class="modern-container-tool" style="--container-max-width: 800px; text-align: center;">
          <h2 style="color: #20b388; margin-bottom: 24px;">📁 Case Manager</h2>
          <div style="background: #f8f9fa; padding: 32px; border-radius: 12px; margin-bottom: 24px;">
            <h3 style="color: #333; margin-bottom: 16px;">🚧 Coming Soon!</h3>
            <p style="font-size: 1.1rem; color: #666; line-height: 1.6; margin-bottom: 20px;">
              A revolutionary front-end case management system where <strong>Excel acts as your portable database</strong>. 
              No accounts, no cloud dependency—just pure, portable data control.
            </p>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin: 24px 0;">
              <div style="background: white; padding: 20px; border-radius: 8px; border-left: 4px solid #20b388;">
                <h4 style="color: #20b388; margin-bottom: 8px;">🆕 New Users</h4>
                <p style="color: #666; font-size: 0.95rem;">Start fresh with a guided case creation wizard</p>
              </div>
              <div style="background: white; padding: 20px; border-radius: 8px; border-left: 4px solid #17a2b8;">
                <h4 style="color: #17a2b8; margin-bottom: 8px;">🔄 Returning Users</h4>
                <p style="color: #666; font-size: 0.95rem;">Upload your Excel file to continue where you left off</p>
              </div>
            </div>
            <p style="font-style: italic; color: #888; font-size: 0.9rem;">
              Think of it like having a digital filing cabinet that travels with you—no vendor lock-in, no lost access.
            </p>
          </div>
        </section>
      `;
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

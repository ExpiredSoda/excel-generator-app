// Debug: Track successful imports
console.log('✓ Navigation imports loaded:', {
  // All imports are now dynamic - loaded when needed
});

import { validateEmployeeEmail } from '../../shared/utils/validation.js';

// Sidebar navigation and dynamic content switching
export function setupNavigation() {
  console.log('🧭 Navigation: Setting up navigation...');
  const navItems = document.querySelectorAll('.nav-item');
  const mainContent = document.querySelector('.main-content');

  console.log('🧭 Navigation: Found elements:', {
    navItems: navItems.length,
    mainContent: !!mainContent
  });

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
    mainContent.innerHTML = `
      <section class="homepage-intro" style="max-width: 800px; margin: 0 auto 40px auto; background: #f8f9fa; border-radius: 16px; box-shadow: 0 2px 16px rgba(32,179,136,0.07); padding: 40px 32px 32px 32px; border: 1.5px solid #e1e8ed;">
        <h1 style="font-size:2.2rem; font-weight:800; color:#20b388; margin-bottom:18px; text-align:center;">Welcome to Free Excel Generators</h1>
        <p style="font-size:1.15rem; color:#222; line-height:1.7; margin-bottom:18px;">
          I’m just a real human—no AI startup, no SaaS company, no hidden subscriptions. I built this site as a personal project to sharpen my skills in JavaScript, HTML, and Excel, while also creating tools that actually help people. I’ve always been frustrated by websites that require logins just to download a spreadsheet, or by AI tools that either cost too much, have a steep learning curve, or don’t quite do what you need. This project is my way of pushing back on all that. It’s made for the folks stuck in admin roles or team support positions who never had time to conquer Excel but still want clean, functional spreadsheets without the headache. Maybe you just hate tools altogether. Maybe you want to impress your boss with a polished calendar or tracker and wink wink pretend you built it—go for it, I won’t stop you. Everything here is free, built to work entirely in your browser, and focused on being genuinely useful. If it helps you out, maybe consider buying me a coffee—or better yet, shoot me an email telling me how much you love the site. Who knows, you might just show up on the future testimonial wall. Thanks for being here
        </p>
        <p style="font-size:1.08rem; color:#444; line-height:1.7; margin-bottom:0;">
          If you want something custom or hand-tailored to your workflow, I do take custom orders—just send me an email. We can set up a time to talk it through, or hash it out the old-fashioned way: through the world’s fastest postal service—email. Got a suggestion? Want to send me praise, hate mail, or just test if I actually read my inbox? Maybe you’re a scam bot who wandered in by accident. Either way, I’m all ears. Shoot me a message—I read everything.
        </p>
      </section>
      <section class="homepage-updates" style="max-width: 800px; margin: 0 auto 40px auto; background: #fff; border-radius: 16px; box-shadow: 0 2px 16px rgba(32,179,136,0.04); padding: 32px 32px 24px 32px; border: 1.5px solid #e1e8ed;">
        <h3 style="font-size:1.25rem; font-weight:700; color:#20b388; margin-bottom:16px;">🔧 What’s Coming Next</h3>
        <ul style="font-size:1.05rem; color:#333; line-height:1.6; margin:0 0 12px 0; padding-left:20px;">
          <li>🆕 Attendance Tracker (coming soon!)</li>
          <li>✨ Improved styling and theme options</li>
          <li>📩 User-suggested features for all tools</li>
        </ul>
        <p style="font-size:1.05rem; color:#444; margin:0;">
          Got a feature idea or tool request? <a href="#" id="suggestLink" style="color:#20b388; text-decoration:underline;">Submit it here</a>.
        </p>
      </section>
    `;
  } else if (page === 'calendar') {
    const template = document.getElementById('calendarPageTemplate');
    if (template) {
      mainContent.innerHTML = template.innerHTML;
      // Wait for DOM to update, then setup calendar handlers
      setTimeout(() => {
        if (window.setupCalendarPage) {
          window.setupCalendarPage();
        }
      }, 100);
    } else {
      mainContent.innerHTML = '<h2>Error: Calendar template not found</h2>';
    }  } else if (page === 'attendance') {
    const attendanceTemplate = document.getElementById('attendancePageTemplate');
    if (attendanceTemplate) {
      mainContent.innerHTML = attendanceTemplate.innerHTML;
      // Wait for DOM to update, then dynamically import and setup attendance tracker
      setTimeout(() => {
        import('./attendanceTracker.js').then(({ setupAttendanceTracker }) => {
          setupAttendanceTracker();
        });
      }, 100);
    } else {
      mainContent.innerHTML = '<h2>Error: Attendance template not found</h2>';
    }
  } else if (page === 'meeting') {
    mainContent.innerHTML = `<h2>Meeting Tracker</h2><p>Coming soon! Create meeting tracking sheets here.</p>`;
  } else if (page === 'contact') {
    mainContent.innerHTML = `
      <section class="contact-section" style="max-width: 600px; margin: 0 auto 40px auto; background: #f8f9fa; border-radius: 16px; box-shadow: 0 2px 16px rgba(32,179,136,0.07); padding: 40px 32px 32px 32px; border: 1.5px solid #e1e8ed;">
        <h2 style="font-size:2rem; font-weight:800; color:#20b388; margin-bottom:18px; text-align:center;">Contact Me</h2>
        <form id="contactForm" style="display:flex; flex-direction:column; gap:18px;">
          <div style="display:flex; flex-direction:column; gap:6px;">
            <label for="contactName" style="font-weight:600; color:#2c3e50;">Name</label>
            <input type="text" id="contactName" name="name" required maxlength="60" style="padding:12px 16px; border:2px solid #e1e8ed; border-radius:8px; font-size:15px;">
          </div>
          <div style="display:flex; flex-direction:column; gap:6px;">
            <label for="contactEmail" style="font-weight:600; color:#2c3e50;">Email</label>
            <input type="email" id="contactEmail" name="email" required maxlength="100" style="padding:12px 16px; border:2px solid #e1e8ed; border-radius:8px; font-size:15px;">
          </div>
          <div style="display:flex; flex-direction:column; gap:6px;">
            <label for="contactMessage" style="font-weight:600; color:#2c3e50;">Message</label>
            <textarea id="contactMessage" name="message" required maxlength="1000" rows="6" style="padding:12px 16px; border:2px solid #e1e8ed; border-radius:8px; font-size:15px;"></textarea>
          </div>
          <button type="submit" class="generate-btn" style="align-self:flex-end; min-width:160px;">Send Message</button>
        </form>
        <div id="contactFormStatus" style="margin-top:18px; font-size:1.05rem;"></div>
      </section>
    `;
    // Add form handler for sending email (using mailto as fallback)
    setTimeout(() => {
      const form = document.getElementById('contactForm');
      const status = document.getElementById('contactFormStatus');
      if (form) {
        form.onsubmit = function(e) {
          e.preventDefault();
          const name = document.getElementById('contactName').value.trim();
          const email = document.getElementById('contactEmail').value.trim();
          const message = document.getElementById('contactMessage').value.trim();
          if (!name || !email || !message) {
            status.textContent = 'Please fill out all fields.';
            status.style.color = '#dc3545';
            return;
          }
          // Use shared validation for email
          const emailResult = validateEmployeeEmail(email);
          if (!emailResult.isValid) {
            status.textContent = emailResult.error || 'Invalid email address.';
            status.style.color = '#dc3545';
            return;
          }
          // Try to open mail client as fallback
          const mailto = `mailto:danielplanos@freeexcelgenerator.com?subject=Contact%20from%20${encodeURIComponent(name)}&body=${encodeURIComponent(message + '\n\nFrom: ' + name + ' (' + email + ')')}`;
          window.location.href = mailto;
          status.textContent = 'Thank you! Your message will be sent via your email client. I usually respond in about 1-2 business days.';
          status.style.color = '#20b388';
        };
      }
    }, 100);
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

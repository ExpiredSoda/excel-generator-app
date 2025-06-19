// Sidebar navigation and dynamic content switching
export function setupNavigation() {
  const navItems = document.querySelectorAll('.nav-item');
  const mainContent = document.querySelector('.main-content');

  console.log('Setting up navigation, found', navItems.length, 'nav items');

  navItems.forEach(item => {
    item.addEventListener('click', (e) => {
      console.log('Nav item clicked:', item.id);
      
      navItems.forEach(i => i.classList.remove('active'));
      item.classList.add('active');
      
      let page = item.id.replace('nav-', '');
      console.log('Navigating to page:', page);
      
      // Map new nav IDs to correct page names
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
      <h2>Welcome to Free Excel Generators!</h2>
      <p>
        This site offers free, easy-to-use tools for creating custom Excel resources like printable calendars and round robin tournament schedules.<br>
        Choose a tool from the sidebar to get started, customize it to your needs, and download your finished Excel file with just a click.
      </p>
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
    }
  } else if (page === 'attendance') {
    mainContent.innerHTML = `<h2>Attendance Tracker</h2><p>Coming soon! Create attendance tracking sheets here.</p>`;
  } else if (page === 'meeting') {
    mainContent.innerHTML = `<h2>Meeting Tracker</h2><p>Coming soon! Create meeting tracking sheets here.</p>`;
  } else {
    mainContent.innerHTML = '<h2>Page Not Found</h2>';
  }
}

// navigation.js: This logic is new in modular, not present in script.js. No monolithic logic to compare.
// color.js: rgbToHex matches script.js logic.
// sanitize.js: sanitizeLegendInput matches script.js logic.
// validation.js: validateLegendInput matches script.js logic.
// zipWriter.js: createZip matches script.js logic.
// No changes needed to any of these files.

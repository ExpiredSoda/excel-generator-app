// Wait until the HTML is fully loaded
document.addEventListener("DOMContentLoaded", function() {
  // Find all nav items and the main content area
  const navHome = document.getElementById("nav-home");
  const navCalendar = document.getElementById("nav-calendar");
  const navRoundRobin = document.getElementById("nav-roundrobin");
  const mainContent = document.querySelector(".main-content");
  const navItems = document.querySelectorAll(".nav-item");

  // Content templates for each page
  const pages = {
    home: `
      <h2>Welcome to Free Excel Generators!</h2>
      <p>
        This site offers free, easy-to-use tools for creating custom Excel resources like printable calendars and round robin tournament schedules.<br>
        Choose a tool from the sidebar to get started, customize it to your needs, and download your finished Excel file with just a click.
      </p>
    `,
    calendar: `
  <h2>Custom Excel Calendar Builder</h2>
  <form id="calendarForm">
    <label for="year">Year:</label>
    <input type="number" id="year" min="1900" max="2100" value="2024" required>
    <label for="month">Month:</label>
    <select id="month" required>
      <option value="0">January</option>
      <option value="1">February</option>
      <option value="2">March</option>
      <option value="3">April</option>
      <option value="4">May</option>
      <option value="5">June</option>
      <option value="6">July</option>
      <option value="7">August</option>
      <option value="8">September</option>
      <option value="9">October</option>
      <option value="10">November</option>
      <option value="11">December</option>
    </select>
    <button type="submit">Generate Calendar</button>
  </form>
  <div id="calendarPreview"></div>
`,
    roundrobin: `
      <h2>Round Robin Sports Scheduler</h2>
      <p>Coming soon: Generate balanced sports schedules and export to Excel.</p>
    `
  };

  // Helper function to show a page
  function showPage(page) {
    // Remove .active from all nav items
    navItems.forEach(item => item.classList.remove("active"));
    // Set .active on the correct nav item
    if (page === "home") navHome.classList.add("active");
    if (page === "calendar") navCalendar.classList.add("active");
    if (page === "roundrobin") navRoundRobin.classList.add("active");
    // Change the main content
    mainContent.innerHTML = pages[page];
  }

  // Event listeners for nav
  navHome.addEventListener("click", () => showPage("home"));
  navCalendar.addEventListener("click", () => showPage("calendar"));
  navRoundRobin.addEventListener("click", () => showPage("roundrobin"));

  // Listen for calendar form submission (dynamic content!)
mainContent.addEventListener("submit", function(event) {
  if (event.target && event.target.id === "calendarForm") {
    event.preventDefault();

    // Get year and month from the form
    const year = parseInt(document.getElementById("year").value, 10);
    const month = parseInt(document.getElementById("month").value, 10);

    // Generate calendar HTML
    const calendarHTML = generateCalendar(year, month);
    document.getElementById("calendarPreview").innerHTML = calendarHTML;
  }
});

// Function to create a simple calendar table
function generateCalendar(year, month) {
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const startDay = new Date(year, month, 1).getDay(); // 0=Sunday
  const monthNames = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  let html = `<h3>${monthNames[month]} ${year}</h3><table border="1" cellpadding="4"><tr>
    <th>Sun</th><th>Mon</th><th>Tue</th><th>Wed</th><th>Thu</th><th>Fri</th><th>Sat</th>
  </tr><tr>`;

  // Fill empty cells until first day
  for (let i = 0; i < startDay; i++) html += "<td></td>";

  // Fill the days of the month
  for (let day = 1; day <= daysInMonth; day++) {
    html += `<td>${day}</td>`;
    if ((startDay + day) % 7 === 0 && day !== daysInMonth) html += "</tr><tr>";
  }

  html += "</tr></table>";
  return html;
}
});
# Excel Generator App

A web app for generating custom Excel resources (calendars, schedules, and more) entirely client-side. Built from scratch for learning, portfolio, and practical use.

---

## Project Structure

- `index.html` — Main HTML page, with banner, sidebar navigation, and dynamic content area.
- `style.css` — All site styles: layout, sidebar, banner, responsive design.
- `script.js` — All logic: navigation, calendar builder, Excel XML/ZIP generation, and download.
- `/images/` — All icons and banner assets (SVGs).

---

## Features & Progress

- [x] Responsive site layout: banner, sidebar, main content area
- [x] Sidebar navigation with active/highlight states
- [x] Dynamic single-page navigation (no reloads)
- [x] Interactive calendar builder form (year/month/event rows)
- [x] Live HTML calendar preview
- [x] Downloadable Excel calendar (.xlsx) with:
  - Compact, professional grid (no empty rows)
  - Floating, styled legend panel (DrawingML, not worksheet cells)
  - User-selectable event rows per day (1–9)
  - User-tunable legend layout (panel/header/pills)
  - No Excel repair errors
- [x] All Excel file generation is 100% client-side (no server)
- [x] Modern, maintainable code with clear comments and section headers
- [x] Fully documented project structure and lessons learned

---

## Notes

- All layout and visual design is handled in `style.css` for maintainability.
- All Excel file generation (XML, ZIP) is handled in `script.js`.
- Banner and sidebar icons are in `/images` and referenced in HTML/CSS.
- No external dependencies or build tools required.
- Project is designed for readability, accessibility, and ease of future extension.

---

## Lessons Learned

- **Separation of Concerns:** HTML for structure, CSS for layout/appearance, JS for logic and interactivity.
- **Flexbox** is ideal for responsive layouts (sidebars, main content, navigation).
- **DOM manipulation** and event listeners enable dynamic, single-page app behavior.
- **Excel Open XML**: You can generate valid Excel files by hand-crafting XML and packaging with ZIP, as long as you follow the spec (no duplicate tags, correct relationships, etc.).
- **DrawingML**: Floating shapes (legend panels, pills) are possible in Excel by generating DrawingML and linking it via relationships.
- **Client-side ZIP**: You can create ZIP files in-browser with just JavaScript and basic byte manipulation.
- **Iterative Design:** Start simple, then refine layout, appearance, and features based on real output and user feedback.
- **Commenting and Sectioning:** Clear section headers and concise comments make code much easier to maintain and extend.
- **No need to memorize everything:** Use documentation, experiment, and keep a reference (like this README) for future work.

---

## Glossary

| Term                | Meaning                                                                                 |
|---------------------|-----------------------------------------------------------------------------------------|
| **HTML**            | HyperText Markup Language; the structure of web pages.                                  |
| **CSS**             | Cascading Style Sheets; controls appearance and layout of HTML elements.                |
| **JavaScript (JS)** | Programming language for interactivity and logic in web apps.                          |
| **DOM**             | Document Object Model; how JS accesses and manipulates HTML elements.                   |
| **Event Listener**  | JS code that responds to user actions (clicks, form submits, etc.).                    |
| **Flexbox**         | CSS layout mode for flexible, responsive designs.                                       |
| **Sidebar**         | Vertical navigation area, usually on the left.                                          |
| **Active/Highlight**| Visual state for the selected navigation item.                                          |
| **DrawingML**       | XML format for shapes/graphics in Excel (used for floating legend panel/pills).         |
| **Excel Open XML**  | The zipped XML file format used by modern Excel (.xlsx).                               |
| **ZIP**             | Compressed archive format; Excel files are ZIPs of XML and assets.                      |
| **Blob**            | JS object for handling binary data (used for downloads).                                |
| **Section Header**  | A comment block marking a major part of the code for clarity.                          |
| **Responsive**      | Design that adapts to different screen sizes.                                           |
| **Merge**           | In Excel, combining multiple cells into one (e.g., for headers).                        |
| **Client-side**     | All code runs in the browser; no server or backend required.                            |
| **Single-page app** | Web app that updates content dynamically without reloading the page.                    |

---

## How to Use / Extend

- To add new Excel tools (e.g., schedules), create a new navigation item, form, and generator function in `script.js`.
- To change the look, edit `style.css` (colors, spacing, layout).
- To update icons or banners, replace SVGs in `/images` and update references in HTML/CSS.
- Use this README as your reference for structure, terminology, and best practices.

---

## Author & License

- Created by [Your Name] as a learning and portfolio project.
- MIT License. Free to use, modify, and share.
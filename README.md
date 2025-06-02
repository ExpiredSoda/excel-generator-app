# Excel Generator App

A beginner-friendly web app for generating custom Excel tools (calendars, schedules, and more).  
This project is coded from scratch as a learning and portfolio experience, with step-by-step documentation.

## Project Setup

- Project folder structure created:
  - `index.html` (main page)
  - `style.css` (site styles)
  - `script.js` (site logic)
  - `/images` (all banners and icons)
- Assets designed and imported (icons, banner).
- Initialized git, created GitHub repository, pushed initial files.

## Progress Log

- [X] Site layout (banner, sidebar, main content area)
- [x] Updated HTML: Removed redundant banner text, added dynamic ad slot, clarified sidebar/content structure, and added homepage welcome/description.
- [x] Built project folder structure in VS Code and added initial assets (banner and icons).
- [x] Connected local project to GitHub with version control.
- [x] Created HTML skeleton with banner, sidebar, main content area, and placeholder navigation.
- [x] Documented and committed each feature and major layout change.
- [x] Implemented full-width banner using CSS background image, with custom height for best fit.
- [x] Refined sidebar navigation with added padding and spacing for usability.
- [x] Styled sidebar navigation with improved active/hover highlighting for accessibility and clarity.
- [x] Added and styled main content welcome panel.
- [x] Verified responsive layout and updated CSS for mobile/desktop flexibility.
- [x] Actively documenting lessons learned and coding concepts (HTML/CSS/box model/symbols).
- [x] Added interactive calendar builder form (year/month input, generate button).
- [x] Implemented dynamic calendar preview using JavaScript.
- [x] Learned about event listeners for forms and generating HTML dynamically.
- [x] Completed navigation logic: sidebar switches content and highlights selected tool.
- [x] Built and wired up a working calendar builder form (year/month), previewed as a dynamic HTML table.
- [x] Used JavaScript event listeners and DOM manipulation for a single-page app experience.
- [ ] Calendar tool form and UI
- [ ] HTML table preview for calendar
- [ ] CSV export functionality
- [ ] README updates at every major step

## Notes

- All layout, spacing, and visual design is achieved through external CSS (`style.css`) for maintainability.
- Banner image and sidebar icons are stored in the `/images` directory and referenced in the HTML and CSS.
- Project is designed for maximum readability, accessibility, and ease of use as both a learning project and portfolio piece.

## Lessons Learned

- No need to memorize all code—practice, experimentation, and reading docs is how developers really learn.
- Keeping a well-organized project structure (separate folders for images, CSS, etc.) makes development faster and less confusing.
- HTML provides the structure (“bones”) of the website, while CSS controls all layout, color, and visual style.
- Most site elements (like sidebars, banners) start off looking plain and unstyled until CSS is added.
- Flexbox is a powerful CSS layout tool for creating sidebars and responsive designs.
- “Padding” adds space inside an element’s border, while “margin” adds space outside it.
- It’s easier to iterate and make design tweaks with external CSS rather than inline styles.
- JavaScript lets you respond to user actions and update the page without reloading.
- DOM (Document Object Model) methods like `getElementById` and `querySelector` let you find and change elements in the HTML.
- Event listeners make your site interactive (e.g., `element.addEventListener("click", ...)`).
- Functions keep your code organized and reusable.

## Glossary of Terms Learned

| Term             | Meaning                                                                                    |
|------------------|--------------------------------------------------------------------------------------------|
| **HTML**         | HyperText Markup Language; the basic structure of web pages.                               |
| **CSS**          | Cascading Style Sheets; defines the appearance and layout of HTML elements.                |
| **Selector**     | In CSS, a way to target HTML elements (like `.sidebar` or `nav`).                          |
| **Class (`.`)**  | A reusable label for styling/grouping HTML elements, e.g., `<div class="sidebar">`.        |
| **Property**     | In CSS, a style rule (like `color`, `background`, `padding`).                              |
| **Value**        | The setting for a property (e.g., `color: #20b388;`).                                    |
| **Block `{}`**   | Groups properties in CSS or code statements in JS.                                         |
| **Semicolon `;`**| Ends a statement in CSS or JS.                                                             |
| **Padding**      | Space inside an element, between the border and content.                                   |
| **Margin**       | Space outside an element’s border, separating it from other elements.                      |
| **Flexbox**      | CSS layout mode for aligning and distributing space among items in a container.            |
| **Responsive**   | Designs that adapt to different screen sizes (mobile, tablet, desktop).                    |
| **Sidebar**      | A vertical navigation area, often on the left side of the screen.                          |
| **Active**       | The current or selected navigation item, highlighted for the user.                         |
| **Hover**        | The style shown when the mouse is over an element.                                         |
| **Commit**       | A saved change in your code tracked by git (version control).                              |
| **Push**         | Uploading your commits to GitHub or another remote repository.                             |
| **README.md**    | A markdown file describing your project, instructions, and progress log.                   |
| **Dom**          | Document Object Model; how JavaScript “sees” and changes HTML                              |
| **EventListener**| Code that waits for and responds to user actions (like clicks)                             |
|**Function**      | A reusable block of code that does something specific                                      |
| **innerHTML**    | The HTML content inside an element; you can set/change it                                  |
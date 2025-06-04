# Free Excel Generators

A web app for generating custom Excel resources entirely client-side. Currently features a powerful calendar generator with plans for additional tools like tournament schedulers.

---

## Project Structure

- `index.html` — Main HTML page with banner, sidebar navigation, and dynamic content area
- `style.css` — Complete site styling: layout, sidebar, banner, responsive design
- `script.js` — Core application logic: navigation, Excel generation, ZIP creation, and downloads
- `images/` — SVG icons and assets for the interface
- `README.md` — Project documentation and technical reference

---

## Features & Progress

### ✅ Completed Features
- **Responsive Web Interface**: Banner, sidebar navigation, and main content area
- **Single-Page Application**: Dynamic navigation without page reloads
- **Calendar Generator**: Interactive form with year/month selection
- **Event Configuration**: User-selectable event rows per day (1-9 slots)
- **Optional Tracker Sheet**: Checkbox to include event tracking worksheet
- **Live Preview**: HTML calendar preview before Excel generation
- **Excel Export**: Clean .xlsx files with three worksheets:
  - **Instructions Sheet**: User guide with merged cells and formatting
  - **Calendar Sheet**: Professional grid layout with cell-based legend
  - **Tracker Sheet**: Optional event counting and analytics
- **Client-Side Processing**: 100% browser-based, no server required
- **Error Prevention**: Eliminated Excel corruption issues through simplified architecture

### 🔄 Architecture Highlights
- **ExcelBuilder Library**: Custom classes for Excel XML generation
- **ZIP Generation**: Browser-based ZIP creation for .xlsx format
- **XML Escaping**: Proper handling of special characters
- **Color Consistency**: 9-color palette shared between calendar and tracker
- **Clean Code Structure**: Commented sections and maintainable functions

### 🚧 Planned Features
- **Round Robin Tournament Scheduler**: Balanced sports scheduling tool
- **Additional Excel Templates**: Expense trackers, project planners, etc.
- **Enhanced Customization**: User-defined color schemes and layouts

---

## Technical Implementation

### Excel Generation Process
1. **Form Input**: Collect year, month, event rows, and tracker preference
2. **XML Creation**: Generate all required Excel XML files using custom classes
3. **ZIP Assembly**: Package XML files into proper .xlsx structure
4. **Download**: Create browser download using Blob API

### Key Components
- **ExcelBuilder Classes**: `ExcelCell`, `ExcelRow`, `ExcelSheet`, `ExcelBuilder`
- **XML Generators**: Functions for workbook, worksheet, styles, and relationships
- **ZIP Writer**: Minimal ZIP creation without external dependencies
- **Event Handlers**: Form submission, navigation, and download management

### Excel File Structure
```
calendar.xlsx (ZIP container)
├── [Content_Types].xml
├── _rels/
│   └── .rels
└── xl/
    ├── workbook.xml
    ├── styles.xml
    ├── _rels/
    │   └── workbook.xml.rels
    └── worksheets/
        ├── sheet1.xml (Instructions)
        ├── sheet2.xml (Calendar)
        └── sheet3.xml (Tracker - optional)
```

---

## Development Notes

### Design Principles
- **Separation of Concerns**: HTML for structure, CSS for styling, JavaScript for logic
- **No External Dependencies**: Pure vanilla JavaScript and CSS
- **Client-Side Only**: All processing happens in the browser
- **Maintainable Code**: Clear section headers, comprehensive comments
- **Responsive Design**: Works on desktop and mobile devices

### Browser Compatibility
- **Modern Browsers**: Chrome, Firefox, Safari, Edge (ES6+ required)
- **JavaScript Features**: Classes, arrow functions, template literals, Map/Set
- **APIs Used**: DOM manipulation, Blob API, URL.createObjectURL

---

## Lessons Learned

### Technical Insights
- **Excel Open XML**: Hand-crafted XML generation is viable for creating valid .xlsx files
- **ZIP Generation**: Browser-based ZIP creation using byte manipulation
- **XML Escaping**: Critical for preventing corruption when handling user input
- **Cell-Based Legends**: Simpler and more reliable than complex DrawingML shapes
- **Event-Driven Architecture**: Clean separation between UI and business logic

### Best Practices Applied
- **Progressive Enhancement**: Start with working HTML/CSS, enhance with JavaScript
- **Error Prevention**: Validate inputs and handle edge cases gracefully
- **User Experience**: Live previews and clear feedback for all actions
- **Code Organization**: Logical grouping with clear section headers
- **Documentation**: Inline comments and comprehensive README

---

## Technical Reference

### Key Technologies
| Technology | Purpose | Implementation |
|------------|---------|----------------|
| **HTML5** | Page structure and semantic markup | Single-page layout with dynamic content |
| **CSS3** | Responsive styling and layout | Flexbox-based design, mobile-first approach |
| **JavaScript ES6+** | Application logic and Excel generation | Classes, modules, async operations |
| **Excel Open XML** | .xlsx file format specification | Hand-crafted XML for workbooks and worksheets |
| **ZIP Format** | Container for Excel files | Custom implementation using byte arrays |
| **DOM API** | User interface manipulation | Event handling, form processing, downloads |

### Color Palette
The app uses a consistent 9-color palette for calendar events:
```javascript
[
  "FFDC143C", // Crimson Red
  "FF228B22", // Forest Green  
  "FF1E90FF", // Dodger Blue
  "FFFFA500", // Orange
  "FF800080", // Purple
  "FFFFFF00", // Yellow
  "FF00CED1", // Dark Turquoise
  "FF8B4513", // Saddle Brown
  "FF4682B4"  // Steel Blue
]
```

---

## Usage Instructions

### For End Users
1. **Navigate**: Use sidebar to access the Calendar Generator
2. **Configure**: Select year, month, and number of event rows per day
3. **Preview**: Click "Generate Calendar" to see HTML preview
4. **Export**: Click "Download ZIP" to get your Excel file
5. **Customize**: Open in Excel to add events and customize colors

### For Developers
1. **Setup**: No build process required - open `index.html` in a browser
2. **Extend**: Add new generators by creating form templates and XML functions
3. **Customize**: Modify colors, layouts, and styles in `style.css`
4. **Debug**: Use browser developer tools - all errors logged to console

---

## Project Status

**Current Version**: 1.0 - Production Ready  
**Last Updated**: June 2025  
**Status**: ✅ Stable - Excel corruption issues resolved

### Recent Fixes
- ✅ Removed complex DrawingML to prevent Excel corruption
- ✅ Implemented robust cell-based legend system  
- ✅ Added proper XML escaping for special characters
- ✅ Fixed all JavaScript syntax errors
- ✅ Ensured consistent color palette across all sheets

---

## License & Attribution

**License**: MIT License - Free to use, modify, and distribute  
**Created by**: Daniel Planos  
**Purpose**: Learning project and portfolio demonstration  
**Repository**: Local development project
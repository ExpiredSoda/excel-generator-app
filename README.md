# Excel Generator App

A modern, modular web application for generating customizable Excel worksheets with advanced OOXML features. Built with vanilla JavaScript and a domain-driven architecture, this application provides powerful Excel generation capabilities with professional formatting, data validation, and dynamic formulas.

## 🚀 Features

### 📅 Smart Calendar Builder
- **Custom Monthly Calendars**: Generate Excel calendars for any year and month
- **Event Rows**: Configurable 1-9 event rows per day for scheduling
- **Optional Tracker Sheet**: Include a companion tracking worksheet
- **Custom Legend Colors**: Define color-coded categories with custom values
- **Professional Formatting**: Automatic sizing, styling, and conditional formatting
- **Data Validation**: Dropdown validation for consistent data entry

### 👥 Advanced Employee Shift Tracker
- **Employee Management**: Add employees with details (name, ID, email, title, phone)
- **Shift Scheduling**: Define custom shift times with break periods
- **Quick Templates**: Pre-defined shift presets (1st, 2nd, 3rd shift)
- **Color Coding**: Visual employee identification with custom colors
- **Time Tracking**: Weekly schedule grid with automatic total hours calculation
- **Legend System**: Custom status tracking with dropdown validation (Present, Absent, Holiday, etc.)
- **Reference Sheet**: Comprehensive employee directory with shift analytics
- **Dynamic Formulas**: Real-time legend usage analytics with COUNTIF functions
- **Data Validation**: Dropdown menus prevent data entry errors
- **Conditional Formatting**: Automatic highlighting for improved readability

### 🔧 Technical Excellence
- **Pure OOXML Generation**: Built from scratch without external Excel libraries
- **Advanced XML Features**: Proper namespace handling, merged cells, data validation
- **Performance Optimized**: Efficient memory usage and fast file generation
- **Professional Styling**: 27+ predefined styles with semantic meaning
- **Cross-Sheet Formulas**: Dynamic references and calculations
- **Data Integrity**: Input validation and sanitization throughout

### 🛠️ Meeting Tracker *(Coming Soon)*
- Meeting scheduling and tracking functionality

## 🏗️ Architecture

The application follows a clean, modular architecture with clear separation of concerns:

```
src/
├── web/                    # User Interface Layer
│   ├── ui/                 # UI Components & Navigation
│   │   ├── navigation.js   # Sidebar navigation and page routing
│   │   ├── calendarBuilder.js  # Calendar form handling
│   │   └── attendanceTracker.js # Employee management UI
│   └── utils/              # UI-specific utilities
│       ├── previewCalendar.js  # Calendar preview generation
│       └── previewEmployee.js  # Employee card previews
│
├── excel/                  # Excel Generation Layer
│   ├── core/               # Core Excel functionality
│   │   ├── excelBuilder.js # Main Excel workbook builder
│   │   ├── excelSheet.js   # Worksheet management
│   │   ├── excelCell.js    # Cell creation and formatting
│   │   ├── excelRow.js     # Row management
│   │   ├── conditionalFormatting.js # Conditional formatting rules
│   │   ├── dataValidation.js # Data validation and dropdown rules
│   │   ├── xmlUtils.js     # XML utility functions
│   │   └── index.js        # Core module exports
│   ├── generators/         # Specialized Excel generators
│   │   ├── calendar/       # Calendar-specific generation
│   │   └── attendance/     # Attendance tracker generation
│   │       ├── attendanceTrackerSheet.js # Main tracker sheet
│   │       ├── referenceSheet.js # Quick reference with analytics
│   │       ├── legendSheet.js # Legend definitions
│   │       ├── instructionsSheet.js # User instructions
│   │       ├── contentTypesXml.js # Content type definitions
│   │       ├── stylesXml.js # Style definitions
│   │       └── workbookXml.js # Workbook structure
│   └── utils/              # Excel utilities
│       └── zipUtils.js     # ZIP file creation for Excel format
│
├── presentation/           # Universal Styling & Formatting
│   ├── styles/             # Style definitions (27+ semantic styles)
│   │   ├── fonts.js        # Font configurations (Calibri, Arial, etc.)
│   │   ├── colors.js       # Color palette with semantic meaning
│   │   ├── fills.js        # Cell fill patterns and backgrounds
│   │   ├── borders.js      # Border styles and combinations
│   │   ├── styleIds.js     # Style ID management with constants
│   │   └── stylesXml.js    # XML style generation
│   ├── formatting/         # Cell formatting
│   │   ├── cellFormats.js  # Cell format definitions with alignment
│   │   └── dxfFormats.js   # Differential formatting for conditionals
│   ├── sizing/             # Layout and sizing
│   │   └── excelSizing.js  # Dynamic column widths and row heights
│   └── index.js            # Presentation layer exports
│
├── shared/                 # Cross-domain utilities
│   └── utils/
│       ├── validation.js   # Input validation patterns
│       ├── sanitize.js     # Data sanitization for XSS protection
│       └── timeUtils.js    # Time formatting and calculations
│
└── images/                 # UI assets and icons
```

## 🎨 Key Design Principles

### Universal Style System
- **Centralized Styling**: All Excel styles managed through `src/presentation/`
- **Semantic Style IDs**: Named constants for Instructions, Tables, Calendar, and Utility styles
- **Professional Formatting**: Consistent fonts, colors, borders, and alignment
- **Reusable Components**: Modular style definitions shared across generators
- **27+ Predefined Styles**: Complete style system covering all use cases

### Advanced OOXML Features
- **Data Validation**: Dropdown lists with formula references
- **Conditional Formatting**: Automatic highlighting and color coding
- **Merged Cells**: Professional table headers and titles
- **Cross-Sheet Formulas**: Dynamic COUNTIF and percentage calculations
- **Dynamic Sizing**: Automatic column width and row height optimization
- **Proper XML Structure**: Namespace compliance and Excel validation

### Domain-Driven Design
- **Clear Boundaries**: Separate domains for UI, Excel generation, presentation, and shared utilities
- **Dependency Direction**: Clean dependencies flowing inward toward core functionality
- **Modular Components**: Each module has a single responsibility
- **Testable Architecture**: Isolated components for easy testing

### Security & Validation
- **Input Sanitization**: All user inputs are sanitized before processing
- **XSS Protection**: Protection against script injection in Excel content
- **Safe XML Generation**: Proper XML escaping for all user-generated content
- **Data Validation**: Client-side and Excel-side validation rules

## 🚀 Quick Start

### Prerequisites
- Modern web browser with ES6 module support
- Local web server (for CORS compliance)

### Running the Application

#### Option 1: Python Server
```bash
# Navigate to project directory
cd excel-generator-app

# Python 3
python -m http.server 8000

# Python 2
python -m SimpleHTTPServer 8000

# Open http://localhost:8000/src/
```

#### Option 2: Node.js Server
```bash
# Install a simple HTTP server
npm install -g http-server

# Navigate to project directory
cd excel-generator-app

# Start server
http-server

# Open http://localhost:8080/src/
```

#### Option 3: VS Code Live Server
1. Install the "Live Server" extension in VS Code
2. Right-click on `src/index.html`
3. Select "Open with Live Server"

## 📖 Usage Guide

### Creating a Calendar
1. Click **Calendar Builder** in the sidebar
2. Select year and month
3. Choose number of event rows per day (1-9)
4. Optionally enable **Include Tracker Sheet**
5. Add custom legend colors and values if desired
6. Click **Generate Calendar** to create Excel file with data validation

### Managing Employee Shifts
1. Click **Attendance Tracker** in the sidebar
2. Fill in employee information (name and title required)
3. Set shift times or use quick presets
4. Choose a color code for the employee
5. Add custom legends for status tracking (Present, Absent, Holiday, etc.)
6. Click **Add Employee** to build your team
7. Click **Generate Tracker** to create comprehensive Excel workbook

### Understanding Generated Excel Files

#### Attendance Tracker Features:
- **Main Tracker Sheet**: Employee schedule grid with dropdown validation
- **Quick Reference Sheet**: Employee details with real-time legend usage analytics
- **Legends Sheet**: Centralized legend definitions with color coding
- **Instructions Sheet**: Comprehensive usage guide

#### Advanced Excel Features:
- **Data Validation**: Dropdown menus ensure consistent data entry
- **Dynamic Formulas**: Real-time calculations using COUNTIF and percentage formulas
- **Conditional Formatting**: Automatic highlighting for better readability
- **Merged Headers**: Professional table presentation
- **Cross-Sheet References**: Data consistency across all worksheets

## 🛠️ Development

### Adding New Features
1. **UI Components**: Add to `src/web/ui/`
2. **Excel Generators**: Create in `src/excel/generators/[feature]/`
3. **Shared Utilities**: Add to `src/shared/utils/`
4. **Styling**: Extend `src/presentation/styles/`

### Style System Usage
```javascript
// Import universal styles
import { STYLE_IDS, getCustomStyleId } from '../../../presentation/index.js';

// Use semantic style constants
const titleStyle = STYLE_IDS.TABLE_TITLE;        // Green background, center aligned
const headerStyle = STYLE_IDS.TABLE_HEADER;      // Green background, borders
const dataStyle = STYLE_IDS.TABLE_DATA;          // Plain with borders
const inputStyle = STYLE_IDS.TABLE_INPUT;        // Light blue for user input

// Create dynamic custom styles for legends
const customStyle = getCustomStyleId(colorIndex);
```

### Adding Data Validation
```javascript
import { DataValidationRule, generateDataValidationsXML } from '../core/dataValidation.js';

// Create dropdown validation
const rule = new DataValidationRule(
  "C2:C100",                           // Cell range
  "=Legends!$A$3:$A$10",              // Formula reference
  { showDropDown: true, allowBlank: false }
);

const validationXML = generateDataValidationsXML([rule]);
```

### Adding New Excel Generators
1. Create generator directory in `src/excel/generators/[name]/`
2. Implement main sheet builder with proper XML structure
3. Add supporting XML generators (styles, workbook, content types)
4. Use universal presentation system for consistent styling
5. Include data validation and conditional formatting
6. Export functions and integrate with UI

## 📁 File Structure Details

### Core Excel Components
- **ExcelBuilder**: Main workbook construction and ZIP generation
- **ExcelSheet**: Worksheet management with automatic sizing
- **ExcelCell**: Individual cell creation with formatting
- **ExcelRow**: Row management with height optimization
- **ConditionalFormatting**: Dynamic formatting rules and highlighting
- **DataValidation**: Dropdown validation and data integrity rules

### Presentation Layer
- **Fonts**: Typography definitions (Calibri 11pt, bold variants, etc.)
- **Colors**: Comprehensive color palette with semantic meaning
- **Fills**: Background patterns for different content types
- **Borders**: Professional border combinations
- **Sizing**: Dynamic column widths and row heights based on content

### Security Features
- **Sanitization**: Removes potentially harmful characters from all inputs
- **Validation**: Checks input patterns, formats, and data types
- **XML Escaping**: Prevents XML injection attacks in generated files
- **CORS Compliance**: Proper handling of cross-origin requests

## 🔧 Technical Stack

- **Frontend**: Vanilla JavaScript (ES6 modules)
- **Styling**: Pure CSS with BEM methodology and responsive design
- **Excel Format**: Office Open XML (.xlsx) with full OOXML compliance
- **Architecture**: Domain-driven design with modular components
- **Security**: Input sanitization, validation, and XSS protection
- **Performance**: Efficient memory usage and optimized file generation

## 🎯 Browser Support

- Chrome 61+
- Firefox 60+
- Safari 11+
- Edge 16+

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Follow the existing architecture patterns
4. Add comprehensive documentation
5. Test your changes thoroughly
6. Submit a pull request

## 🔮 Roadmap

- **Meeting Tracker**: Complete meeting scheduling functionality with calendar integration
- **Template System**: Save and reuse custom templates and configurations
- **Export Formats**: Support for additional file formats (CSV, PDF reports)
- **Advanced Formulas**: More complex calculations and business logic
- **Cloud Integration**: Save/load configurations from cloud storage
- **API Development**: RESTful API for programmatic access
- **Advanced Validation**: Custom validation rules and error handling
- **Performance Optimization**: Large dataset handling and memory efficiency

---

Built with ❤️ for creating professional Excel worksheets with enterprise-grade features and modern web technology.

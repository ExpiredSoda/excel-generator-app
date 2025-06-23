# Excel Generator App

A modern, modular web application for generating customizable Excel worksheets. Built with vanilla JavaScript and a domain-driven architecture, this application provides powerful Excel generation capabilities with an intuitive user interface.

## 🚀 Features

### 📅 Smart Calendar Builder
- **Custom Monthly Calendars**: Generate Excel calendars for any year and month
- **Event Rows**: Configurable 1-9 event rows per day for scheduling
- **Optional Tracker Sheet**: Include a companion tracking worksheet
- **Custom Legend Colors**: Define color-coded categories with custom values
- **Professional Formatting**: Automatic sizing, styling, and conditional formatting

### 👥 Employee Shift Tracker
- **Employee Management**: Add employees with details (name, ID, email, title, phone)
- **Shift Scheduling**: Define custom shift times with break periods
- **Quick Templates**: Pre-defined shift presets (1st, 2nd, 3rd shift)
- **Color Coding**: Visual employee identification with custom colors
- **Time Tracking**: Weekly schedule grid with total hours calculation
- **Reference Sheet**: Comprehensive employee directory

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
│   │   ├── xmlUtils.js     # XML utility functions
│   │   └── index.js        # Core module exports
│   ├── generators/         # Specialized Excel generators
│   │   ├── calendar/       # Calendar-specific generation
│   │   └── attendance/     # Attendance tracker generation
│   └── utils/              # Excel utilities
│       └── zipUtils.js     # ZIP file creation for Excel format
│
├── presentation/           # Universal Styling & Formatting
│   ├── styles/             # Style definitions
│   │   ├── fonts.js        # Font configurations
│   │   ├── colors.js       # Color palette
│   │   ├── fills.js        # Cell fill patterns
│   │   ├── borders.js      # Border styles
│   │   ├── styleIds.js     # Style ID management
│   │   └── stylesXml.js    # XML style generation
│   ├── formatting/         # Cell formatting
│   │   ├── cellFormats.js  # Cell format definitions
│   │   └── dxfFormats.js   # Differential formatting
│   ├── sizing/             # Layout and sizing
│   │   └── excelSizing.js  # Column widths and row heights
│   └── index.js            # Presentation layer exports
│
├── shared/                 # Cross-domain utilities
│   └── utils/
│       ├── validation.js   # Input validation
│       └── sanitize.js     # Data sanitization
│
└── images/                 # UI assets and icons
```

## 🎨 Key Design Principles

### Universal Style System
- **Centralized Styling**: All Excel styles managed through `src/presentation/`
- **Style ID Management**: Named constants instead of magic numbers
- **Reusable Components**: Modular style definitions shared across generators
- **Consistent Formatting**: Uniform appearance across all generated Excel files

### Domain-Driven Design
- **Clear Boundaries**: Separate domains for UI, Excel generation, presentation, and shared utilities
- **Dependency Direction**: Clean dependencies flowing inward toward core functionality
- **Modular Components**: Each module has a single responsibility

### Security & Validation
- **Input Sanitization**: All user inputs are sanitized before processing
- **XSS Protection**: Protection against script injection in Excel content
- **Safe XML Generation**: Proper XML escaping for all user-generated content

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
6. Click **Generate Calendar**
7. Download the generated Excel file

### Managing Employee Shifts
1. Click **Attendance Tracker** in the sidebar
2. Fill in employee information (name and title required)
3. Set shift times or use quick presets
4. Choose a color code for the employee
5. Click **Add Employee**
6. Repeat for all employees
7. Click **Generate Tracker** to create Excel file

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

// Use predefined styles
const headerStyle = STYLE_IDS.HEADER;

// Create custom styles
const customStyle = getCustomStyleId('myCustomStyle', customConfig);
```

### Adding New Excel Generators
1. Create generator directory in `src/excel/generators/[name]/`
2. Implement main sheet builder
3. Add supporting XML generators (styles, workbook, content types)
4. Export functions and integrate with UI
5. Use universal presentation system for styling

## 📁 File Structure Details

### Core Excel Components
- **ExcelBuilder**: Main workbook construction and ZIP generation
- **ExcelSheet**: Worksheet management with automatic sizing
- **ExcelCell**: Individual cell creation with formatting
- **ExcelRow**: Row management with height optimization
- **ConditionalFormatting**: Dynamic formatting rules

### Presentation Layer
- **Fonts**: Typography definitions (Calibri, Arial, etc.)
- **Colors**: Comprehensive color palette for themes
- **Fills**: Background patterns and gradients
- **Borders**: Border styles and thickness options
- **Sizing**: Column widths and row heights for different content types

### Security Features
- **Sanitization**: Removes potentially harmful characters
- **Validation**: Checks input patterns and formats
- **XML Escaping**: Prevents XML injection attacks

## 🔧 Technical Stack

- **Frontend**: Vanilla JavaScript (ES6 modules)
- **Styling**: Pure CSS with BEM methodology
- **Excel Format**: Office Open XML (.xlsx)
- **Architecture**: Domain-driven design with modular components
- **Security**: Input sanitization and validation

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

- **Meeting Tracker**: Complete meeting scheduling functionality
- **Template System**: Save and reuse custom templates
- **Export Formats**: Support for additional file formats
- **Advanced Formatting**: More styling and layout options
- **Cloud Integration**: Save/load from cloud storage
- **API Development**: RESTful API for programmatic access

---

Built with ❤️ for creating professional Excel worksheets with ease.

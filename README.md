# Excel Generator Web App

A modern web application for generating custom Excel files with smart calendars and schedules featuring advanced conditional formatting and customizable legends.

## ✨ Features

- **🗓️ Smart Calendar Builder**: Generate monthly calendars with intelligent conditional formatting
- **🎨 Custom Legend System**: Personalize category names with automatic color coding
- **📊 Auto-Tracking**: Optional tracker sheets with real-time formula-based counting
- **🔧 Round Robin Scheduler**: (Coming Soon) Create balanced sports schedules
- **💻 Modern UI**: Clean, responsive design with professional styling
- **📥 Excel Export**: Download as proper Excel (.xlsx) files with full formatting

## 🚀 Current Tools

### 📅 Calendar Builder
- **Flexible Dates**: Select any month and year (1900-2100)
- **Customizable Events**: Choose 1-9 event rows per day
- **Smart Legend System**: 
  - Customize category names (Meeting, Workout, Appointment, etc.)
  - Real-time input validation and sanitization
  - Visual color indicators matching Excel formatting
- **Conditional Formatting**: 
  - Automatic cell highlighting based on legend values
  - Case-insensitive text matching
  - 9 distinct color palette
- **Tracker Integration**: Optional automatic counting of legend entries
- **Professional Output**: Multi-sheet Excel files with instructions

## 🎯 Getting Started

1. **Launch**: Open `index.html` in any modern web browser
2. **Navigate**: Use the sidebar to access different tools
3. **Customize**: 
   - Set your desired month/year
   - Choose event rows per day
   - Personalize legend values with your own categories
4. **Generate**: Click "Generate Calendar" to create preview
5. **Download**: Get your custom Excel file with full formatting

## 🔧 Technical Highlights

### Core Technologies
- **Pure JavaScript**: No external dependencies
- **Advanced Excel Generation**: 
  - Valid OOXML format
  - Conditional formatting rules
  - Multi-sheet workbooks
  - Custom styling and DXF formatting
- **Security**: Input sanitization and XSS protection
- **ZIP Compression**: Proper .xlsx file structure

### Excel Features Generated
- **Conditional Formatting**: Auto-highlights cells matching legend values
- **Formula Integration**: Tracker sheets use COUNTIF formulas
- **Professional Styling**: Custom fonts, colors, and borders
- **Multi-Sheet Structure**: Instructions, Calendar, and optional Tracker
- **Cross-Platform**: Works in Excel, Google Sheets, and other spreadsheet apps

## 📁 File Structure

```
excel-generator-app/
├── index.html          # Main application page
├── style.css           # Modern responsive styling
├── script.js           # Core logic with ExcelBuilder library
├── images/             # UI icons (Calendar, Gear, Download SVGs)
│   ├── Calendar Icon.svg
│   ├── Gear Icon.svg
│   └── Download Icon.svg
└── README.md           # This documentation
```

## 🛡️ Security Features

- **Input Sanitization**: Removes HTML tags, scripts, and dangerous characters
- **Length Limits**: Prevents abuse with 50-character limits
- **Pattern Detection**: Identifies and blocks suspicious input patterns
- **XSS Protection**: Comprehensive validation for all user inputs

## 🎨 UI/UX Improvements

- **Modern Form Design**: Grid layout with professional styling
- **Visual Feedback**: Color indicators, hover effects, and smooth transitions
- **Responsive Layout**: Adapts to different screen sizes
- **Icon Integration**: SVG icons for better visual hierarchy
- **Real-time Validation**: Immediate feedback on legend inputs

## 🔄 Recent Updates

### Legend Customization System
- Dynamic legend field generation based on event rows
- Real-time input validation and sanitization
- Visual color indicators matching Excel palette
- Custom legend values integrated into Excel generation

### Enhanced UI
- Professional form styling with gradients and shadows
- Custom checkbox design
- Circular icon containers for headings
- Improved button designs with hover effects

### Excel Integration
- Conditional formatting rules using custom legend values
- Tracker sheet formulas reference user-defined categories
- Case-insensitive text matching in Excel
- Professional multi-sheet structure with instructions

## 🚀 Future Enhancements

- Round Robin tournament scheduler
- Additional calendar layouts and themes
- Export to other formats (PDF, CSV)
- Advanced scheduling features
- Team management tools
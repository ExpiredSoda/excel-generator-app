# Excel Generator App - UAT Version ✅

This is the **working** User Acceptance Testing (UAT) version of the Excel Generator App, featuring a modular architecture with separated concerns for better maintainability and testing.

## 🚀 Features

✅ **Smart Calendar Builder**: Generate custom Excel calendars with conditional formatting  
✅ **Legend Customization**: Add custom values and colors that automatically highlight matching calendar entries  
✅ **Color Picker Integration**: Visual color selection with real-time preview  
✅ **Merged Legend Display**: Clean legend appearance spanning both columns  
✅ **Tracker Integration**: Optional tracker sheet that automatically counts legend value occurrences  
✅ **Modern UI**: Clean, responsive interface with intuitive navigation  
✅ **Modular Architecture**: Separated concerns for easier maintenance and testing  
✅ **Cross-Platform**: Works in Excel, Google Sheets, and other spreadsheet applications  

## 📁 Project Structure

```
excel-generator-app-UAT/
├── index.html              # Main HTML entry point
├── main.js                 # Application initialization and module imports
├── style.css              # Global styles and responsive design
├── core/
│   └── excelCore.js       # Core Excel building classes (ExcelCell, ExcelRow, ExcelSheet, ConditionalFormattingRule)
├── generators/
│   ├── calendarSheet.js   # Calendar worksheet builder with conditional formatting
│   ├── stylesXml.js       # Excel styles.xml generator with custom colors
│   ├── contentTypesXml.js # Content types and relationships XML
│   ├── workbookXml.js     # Workbook structure XML
│   └── trackerSheet.js    # Tracker worksheet with formulas
├── ui/
│   ├── navigation.js      # Page navigation and routing
│   └── eventHandlers.js   # Form handling and user interactions
├── utils/
│   └── zipWriter.js       # ZIP file creation for Excel format
└── images/                # UI icons and assets
```

## 🔧 Technology Stack

- **Frontend**: Vanilla JavaScript ES6 modules, HTML5, CSS3
- **Excel Generation**: Custom XML generation with ZIP packaging
- **Architecture**: Modular design with separated concerns
- **Features**: Conditional formatting, custom colors, formula integration

## 🚀 Quick Start

1. **Open the application**: Simply open `index.html` in a modern web browser
2. **Navigate to Calendar Builder**: Click the "Calendar Builder" option in the sidebar
3. **Customize your calendar**:
   - Select year and month
   - Choose number of event rows per day (1-9)
   - Customize legend values and colors using the color pickers
   - Toggle tracker sheet inclusion
4. **Generate**: Click "Generate Calendar" to create preview
5. **Download**: Click "Download Excel File" to get your .xlsx file

## ✨ Key Features Explained

### Smart Conditional Formatting
- Type any legend value in calendar cells
- Cells automatically highlight with matching legend colors
- Case-insensitive matching (e.g., "meeting" = "Meeting" = "MEETING")

### Color Picker Integration
- Visual color selection for each legend entry
- Real-time preview of selected colors
- Validation to prevent duplicate colors
- Colors properly map to Excel ARGB format

### Merged Legend Display
- Legend entries span both I and J columns for clean appearance
- No awkward cut-off sections
- Proper center alignment and borders

### Tracker Sheet (Optional)
- Automatically counts occurrences of each legend value
- Uses Excel COUNTIF formulas for real-time updating
- Links to legend values for dynamic tracking

## 🧪 Testing Status

This UAT version has been tested and verified for:
- ✅ Excel file generation without corruption
- ✅ Conditional formatting functionality
- ✅ Custom color application
- ✅ Legend value customization
- ✅ Tracker sheet formulas
- ✅ Cross-browser compatibility
- ✅ Responsive design

## 🔄 Development Workflow

The modular architecture allows for easy:
- **Testing**: Individual components can be tested in isolation
- **Maintenance**: Changes to one module don't affect others
- **Extension**: New generators can be added easily
- **Debugging**: Clear separation of concerns

## 📊 Module Dependencies

```
main.js
├── ui/navigation.js
├── ui/eventHandlers.js
│   ├── generators/calendarSheet.js
│   │   └── core/excelCore.js
│   ├── generators/stylesXml.js
│   ├── generators/contentTypesXml.js
│   ├── generators/workbookXml.js
│   ├── generators/trackerSheet.js
│   └── utils/zipWriter.js
```

## 🚧 Future Enhancements

Potential areas for expansion:
- Additional calendar layouts (weekly, yearly)
- More tracker sheet options
- Template system for common use cases
- Export to other formats (PDF, CSV)
- Advanced styling options

---

**Status**: ✅ Fully Functional  
**Last Updated**: Current  
**Next Step**: Refactor excel-generator-lib for library distribution
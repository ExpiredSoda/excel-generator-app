# Excel Generator Library

A lightweight, dependency-free JavaScript library for generating Excel (.xlsx) files with advanced features like conditional formatting, custom styling, and formula support.

## 🚀 Features

- ✅ **Zero Dependencies**: Pure JavaScript with no external libraries
- ✅ **Excel Compatibility**: Generates valid .xlsx files that work in Excel, Google Sheets, and other spreadsheet applications  
- ✅ **Conditional Formatting**: Smart highlighting based on cell values
- ✅ **Custom Styling**: Colors, fonts, borders, and alignment
- ✅ **Formula Support**: Excel formulas for dynamic calculations
- ✅ **Modular Architecture**: Use only the components you need
- ✅ **No Build Process**: Direct ES6 imports, works immediately

## 📦 Installation

### ES6 Modules
Import individual modules as needed:

```javascript
import { ExcelBuilder, ExcelSheet, ExcelCell } from './lib/index.js';
import { createZip } from './lib/utils/zipWriter.js';
```

## 🚀 Quick Start

### Easy Method: Complete Excel Generation

```javascript
import { generateCompleteExcel, createZip } from './lib/index.js';

// Generate calendar with all options
const options = {
  year: 2024,
  month: 0, // January
  eventRows: 3,
  includeTracker: true, // Add tracker sheet
  legendValues: ['Meeting', 'Holiday', 'Personal'],
  customColors: ['FFDC143C', 'FF228B22', 'FF1E90FF']
};

const files = generateCompleteExcel(options);
const zipBytes = createZip(files);

// Download
const blob = new Blob([zipBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'calendar.xlsx';
a.click();
```

### Advanced Method: Manual Component Generation

```javascript
import { 
  buildCalendarSheetWithExcelBuilder,
  getStylesXML,
  getContentTypesXML,
  getWorkbookXML,
  getTrackerSheetXML,
  createZip
} from './lib/index.js';

// Generate individual components
const includeTracker = true;
const legendValues = ['Meeting', 'Holiday', 'Personal'];
const customColors = ['FFDC143C', 'FF228B22', 'FF1E90FF'];

const calendarXML = buildCalendarSheetWithExcelBuilder(2024, 0, 3, false, legendValues, customColors);
const stylesXML = getStylesXML(3, customColors);
const contentTypesXML = getContentTypesXML(includeTracker);
const workbookXML = getWorkbookXML(includeTracker);

// Base files
const files = [
  { name: '[Content_Types].xml', content: contentTypesXML },
  { name: 'xl/worksheets/sheet1.xml', content: calendarXML },
  { name: 'xl/styles.xml', content: stylesXML },
  { name: 'xl/workbook.xml', content: workbookXML }
];

// Add tracker if requested
if (includeTracker) {
  const trackerXML = getTrackerSheetXML(legendValues);
  files.push({ name: 'xl/worksheets/sheet2.xml', content: trackerXML });
}

const zipBytes = createZip(files);
// Use zipBytes for download...
```

## 📊 Tracker Sheet Feature

The library includes an optional **tracker sheet** that automatically counts occurrences of each legend value in your calendar.

### How it Works:
1. **Enable tracker**: Set `includeTracker: true` in options
2. **Automatic counting**: Uses Excel COUNTIF formulas to count legend values
3. **Real-time updates**: Counts update automatically when you modify calendar entries
4. **Visual summary**: Shows total count for each legend category

### Tracker Sheet Contents:
- **Legend Value**: Lists each custom legend value
- **Color**: Shows the associated color 
- **Count**: Formula that counts occurrences in the calendar sheet
- **Percentage**: Shows what percentage each category represents

### Example Tracker Output:
```
Legend Value    Count    Percentage
Meeting         8        40%
Holiday         3        15%
Personal        9        45%
```

### Usage:
```javascript
// Generate calendar with tracker
const files = generateCompleteExcel({
  year: 2024,
  month: 0,
  includeTracker: true, // This adds the tracker sheet
  legendValues: ['Meeting', 'Holiday', 'Personal'],
  customColors: ['FFDC143C', 'FF228B22', 'FF1E90FF']
});

// Result: Excel file with 2 sheets
// Sheet 1: Calendar with conditional formatting
// Sheet 2: Tracker with automatic counting
```

## 📋 API Reference

### ExcelBuilder

```javascript
const builder = new ExcelBuilder();
builder.addSheet(sheet);              // Add a worksheet
builder.setStyles(stylesXML);         // Set custom styles
builder.generateFiles();              // Generate all Excel files
```

### ExcelSheet  

```javascript
const sheet = new ExcelSheet(name);
sheet.addRow(row);                    // Add a row
sheet.addMerge(range);                // Merge cells (e.g., "A1:B1")
sheet.setCols(colDefs);               // Set column widths
sheet.addConditionalFormatting(rule); // Add conditional formatting
```

### ExcelCell Helper Methods

```javascript
// Different cell types
const textCell = ExcelCell.text('A', 1, 'Hello');
const numberCell = ExcelCell.number('B', 1, 42);
const formulaCell = ExcelCell.formula('C', 1, '=A1+B1');
const emptyCell = ExcelCell.empty('D', 1, 2); // style 2
```

## 🌐 Browser Support

- ✅ Chrome 60+
- ✅ Firefox 60+  
- ✅ Safari 12+
- ✅ Edge 79+

Requires ES6 module support.

## 📄 License

MIT License - feel free to use in both personal and commercial projects.

---

**No npm install required** - Just copy the lib folder and start using!
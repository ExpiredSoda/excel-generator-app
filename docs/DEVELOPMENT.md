# Development Guidelines 🔧

## Development Workflow

The `/src` folder is the **main development environment** for the Excel Generator project. This modular architecture allows for rapid feature development and easy maintenance.

## 🏗️ Architecture Overview

```
src/
├── main.js                     # Application entry point & module coordination
├── index.html                 # Main UI & page structure
├── style.css                  # Global styles & responsive design
├── core/
│   └── excelCore.js           # Core Excel building classes
├── generators/
│   ├── calendarSheet.js       # Calendar worksheet generation
│   ├── stylesXml.js           # Excel styles with custom colors
│   ├── contentTypesXml.js     # Content types & relationships
│   ├── workbookXml.js         # Workbook structure
│   └── trackerSheet.js        # Tracker formulas & counting
├── ui/
│   ├── navigation.js          # Page routing & navigation
│   └── eventHandlers.js       # Form handling & user interactions
├── utils/
│   └── zipWriter.js           # ZIP file creation for Excel
└── images/                    # UI icons & assets
```

## 🚀 Adding New Features

### 1. Adding New Generators
Create new XML generators in `/generators`:

```javascript
// generators/newFeature.js
export function generateNewFeatureXML(options) {
  // Your generator logic
  return xmlString;
}
```

Import in `main.js`:
```javascript
import { generateNewFeatureXML } from './generators/newFeature.js';
```

### 2. Adding UI Components
Add new UI handlers in `/ui`:

```javascript
// ui/newComponent.js
export function setupNewComponent() {
  // UI setup logic
}
```

### 3. Adding Core Functionality
Extend core classes in `/core/excelCore.js`:

```javascript
export class NewExcelFeature {
  constructor(options) {
    // Core functionality
  }
}
```

## 🧪 Testing Workflow

1. **Development**: Make changes in `/src`
2. **Test**: Verify functionality works as expected
3. **Validate**: Ensure Excel files open without corruption
4. **Cross-browser**: Test in Chrome, Firefox, Safari, Edge
5. **Update Library**: Extract reusable components to `/lib` if needed

## 🔄 Module Dependencies

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

## 📱 Feature Roadmap

### Immediate (Current Sprint)
- ✅ Conditional formatting working
- ✅ Custom colors integration
- ✅ Merged legend display

### Next Sprint
- [ ] Additional calendar layouts (weekly, yearly)
- [ ] More tracker sheet options
- [ ] Template system for common use cases
- [ ] Advanced styling options

### Future Enhancements
- [ ] Export to other formats (PDF, CSV)
- [ ] Recurring event support
- [ ] Multi-sheet workbooks
- [ ] Formula builder UI

## 🛠️ Code Standards

### JavaScript
- Use ES6 modules
- Prefer `const` over `let` over `var`
- Use descriptive function and variable names
- Add JSDoc comments for public APIs

### File Organization
- One main class/function per file
- Group related utilities together
- Keep generators focused on single responsibility
- Maintain clear separation between UI and logic

### XML Generation
- Always escape user input with `escapeXml()`
- Use proper ARGB color formatting
- Validate Excel formulas before generating
- Maintain proper XML structure and namespaces

## 🐛 Debugging Tips

### Excel File Corruption
1. Check XML syntax with browser dev tools
2. Validate color format (ARGB with FF prefix)
3. Ensure proper ZIP file structure
4. Test UTF-8 encoding in zipWriter

### Conditional Formatting Issues
1. Verify DXF styles match fill colors exactly
2. Check formula references (`$I$2` format)
3. Ensure range references are correct
4. Test with simple values first

### Performance
1. Minimize DOM updates in event handlers
2. Cache color picker values
3. Batch XML generation operations
4. Use efficient string concatenation

## 📦 Deployment

For deployment, the entire `/src` folder can be:
1. Served as static files
2. Deployed to any web server
3. Used with GitHub Pages
4. Integrated into larger applications

No build process required - pure JavaScript ES6 modules!
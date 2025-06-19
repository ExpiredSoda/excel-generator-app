# Excel Generator Project 📊

A comprehensive Excel file generation solution that creates smart calendars with conditional formatting, custom styling, and formula support. Built with pure JavaScript and zero dependencies.

## 🎯 Quick Start

**👉 Want to use the app?** Open `/src/index.html` in your browser!  
**👉 Want to develop with it?** Check out the `/lib` folder for reusable components!  
**👉 Want to see the evolution?** Compare `/archive` with `/src` to see the refactoring benefits!  

## 📁 Repository Structure

This repository contains **three complete versions** of the Excel Generator:

```
📦 excel-generator-app/
├── 🚀 src/                     # CURRENT VERSION - Use this!
│   ├── index.html              # 👈 Main application - Open this file!
│   ├── main.js                 # Modular architecture
│   ├── style.css               # Modern UI
│   ├── core/                   # Excel building classes
│   ├── generators/             # XML generators
│   ├── ui/                     # User interface modules
│   ├── utils/                  # Utility functions
│   └── images/                 # UI assets
│
├── 🗄️ archive/                 # Original monolithic version (preserved)
│   ├── script.js               # Single 800+ line file
│   ├── index.html              # Original HTML
│   └── style.css               # Original styles
│
├── 📚 lib/                     # Reusable library for developers
│   ├── index.js                # Clean API for integration
│   ├── core/                   # Core classes
│   ├── generators/             # Generator modules
│   └── utils/                  # Utility functions
│
└── 📖 docs/                    # Project documentation
    ├── DEVELOPMENT.md          # How to add features
    └── DEPLOYMENT.md           # How to deploy
```

## ✨ Features (All Versions)

- 🗓️ **Smart Calendar Generation** with conditional formatting
- 🎨 **Visual Color Picker** with real-time preview  
- 📋 **Custom Legend Values** with automatic highlighting
- 📊 **Tracker Integration** with Excel formulas
- 📱 **Cross-Platform** - Works in Excel, Google Sheets, and other apps
- 🚫 **Zero Dependencies** - Pure JavaScript implementation

## 🚀 Which Version Should I Use?

### 👥 **For End Users**
**Use `/src/index.html`** - The current, feature-complete application
- ✅ Latest features and bug fixes
- ✅ Modern modular architecture  
- ✅ Easy to customize and extend
- ✅ Actively maintained

### 👨‍💻 **For Developers** 
**Use `/lib/`** - Clean, reusable components
```javascript
import { ExcelBuilder, CalendarGenerator } from './lib/index.js';

const calendar = new CalendarGenerator({
  year: 2024,
  month: 0,
  legendValues: ['Meeting', 'Holiday'],
  legendColors: ['FFDC143C', 'FF228B22']
});
```

### 📚 **For Learning/Reference**
**Check `/archive/`** - Original monolithic version
- See the evolution from monolithic to modular
- Understand refactoring benefits
- Historical reference for the project

## 🔄 Project Evolution

```
🗄️ Archive (v1.0)          🚀 Current (v2.0)
├── script.js (800+ lines)  ├── main.js (clean entry point)
├── index.html              ├── core/ (Excel classes)
└── style.css               ├── generators/ (XML creation)
                            ├── ui/ (interface logic)
                            └── utils/ (helper functions)

❌ Hard to maintain         ✅ Easy to extend
❌ Difficult to test        ✅ Modular testing  
❌ Code duplication         ✅ Clean separation
❌ Merge conflicts          ✅ Parallel development
```

## 🛠️ Development & Deployment

- **Development**: See `/docs/DEVELOPMENT.md` for adding features
- **Deployment**: See `/docs/DEPLOYMENT.md` for hosting options
- **Library Usage**: See `/lib/README.md` for integration guide

## 🧪 Testing

Each folder contains a complete, working version:
- **Archive**: `open archive/index.html` (works with script.js)
- **Current**: `open src/index.html` (works with modular files)
- **Library**: See examples in `/lib/examples/`

## 📈 Future Roadmap

- [ ] Additional calendar layouts (weekly, yearly)
- [ ] More tracker sheet options  
- [ ] Template system for common use cases
- [ ] Export to other formats (PDF, CSV)
- [ ] Advanced styling options
- [ ] Recurring event support

## 📄 License

MIT License - See individual folders for specific details.

---

**🎯 TL;DR**: Open `/src/index.html` to use the app, check `/lib/` to integrate it, and peek at `/archive/` to see where we started!
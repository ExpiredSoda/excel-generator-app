# Library Verification Guide 🧪

## 🚀 How to Verify the Library Works

### Method 1: Quick Demo Test
1. **Serve the library folder** with a local server:
   ```bash
   # Using Python
   cd lib
   python -m http.server 8000
   
   # Using Node.js (if you have it)
   npx serve .
   
   # Using VS Code Live Server extension
   # Right-click quick-demo.html -> "Open with Live Server"
   ```

2. **Open the demo page**: `http://localhost:8000/quick-demo.html`
3. **Click "Generate & Download Calendar"**
4. **Verify the Excel file opens correctly** in Excel/Google Sheets

### Method 2: Comprehensive Testing
1. **Open**: `http://localhost:8000/test-library.html`
2. **Run all tests** by clicking each test button
3. **Verify all tests pass** (should show green ✅)
4. **Download the test file** and open in Excel

### Method 3: Manual Integration Test
Create a simple HTML file that imports your library:

```html
<!DOCTYPE html>
<html>
<head><title>Library Test</title></head>
<body>
    <button onclick="test()">Test Library</button>
    <script type="module">
        import { 
            buildCalendarSheetWithExcelBuilder, 
            createZip 
        } from './lib/index.js';
        
        window.test = async function() {
            try {
                const xml = buildCalendarSheetWithExcelBuilder(2024, 0, 2);
                console.log('✅ Library works!', xml.length);
                alert('✅ Library imported and working!');
            } catch (error) {
                console.error('❌ Library error:', error);
                alert('❌ Library error: ' + error.message);
            }
        };
    </script>
</body>
</html>
```

## 🔍 What to Verify

### ✅ Import Success
- [ ] All modules import without errors
- [ ] Functions are available and callable
- [ ] No console errors during import

### ✅ Calendar Generation
- [ ] Calendar XML is generated (>1000 bytes)
- [ ] Contains proper Excel worksheet structure
- [ ] Includes conditional formatting rules
- [ ] Legend values are present

### ✅ File Creation
- [ ] ZIP file is created successfully
- [ ] File size is reasonable (>5KB)
- [ ] Downloaded file has .xlsx extension

### ✅ Excel Compatibility
- [ ] File opens in Microsoft Excel without errors
- [ ] File opens in Google Sheets
- [ ] Conditional formatting works (cells highlight when typing legend values)
- [ ] Colors match the selected values

## 🐛 Common Issues & Solutions

### Import Errors
```
❌ Failed to resolve module specifier
```
**Solution**: Serve files through HTTP server, not file:// protocol

### CORS Errors
```
❌ Access to script blocked by CORS policy
```
**Solution**: Use a local server instead of opening HTML directly

### Missing Functions
```
❌ buildCalendarSheetWithExcelBuilder is not a function
```
**Solution**: Check that all generators are copied from /src

### Invalid Excel Files
```
❌ Excel says file is corrupted
```
**Solution**: Verify ZIP structure and XML syntax

## 🎯 Success Criteria

Your library is working correctly if:

1. **✅ No JavaScript errors** in browser console
2. **✅ All test functions return expected results**
3. **✅ Generated Excel files open without corruption**
4. **✅ Conditional formatting highlights cells correctly**
5. **✅ Custom colors appear as selected**

## 📊 Performance Benchmarks

Expected performance:
- **Import time**: <100ms
- **Calendar generation**: <500ms  
- **ZIP creation**: <200ms
- **File size**: 15-30KB for typical calendar

## 🔄 Continuous Verification

Add this to your development workflow:
1. **Before commits**: Run quick-demo.html test
2. **Before releases**: Run full test-library.html suite
3. **Regular checks**: Verify in different browsers
4. **Excel validation**: Test in both Excel and Google Sheets

---

**Goal**: Ensure developers can integrate your library with confidence! 🎯
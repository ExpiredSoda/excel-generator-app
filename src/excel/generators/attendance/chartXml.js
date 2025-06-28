// generators/attendance/chartXml.js
// Generates chart XML for attendance tracker legend usage

import { PieChart, ChartSeries, ChartDataRange } from '../../core/index.js';

/**
 * Generate attendance legend usage pie chart
 * @param {Array} legends - Array of legend objects with label and color
 * @param {string} sheetName - Name of the Quick Reference sheet
 * @returns {string} Chart XML
 */
export function buildAttendanceChart(legends = [], sheetName = "Quick Reference") {
  if (legends.length === 0) {
    throw new Error('Cannot create chart without legends data');
  }

  // Create pie chart instance
  const chart = new PieChart('Legend Usage Analytics');
  chart.setLegend('r', false); // Right legend, no overlay (matches working chart)

  // Remove custom color setting since we're using Excel's theme colors only
  // The chart will use accent6 with different shades/tints

  // Calculate data range based on legends
  const firstDataRow = 4; // Legends data starts at row 4
  const lastDataRow = firstDataRow + legends.length - 1;

  // Create data ranges for categories (legend labels) and values (usage counts)
  const categoriesRange = new ChartDataRange(
    sheetName, 
    `F${firstDataRow}`, 
    `F${lastDataRow}`, 
    true // string data
  );

  const valuesRange = new ChartDataRange(
    sheetName, 
    `G${firstDataRow}`, 
    `G${lastDataRow}`, 
    false // numeric data
  );

  // Create chart series
  const series = new ChartSeries('Usage', categoriesRange, valuesRange);
  series.setIndexOrder(0, 0);
  chart.addSeries(series);

  // Generate and return chart XML
  return chart.toXML();
}

/**
 * Generate chart colors XML file
 * @returns {string} Chart colors XML
 */
export function buildChartColorsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">
  <a:schemeClr val="accent1"/>
  <a:schemeClr val="accent2"/>
  <a:schemeClr val="accent3"/>
  <a:schemeClr val="accent4"/>
  <a:schemeClr val="accent5"/>
  <a:schemeClr val="accent6"/>
</cs:colorStyle>`;
}

/**
 * Generate chart style XML file
 * @returns {string} Chart style XML
 */
export function buildChartStyleXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="102">
  <cs:dataPoint>
    <cs:lnRef idx="1"/>
    <cs:fillRef idx="1"/>
    <cs:effectRef idx="0"/>
  </cs:dataPoint>
  <cs:legend>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:defRPr sz="900" kern="1200"/>
  </cs:legend>
  <cs:title>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:defRPr sz="1400" kern="1200"/>
  </cs:title>
</cs:chartStyle>`;
} 
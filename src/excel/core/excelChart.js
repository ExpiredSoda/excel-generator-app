// excel/core/excelChart.js
// Base chart functionality for Excel generation

import { escapeXml } from './xmlUtils.js';

/**
 * Base class for Excel charts
 */
export class ExcelChart {
  constructor(title = '', type = 'pieChart') {
    this.title = title;
    this.type = type;
    this.series = [];
    this.legend = { position: 'b', overlay: false }; // bottom by default
    this.plotArea = { layout: {} };
    this.varyColors = true;
    this.showDataLabels = false;
  }

  /**
   * Add a data series to the chart
   * @param {ChartSeries} series - Chart data series
   */
  addSeries(series) {
    this.series.push(series);
  }

  /**
   * Set legend configuration
   * @param {string} position - Legend position: 'b', 't', 'l', 'r'
   * @param {boolean} overlay - Whether legend overlays the chart
   */
  setLegend(position = 'b', overlay = false) {
    this.legend = { position, overlay };
  }

  /**
   * Generate chart XML structure - matches working Excel chart exactly
   * @returns {string} Chart XML
   */
  toXML() {
    const titleXML = this.generateTitleXML();
    const plotAreaXML = this.generatePlotAreaXML();
    const legendXML = this.generateLegendXML();

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="108"/></mc:Choice><mc:Fallback><c:style val="8"/></mc:Fallback></mc:AlternateContent><c:chart>${titleXML}<c:autoTitleDeleted val="0"/>${plotAreaXML}${legendXML}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:noFill/><a:round/></a:ln><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw><a:softEdge rad="63500"/></a:effectLst></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
  }

  /**
   * Generate title XML
   * @returns {string} Title XML
   */
  generateTitleXML() {
    if (!this.title) {
      return '<c:autoTitleDeleted val="1"/>';
    }

    return `<c:title><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:title>`;
  }

  /**
   * Generate legend XML - matches working Excel chart styling
   * @returns {string} Legend XML
   */
  generateLegendXML() {
    return `<c:legend><c:legendPos val="${this.legend.position}"/><c:overlay val="${this.legend.overlay ? '1' : '0'}"/><c:spPr><a:solidFill><a:schemeClr val="lt1"><a:lumMod val="95000"/><a:alpha val="39000"/></a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:ln w="0" cap="sq"><a:noFill/></a:ln><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:legend>`;
  }

  /**
   * Generate plot area XML - to be overridden by chart type implementations
   * @returns {string} Plot area XML
   */
  generatePlotAreaXML() {
    throw new Error('generatePlotAreaXML must be implemented by chart type subclasses');
  }
}

/**
 * Chart data series class
 */
export class ChartSeries {
  constructor(name, categories, values) {
    this.name = name;
    this.categories = categories; // ChartDataRange
    this.values = values; // ChartDataRange
    this.index = 0;
    this.order = 0;
  }

  /**
   * Set series index and order
   * @param {number} index - Series index
   * @param {number} order - Series order
   */
  setIndexOrder(index, order = index) {
    this.index = index;
    this.order = order;
  }
}

/**
 * Chart data range class for referencing worksheet cells
 */
export class ChartDataRange {
  constructor(sheetName, startCell, endCell, isStringData = false) {
    this.sheetName = sheetName;
    this.startCell = startCell;
    this.endCell = endCell;
    this.isStringData = isStringData;
  }

  /**
   * Get Excel formula reference
   * @returns {string} Excel formula (e.g., 'Sheet1'!$A$1:$A$10)
   */
  getFormulaReference() {
    const range = this.startCell === this.endCell ? this.startCell : `${this.startCell}:${this.endCell}`;
    return `'${this.sheetName}'!$${range.replace(/:/g, ':$')}`;
  }

  /**
   * Generate XML reference
   * @returns {string} Reference XML
   */
  toXML() {
    const refType = this.isStringData ? 'strRef' : 'numRef';
    const cacheType = this.isStringData ? 'strCache' : 'numCache';
    
    // Generate cache data based on range
    let cacheXML = '';
    if (this.isStringData) {
      // For string data, create placeholder cache entries
      const count = this.getRowCount();
      cacheXML = `<c:${cacheType}><c:ptCount val="${count}"/>`;
      for (let i = 0; i < count; i++) {
        cacheXML += `<c:pt idx="${i}"><c:v>Legend ${i + 1}</c:v></c:pt>`;
      }
      cacheXML += `</c:${cacheType}>`;
    } else {
      // For numeric data, create placeholder cache entries
      const count = this.getRowCount();
      cacheXML = `<c:${cacheType}><c:formatCode>General</c:formatCode><c:ptCount val="${count}"/>`;
      for (let i = 0; i < count; i++) {
        cacheXML += `<c:pt idx="${i}"><c:v>0</c:v></c:pt>`;
      }
      cacheXML += `</c:${cacheType}>`;
    }
    
    return `<c:${refType}><c:f>${escapeXml(this.getFormulaReference())}</c:f>${cacheXML}</c:${refType}>`;
  }

  /**
   * Calculate number of rows in the range
   * @returns {number} Row count
   */
  getRowCount() {
    if (this.startCell === this.endCell) {
      return 1;
    }
    
    // Extract row numbers from cell references (e.g., F4:F11 -> 8 rows)
    const startRow = parseInt(this.startCell.replace(/[A-Z]/g, ''));
    const endRow = parseInt(this.endCell.replace(/[A-Z]/g, ''));
    return endRow - startRow + 1;
  }
} 
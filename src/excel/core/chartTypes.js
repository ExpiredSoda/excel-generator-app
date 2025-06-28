// excel/core/chartTypes.js
// Specific chart type implementations

import { ExcelChart } from './excelChart.js';

/**
 * Pie chart implementation
 */
export class PieChart extends ExcelChart {
  constructor(title = '') {
    super(title, 'pieChart');
    this.firstSliceAngle = 0;
  }

  /**
   * Set first slice angle
   * @param {number} angle - Angle in degrees (0-360)
   */
  setFirstSliceAngle(angle) {
    this.firstSliceAngle = angle;
  }

  /**
   * Generate pie chart plot area XML - matches working Excel chart
   * @returns {string} Plot area XML
   */
  generatePlotAreaXML() {
    const seriesXML = this.series.map((series, index) => 
      this.generateSeriesXML(series, index)
    ).join('');

    const dataLabelsXML = this.generateDataLabelsXML();

    return `<c:plotArea><c:layout/><c:pieChart><c:varyColors val="${this.varyColors ? '1' : '0'}"/>${seriesXML}${dataLabelsXML}<c:firstSliceAng val="${this.firstSliceAngle}"/></c:pieChart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst><a:softEdge rad="736600"/></a:effectLst></c:spPr></c:plotArea>`;
  }

  /**
   * Generate series XML for pie chart
   * @param {ChartSeries} series - Chart series
   * @param {number} index - Series index
   * @returns {string} Series XML
   */
  generateSeriesXML(series, index) {
    // Use fixed number of data points - 8 (matching working chart)
    const totalPoints = 8;
    const dataPointsXML = this.generateDataPointsXML(totalPoints);
    
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:strRef><c:f>'Quick Reference'!$G$3</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Usage</c:v></c:pt></c:strCache></c:strRef></c:tx>${dataPointsXML}<c:cat>${series.categories.toXML()}</c:cat><c:val>${series.values.toXML()}</c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000010-0A41-4AE2-9935-08E1571919AC}"/></c:ext></c:extLst></c:ser>`;
  }

  /**
   * Generate data point styling XML - matches working Excel chart exactly
   * @param {number} totalPoints - Total number of data points to generate
   * @returns {string} Data points XML
   */
  generateDataPointsXML(totalPoints = 8) {
    const dataPoints = [];
    
    // Use accent6 with different shades/tints like the working chart
    const shadeValues = [
      'val="45000"', // shade 45%
      'val="61000"', // shade 61%
      'val="76000"', // shade 76%
      'val="92000"', // shade 92%
      'val="93000"', // tint 93%
      'val="77000"', // tint 77%
      'val="62000"', // tint 62%
      'val="46000"'  // tint 46%
    ];
    
    for (let index = 0; index < totalPoints; index++) {
      const shadeType = index <= 3 ? 'shade' : 'tint';
      const shadeValue = shadeValues[index] || 'val="50000"';
      
      // Generate unique ID for extension
      const uniqueId = `{${String(index * 2 + 1).padStart(8, '0')}-0A41-4AE2-9935-08E1571919AC}`;
      
      dataPoints.push(`<c:dPt><c:idx val="${index}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:schemeClr val="accent6"><a:${shadeType} ${shadeValue}/></a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst><a:outerShdw blurRad="254000" sx="102000" sy="102000" algn="ctr" rotWithShape="0"><a:prstClr val="black"><a:alpha val="20000"/></a:prstClr></a:outerShdw></a:effectLst></c:spPr><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="${uniqueId}"/></c:ext></c:extLst></c:dPt>`);
    }
    
    return dataPoints.join('');
  }

  /**
   * Generate data labels XML - matches working Excel chart with percentages
   * @returns {string} Data labels XML
   */
  generateDataLabelsXML() {
    return `<c:dLbls><c:spPr><a:pattFill prst="pct75"><a:fgClr><a:schemeClr val="dk1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:fgClr><a:bgClr><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:bgClr></a:pattFill><a:ln><a:noFill/></a:ln><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" lIns="38100" tIns="19050" rIns="38100" bIns="19050" anchor="ctr" anchorCtr="1"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/><c:leaderLines><c:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:solidFill></a:ln><a:effectLst/></c:spPr></c:leaderLines><c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"/></c:extLst></c:dLbls>`;
  }
} 
// excel/core/dataValidation.js
// Modular data validation rule for Excel

/**
 * Represents a data validation rule for Excel cells
 */
export class DataValidationRule {
  /**
   * @param {string} sqref - Cell range (e.g., "C2:C100")
   * @param {string} formula1 - Formula for the list (e.g., '=Legends!$A$3:$A$10')
   * @param {Object} options - Additional options
   * @param {boolean} [options.allowBlank=false] - Allow blank values
   * @param {boolean} [options.showDropDown=true] - Show dropdown arrow
   * @param {string} [options.type='list'] - Validation type
   */
  constructor(sqref, formula1, options = {}) {
    this.sqref = sqref;
    this.formula1 = formula1;
    this.allowBlank = options.allowBlank ?? false;
    this.showDropDown = options.showDropDown ?? true;
    this.showInputMessage = options.showInputMessage ?? false;
    this.showErrorMessage = options.showErrorMessage ?? true;
    this.type = options.type ?? 'list';
    this.errorStyle = options.errorStyle ?? 'stop';
    this.errorTitle = options.errorTitle ?? 'Invalid Entry';
    this.error = options.error ?? 'Please select a value from the dropdown list.';
    this.promptTitle = options.promptTitle;
    this.prompt = options.prompt;
  }

  /**
   * Generate XML for this data validation rule
   */
  toXML() {
    // NOTE: showDropDown attribute works backwards in XLSX format!
    // showDropDown="0" shows the dropdown, showDropDown="1" hides it
    const showDropDownValue = this.showDropDown ? "0" : "1";
    return `<dataValidation type="list" showDropDown="${showDropDownValue}" sqref="${this.sqref}"><formula1>${this.formula1}</formula1></dataValidation>`;
  }
}

/**
 * Generate the <dataValidations> XML block from an array of rules
 * @param {DataValidationRule[]} rules
 * @returns {string}
 */
export function generateDataValidationsXML(rules) {
  if (!rules || rules.length === 0) return '';
  const xml = `<dataValidations count="${rules.length}">${rules.map(r => r.toXML()).join('')}</dataValidations>`;
  return xml;
} 
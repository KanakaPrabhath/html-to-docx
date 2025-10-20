import { convertHtmlToDocx, convertHtmlToDocxFile } from './converter.js';

class HtmlToDocx {
  constructor(options = {}) {
    this.options = {
      fontSize: options.fontSize || 11,
      fontFamily: options.fontFamily || 'Calibri',
      lineHeight: options.lineHeight || 1.15,
      ...options
    };
  }

  /**
   * Convert HTML string to DOCX buffer
   * @param {string} html - HTML content to convert
   * @param {Object} options - Additional conversion options
   * @returns {Promise<Buffer>} - DOCX file buffer
   */
  async convertHtmlToDocx(html, options = {}) {
    const mergedOptions = { ...this.options, ...options };
    return convertHtmlToDocx(html, mergedOptions);
  }

  /**
   * Convert HTML string to DOCX file
   * @param {string} html - HTML content to convert
   * @param {string} outputPath - Path where to save the DOCX file
   * @param {Object} options - Additional conversion options
   * @returns {Promise<void>}
   */
  async convertHtmlToDocxFile(html, outputPath, options = {}) {
    const mergedOptions = { ...this.options, ...options };
    return convertHtmlToDocxFile(html, outputPath, mergedOptions);
  }
}

export default HtmlToDocx;
export { HtmlToDocx, convertHtmlToDocx, convertHtmlToDocxFile };

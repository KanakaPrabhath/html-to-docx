import { escapeXml, getAlignment } from './utils.js';

/**
 * OOXML generation utilities for HTML to DOCX conversion
 */

/**
 * Create OOXML paragraph
 * @param {string} text - Paragraph text
 * @param {Object} styles - Paragraph styles
 * @param {Element} element - Original DOM element (optional)
 * @param {Function} processChildNodes - Function to process child nodes
 * @returns {string} - OOXML paragraph
 */
export function createParagraph(text, styles, element = null, processChildNodes = null) {
  let runs = '';

  if (element && processChildNodes) {
    // Process child nodes for complex formatting
    runs = processChildNodes(element, styles);
  } else {
    // Simple text
    runs = createRun(text, styles);
  }

  const pPr = createParagraphProperties(styles);

  return `<w:p>${pPr}${runs}</w:p>`;
}

/**
 * Create OOXML run (text with formatting)
 * @param {string} text - Run text
 * @param {Object} styles - Run styles
 * @returns {string} - OOXML run
 */
export function createRun(text, styles) {
  if (!text) return '';

  const rPr = createRunProperties(styles);
  const escapedText = escapeXml(text);

  return `<w:r>${rPr}<w:t xml:space="preserve">${escapedText}</w:t></w:r>`;
}

/**
 * Create paragraph properties OOXML
 * @param {Object} styles - Paragraph styles
 * @returns {string} - OOXML paragraph properties
 */
export function createParagraphProperties(styles) {
  let pPr = '';

  if (styles.alignment) {
    const alignment = getAlignment(styles.alignment);
    pPr += `<w:jc w:val="${alignment}"/>`;
  }

  if (styles.heading) {
    pPr += `<w:pStyle w:val="Heading${styles.heading}"/>`;
  }

  if (pPr) {
    return `<w:pPr>${pPr}</w:pPr>`;
  }

  return '';
}

/**
 * Create run properties OOXML
 * @param {Object} styles - Run styles
 * @returns {string} - OOXML run properties
 */
export function createRunProperties(styles) {
  let rPr = '';

  if (styles.bold) {
    rPr += '<w:b/>';
  }

  if (styles.italic) {
    rPr += '<w:i/>';
  }

  if (styles.underline) {
    rPr += '<w:u w:val="single"/>';
  }

  if (styles.color) {
    rPr += `<w:color w:val="${styles.color}"/>`;
  }

  if (styles.backgroundColor) {
    rPr += `<w:shd w:val="clear" w:color="auto" w:fill="${styles.backgroundColor}"/>`;
  }

  if (styles.fontSize) {
    rPr += `<w:sz w:val="${styles.fontSize}"/>`;
    rPr += `<w:szCs w:val="${styles.fontSize}"/>`;
  }

  if (styles.fontFamily) {
    rPr += `<w:rFonts w:ascii="${styles.fontFamily}" w:hAnsi="${styles.fontFamily}"/>`;
  }

  if (rPr) {
    return `<w:rPr>${rPr}</w:rPr>`;
  }

  return '';
}
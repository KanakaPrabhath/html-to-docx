import { JSDOM } from 'jsdom';
import { getTextContent, cleanHtml, parseDimension } from './utils.js';
import { parseStyles } from './styleParser.js';
import { createParagraph, createRun } from './ooxmlGenerator.js';
import { createTextBox } from './vmlGenerator.js';
import { createTable, createTableRow, createTableCell } from './ooxmlGenerator.js';
import { MediaManager, createImageOOXML } from './mediaHandler.js';

/**
 * HTML parsing utilities for HTML to DOCX conversion
 */

const INLINE_STYLE_MAP = {
  'strong': 'bold',
  'b': 'bold',
  'em': 'italic',
  'i': 'italic',
  'u': 'underline',
  'strike': 'strikethrough',
  's': 'strikethrough'
};

/**
 * Check if element should be wrapped in a text box based on styles
 * @param {Object} styles - Parsed styles
 * @returns {boolean} - Whether to wrap in text box
 */
function shouldWrapInTextBox(styles) {
  return styles.backgroundColor && (styles.padding || styles.width || styles.borderRadius);
}

/**
 * Process all child nodes of an element
 * @param {Element} element - Parent element
 * @param {MediaManager} mediaManager - Media manager instance
 * @param {Object} options - Conversion options
 * @param {string} part - The part this content belongs to
 * @returns {Promise<string>} - OOXML content
 */
async function processChildNodesContent(element, mediaManager, options, part) {
  let content = '';
  for (const child of element.childNodes) {
    content += await processNode(child, mediaManager, options, part);
  }
  return content;
}

/**
 * Merge global options into styles, skipping lineHeight if noSpacing is set
 * @param {Object} styles - Parsed styles
 * @param {Object} options - Global options
 */
function mergeGlobalOptions(styles, options) {
  if (!styles.noSpacing) {
    if (options.fontSize && !styles.fontSize) styles.fontSize = options.fontSize * 2; // Convert to half-points
    if (options.fontFamily && !styles.fontFamily) styles.fontFamily = options.fontFamily;
    if (options.lineHeight && !styles.lineHeight) styles.lineHeight = options.lineHeight;
  }
}

/**
 * Wrap content with page breaks if specified in styles
 * @param {string} content - OOXML content
 * @param {Object} styles - Parsed styles
 * @returns {string} - Content wrapped with page breaks
 */
function wrapWithPageBreaks(content, styles) {
  let result = content;

  if (styles.pageBreakBefore) {
    result = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + result;
  }

  if (styles.pageBreakAfter) {
    result = result + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  }

  return result;
}

/**
 * Parse image-specific styles from element
 * @param {Element} element - Image element
 * @returns {Object} - Parsed image styles
 */
function parseImageStyles(element) {
  const styles = {};
  const styleAttr = element.getAttribute('style');

  if (styleAttr) {
    const styleRules = styleAttr.split(';').filter(s => s.trim());

    for (const rule of styleRules) {
      const [property, value] = rule.split(':').map(s => s.trim());

      switch (property) {
        case 'width':
          styles.width = parseDimension(value);
          break;
        case 'height':
          styles.height = parseDimension(value);
          break;
        case 'float':
          styles.float = value;
          break;
        case 'margin':
          styles.margin = parseMargin(value);
          break;
        case 'display':
          styles.display = value;
          break;
      }
    }
  }

  return styles;
}

/**
 * Parse margin value
 * @param {string} value - Margin value (e.g., '0px 0px 10px 10px')
 * @returns {Object} - Margin object with top, right, bottom, left
 */
function parseMargin(value) {
  if (!value) return null;
  const parts = value.split(' ').filter(p => p.trim());
  const margins = {};

  if (parts.length === 1) {
    // Single value: margin: 10px
    const val = parseDimension(parts[0]);
    margins.top = margins.right = margins.bottom = margins.left = val;
  } else if (parts.length === 2) {
    // Two values: margin: 10px 20px
    margins.top = margins.bottom = parseDimension(parts[0]);
    margins.right = margins.left = parseDimension(parts[1]);
  } else if (parts.length === 3) {
    // Three values: margin: 10px 20px 30px
    margins.top = parseDimension(parts[0]);
    margins.right = margins.left = parseDimension(parts[1]);
    margins.bottom = parseDimension(parts[2]);
  } else if (parts.length === 4) {
    // Four values: margin: 10px 20px 30px 40px
    margins.top = parseDimension(parts[0]);
    margins.right = parseDimension(parts[1]);
    margins.bottom = parseDimension(parts[2]);
    margins.left = parseDimension(parts[3]);
  }

  return margins;
}

/**
 * Parse HTML and convert to OOXML body content
 * @param {string} html - HTML content
 * @param {MediaManager} mediaManager - Media manager instance
 * @param {Object} options - Conversion options
 * @param {string} part - The part this content belongs to ('document', 'header', 'footer')
 * @returns {Promise<string>} - OOXML body content
 */
export async function parseHtml(html, mediaManager = null, options = {}, part = 'document') {
  // Clean the HTML first
  const cleanedHtml = cleanHtml(html);

  const dom = new JSDOM(cleanedHtml);
  const document = dom.window.document;
  const body = document.body;

  let bodyContent = '';

  // Process all child nodes
  for (const node of body.childNodes) {
    bodyContent += await processNode(node, mediaManager, options, part);
  }

  // If no content was generated, add a default paragraph
  if (!bodyContent.trim()) {
    bodyContent = await createParagraph(' ', {});
  }

  return bodyContent;
}

/**
 * Process a DOM node and convert to OOXML
 * @param {Node} node - DOM node to process
 * @param {MediaManager} mediaManager - Media manager instance
 * @param {string} part - The part this content belongs to
 * @returns {Promise<string>} - OOXML content
 */
export async function processNode(node, mediaManager = null, options = {}, part = 'document') {
  if (node.nodeType === 3) { // Text node
    const text = node.textContent.trim();
    if (text) {
      return await createParagraph(text, {});
    }
    return '';
  }

  if (node.nodeType === 1) { // Element node
    const tagName = node.tagName.toLowerCase();

    switch (tagName) {
      case 'p':
        return await processParagraph(node, mediaManager, options, part);
      case 'h1':
      case 'h2':
      case 'h3':
      case 'h4':
      case 'h5':
      case 'h6':
        return await processHeading(node, mediaManager, options, part);
      case 'strong':
      case 'b':
        return await processInlineStyle(node, { bold: true }, mediaManager, options, part);
      case 'em':
      case 'i':
        return await processInlineStyle(node, { italic: true }, mediaManager, options, part);
      case 'u':
        return await processInlineStyle(node, { underline: true }, mediaManager, options, part);
      case 'strike':
      case 's':
        return await processInlineStyle(node, { strikethrough: true }, mediaManager, options, part);
      case 'br':
        return await createParagraph('', {});
      case 'span':
        return processSpan(node, mediaManager, part);
      case 'div':
        return await processDiv(node, mediaManager, options, part);
      case 'ul':
      case 'ol':
        return await processList(node, mediaManager, options, part);
      case 'li':
        // Li should be processed within lists, but if standalone, treat as ul
        return await processListItem(node, 'ul', mediaManager, options, part);
      case 'img':
        const imageOOXML = await processImage(node, mediaManager, part);
        return imageOOXML ? `<w:p>${imageOOXML}</w:p>` : '';
      case 'table':
        return await processTable(node, mediaManager, options, part);
      case 'tr':
        return await processTableRow(node, mediaManager, options, part);
      case 'td':
      case 'th':
        return await processTableCell(node, mediaManager, options, part);
      case 'page-break':
        return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      default:
        // Process children for unknown tags
        return await processChildNodesContent(node, mediaManager, options, part);
    }
  }

  return '';
}

/**
 * Process paragraph element
 * @param {Element} element - Paragraph element
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML paragraph
 */
export async function processParagraph(element, mediaManager = null, options = {}, part = 'document') {
  let styles = parseStyles(element);
  
  // Merge global options, but skip lineHeight if noSpacing is set
  mergeGlobalOptions(styles, options);
  
  const text = getTextContent(element);

  const paragraphContent = await createParagraph(text, styles, element, async (el, parentStyles) => await processChildNodes(el, parentStyles, mediaManager, options, part));

  return wrapWithPageBreaks(paragraphContent, styles);
}

/**
 * Process heading element
 * @param {Element} element - Heading element
 * @param {MediaManager} mediaManager - Media manager instance
 * @param {Object} options - Conversion options
 * @param {string} part - The part this content belongs to
 * @returns {Promise<string>} - OOXML paragraph with heading style
 */
export async function processHeading(element, mediaManager = null, options = {}, part = 'document') {
  const level = parseInt(element.tagName.substring(1));
  let styles = parseStyles(element);
  const text = getTextContent(element);

  // Check for heading replacements
  if (options.headingReplacements && options.headingReplacements[level - 1]) {
    let replacementHtml = options.headingReplacements[level - 1];
    replacementHtml = replacementHtml.replace(/HEADING_TEXT/g, text);
    return await parseHtml(replacementHtml, mediaManager, options, part);
  }

  styles.heading = level;

  // Merge global options, but skip lineHeight if noSpacing is set
  mergeGlobalOptions(styles, options);

  // Check if this heading should be wrapped in a text box
  // Wrap in text box if it has background color and other box-like properties
  if (shouldWrapInTextBox(styles)) {
    // Don't use heading style inside textbox to avoid spacing issues
    // Instead, just keep the text formatting (bold, size, etc.)
    delete styles.heading;
    styles.insideTextBox = true;
    
    // Create paragraph content without page breaks for text box
    const paragraphContent = await createParagraph(text, styles, element, async (el, parentStyles) => await processChildNodes(el, parentStyles, mediaManager, options, part));
    
    // Wrap in text box
    const textBoxContent = createTextBox(paragraphContent, styles);
    
    // Apply page breaks to the text box if needed
    return wrapWithPageBreaks(textBoxContent, styles);
  } else {
    // Normal heading processing
    const headingContent = await createParagraph(text, styles, element, async (el, parentStyles) => await processChildNodes(el, parentStyles, mediaManager, options, part));
    return wrapWithPageBreaks(headingContent, styles);
  }
}

/**
 * Process inline styled element
 * @param {Element} element - Inline element
 * @param {Object} styleOverrides - Style overrides
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML content
 */
export async function processInlineStyle(element, styleOverrides, mediaManager = null, options = {}, part = 'document') {
  const styles = { ...parseStyles(element), ...styleOverrides };
  const text = getTextContent(element);
  return await createParagraph(text, styles, element, async (el, parentStyles) => await processChildNodes(el, parentStyles, mediaManager, options, part));
}

/**
 * Process span element
 * @param {Element} element - Span element
 * @param {MediaManager} mediaManager - Media manager instance
 * @param {string} part - The part this content belongs to
 * @returns {string} - OOXML content
 */
export function processSpan(element, mediaManager = null, part = 'document') {
  const styles = parseStyles(element);
  const text = getTextContent(element);
  return createRun(text, styles);
}

/**
 * Process div element
 * @param {Element} element - Div element
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML content
 */
export async function processDiv(element, mediaManager = null, options = {}, part = 'document') {
  const styles = parseStyles(element);
  
  // Check if this div should be wrapped in a text box
  // Wrap in text box if it has background color and other box-like properties (padding, border, width, border-radius)
  if (shouldWrapInTextBox(styles)) {
    // Process all child nodes and collect content
    const content = await processChildNodesContent(element, mediaManager, options, part);
    
    // Wrap in text box if we have content
    if (content.trim()) {
      return createTextBox(content, styles);
    }
  }
  
  // Normal div processing - just process children
  return await processChildNodesContent(element, mediaManager, options, part);
}

/**
 * Process list element
 * @param {Element} element - List element (ul or ol)
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML content
 */
export async function processList(element, mediaManager = null, options = {}, part = 'document') {
  const listType = element.tagName.toLowerCase();
  let content = '';
  
  // Process each list item
  for (const child of element.children) {
    if (child.tagName.toLowerCase() === 'li') {
      content += await processListItem(child, listType, mediaManager, options, part);
    }
  }
  
  return content;
}

/**
 * Process list item element
 * @param {Element} element - List item element
 * @param {string} listType - Type of list ('ul' or 'ol')
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML paragraph with numbering
 */
export async function processListItem(element, listType, mediaManager = null, options = {}, part = 'document') {
  const styles = parseStyles(element);
  styles.listType = listType;
  styles.listItem = true;
  const text = getTextContent(element);

  const listItemContent = await createParagraph(text, styles, element, async (el, parentStyles) => await processChildNodes(el, parentStyles, mediaManager, options, part));

  return wrapWithPageBreaks(listItemContent, styles);
}

/**
 * Process image element
 * @param {Element} element - Image element
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {string} - OOXML image content
 */
export async function processImage(element, mediaManager = null, part = 'document') {
  if (!mediaManager) {
    return ''; // Skip images if no media manager
  }

  const src = element.getAttribute('src');
  const alt = element.getAttribute('alt') || '';
  const width = parseInt(element.getAttribute('width')) || null;
  const height = parseInt(element.getAttribute('height')) || null;

  // Parse style attribute for additional image properties
  const imageStyles = parseImageStyles(element);

  if (!src) {
    return '';
  }

  try {
    const mediaFile = await mediaManager.addImage(src, alt, part);
    if (mediaFile) {
      const options = {
        width: imageStyles.width || width,
        height: imageStyles.height || height,
        float: imageStyles.float,
        margin: imageStyles.margin,
        display: imageStyles.display
      };
      return createImageOOXML(mediaFile, options);
    }
  } catch (error) {
    console.warn('Failed to process image:', error.message);
  }

  return '';
}

/**
 * Process child nodes for runs
 * @param {Element} element - Parent element
 * @param {Object} parentStyles - Parent styles
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML runs
 */
export async function processChildNodes(element, parentStyles, mediaManager = null, options = {}, part = 'document') {
  let runs = '';

  for (const node of element.childNodes) {
    if (node.nodeType === 3) {
      // Text node
      const text = node.textContent;
      if (text.trim() || text === ' ') {
        runs += createRun(text, parentStyles);
      }
    } else if (node.nodeType === 1) {
      // Element node
      const tagName = node.tagName.toLowerCase();
      const childStyles = { ...parentStyles, ...parseStyles(node) };

      const styleProperty = INLINE_STYLE_MAP[tagName];
      if (styleProperty) {
        childStyles[styleProperty] = true;
        runs += await processChildNodes(node, childStyles, mediaManager, options, part);
      } else {
        switch (tagName) {
          case 'span':
            runs += await processChildNodes(node, childStyles, mediaManager, options, part);
            break;
          case 'br':
            runs += '<w:r><w:br/></w:r>';
            break;
          case 'img':
            runs += await processImage(node, mediaManager, part);
            break;
          case 'page-break':
            runs += '<w:r><w:br w:type="page"/></w:r>';
            break;
          default:
            runs += await processChildNodes(node, childStyles, mediaManager, options, part);
        }
      }
    }
  }

  return runs;
}

/**
 * Process table element
 * @param {Element} element - Table element
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML table
 */
export async function processTable(element, mediaManager = null, options = {}, part = 'document') {
  const styles = parseStyles(element);
  let tableContent = '';

  // Process table rows
  const rows = element.querySelectorAll('tr');
  for (const row of rows) {
    tableContent += await processTableRow(row, mediaManager, options, part);
  }

  return createTable(tableContent, styles);
}

/**
 * Process table row element
 * @param {Element} element - Table row element
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML table row
 */
export async function processTableRow(element, mediaManager = null, options = {}, part = 'document') {
  const styles = parseStyles(element);
  let rowContent = '';

  // Process table cells
  const cells = element.querySelectorAll('td, th');
  for (const cell of cells) {
    rowContent += await processTableCell(cell, mediaManager, options, part);
  }

  return createTableRow(rowContent, styles);
}

/**
 * Process table cell element
 * @param {Element} element - Table cell element
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML table cell
 */
export async function processTableCell(element, mediaManager = null, options = {}, part = 'document') {
  const styles = parseStyles(element);
  const isHeader = element.tagName.toLowerCase() === 'th';
  const text = getTextContent(element);
  
  const cellContent = await createParagraph(text, styles, element, async (el, parentStyles) => await processChildNodes(el, parentStyles, mediaManager, options, part));

  return createTableCell(cellContent, styles, isHeader);
}
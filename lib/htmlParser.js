import { JSDOM } from 'jsdom';
import { getTextContent } from './utils.js';
import { parseStyles } from './styleParser.js';
import { createParagraph, createRun } from './ooxmlGenerator.js';

/**
 * HTML parsing utilities for HTML to DOCX conversion
 */

/**
 * Parse HTML and convert to OOXML body content
 * @param {string} html - HTML content
 * @returns {string} - OOXML body content
 */
export function parseHtml(html) {
  const dom = new JSDOM(html);
  const document = dom.window.document;
  const body = document.body;

  let bodyContent = '';

  // Process all child nodes
  for (const node of body.childNodes) {
    bodyContent += processNode(node);
  }

  return bodyContent;
}

/**
 * Process a DOM node and convert to OOXML
 * @param {Node} node - DOM node to process
 * @returns {string} - OOXML content
 */
export function processNode(node) {
  if (node.nodeType === 3) { // Text node
    const text = node.textContent.trim();
    if (text) {
      return createParagraph(text, {});
    }
    return '';
  }

  if (node.nodeType === 1) { // Element node
    const tagName = node.tagName.toLowerCase();

    switch (tagName) {
      case 'p':
        return processParagraph(node);
      case 'h1':
      case 'h2':
      case 'h3':
      case 'h4':
      case 'h5':
      case 'h6':
        return processHeading(node);
      case 'strong':
      case 'b':
        return processInlineStyle(node, { bold: true });
      case 'em':
      case 'i':
        return processInlineStyle(node, { italic: true });
      case 'u':
        return processInlineStyle(node, { underline: true });
      case 'br':
        return createParagraph('', {});
      case 'span':
        return processSpan(node);
      case 'div':
        return processDiv(node);
      case 'ul':
      case 'ol':
        return processList(node);
      case 'li':
        return processListItem(node);
      default:
        // Process children for unknown tags
        let content = '';
        for (const child of node.childNodes) {
          content += processNode(child);
        }
        return content;
    }
  }

  return '';
}

/**
 * Process paragraph element
 * @param {Element} element - Paragraph element
 * @returns {string} - OOXML paragraph
 */
export function processParagraph(element) {
  const styles = parseStyles(element);
  const text = getTextContent(element);

  // Handle empty paragraphs (like <br>)
  if (!text.trim() && element.innerHTML === '<br>') {
    return createParagraph('', styles);
  }

  return createParagraph(text, styles, element, processChildNodes);
}

/**
 * Process heading element
 * @param {Element} element - Heading element
 * @returns {string} - OOXML paragraph with heading style
 */
export function processHeading(element) {
  const level = parseInt(element.tagName.substring(1));
  const styles = parseStyles(element);
  styles.heading = level;
  const text = getTextContent(element);

  return createParagraph(text, styles, element, processChildNodes);
}

/**
 * Process inline styled element
 * @param {Element} element - Inline element
 * @param {Object} styleOverrides - Style overrides
 * @returns {string} - OOXML content
 */
export function processInlineStyle(element, styleOverrides) {
  const styles = { ...parseStyles(element), ...styleOverrides };
  const text = getTextContent(element);
  return createParagraph(text, styles, element, processChildNodes);
}

/**
 * Process span element
 * @param {Element} element - Span element
 * @returns {string} - OOXML content
 */
export function processSpan(element) {
  const styles = parseStyles(element);
  const text = getTextContent(element);
  return createRun(text, styles);
}

/**
 * Process div element
 * @param {Element} element - Div element
 * @returns {string} - OOXML content
 */
export function processDiv(element) {
  let content = '';
  for (const child of element.childNodes) {
    content += processNode(child);
  }
  return content;
}

/**
 * Process list element
 * @param {Element} element - List element (ul or ol)
 * @returns {string} - OOXML content
 */
export function processList(element) {
  let content = '';
  for (const child of element.children) {
    if (child.tagName.toLowerCase() === 'li') {
      content += processListItem(child);
    }
  }
  return content;
}

/**
 * Process list item element
 * @param {Element} element - List item element
 * @returns {string} - OOXML paragraph with numbering
 */
export function processListItem(element) {
  const styles = parseStyles(element);
  styles.listItem = true;
  const text = getTextContent(element);
  return createParagraph(text, styles, element, processChildNodes);
}

/**
 * Process child nodes for runs
 * @param {Element} element - Parent element
 * @param {Object} parentStyles - Parent styles
 * @returns {string} - OOXML runs
 */
export function processChildNodes(element, parentStyles) {
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

      switch (tagName) {
        case 'strong':
        case 'b':
          childStyles.bold = true;
          runs += processChildNodes(node, childStyles);
          break;
        case 'em':
        case 'i':
          childStyles.italic = true;
          runs += processChildNodes(node, childStyles);
          break;
        case 'u':
          childStyles.underline = true;
          runs += processChildNodes(node, childStyles);
          break;
        case 'span':
          runs += processChildNodes(node, childStyles);
          break;
        case 'br':
          runs += '<w:r><w:br/></w:r>';
          break;
        default:
          runs += processChildNodes(node, childStyles);
      }
    }
  }

  return runs;
}
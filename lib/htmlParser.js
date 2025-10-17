import { JSDOM } from 'jsdom';
import { getTextContent, cleanHtml } from './utils.js';
import { parseStyles } from './styleParser.js';
import { createParagraph, createRun } from './ooxmlGenerator.js';
import { MediaManager, createImageOOXML } from './mediaHandler.js';

/**
 * HTML parsing utilities for HTML to DOCX conversion
 */

/**
 * Parse HTML and convert to OOXML body content
 * @param {string} html - HTML content
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML body content
 */
export async function parseHtml(html, mediaManager = null) {
  // Clean the HTML first
  const cleanedHtml = cleanHtml(html);

  const dom = new JSDOM(cleanedHtml);
  const document = dom.window.document;
  const body = document.body;

  let bodyContent = '';

  // Process all child nodes
  for (const node of body.childNodes) {
    bodyContent += await processNode(node, mediaManager);
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
 * @returns {Promise<string>} - OOXML content
 */
export async function processNode(node, mediaManager = null) {
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
        return await processParagraph(node, mediaManager);
      case 'h1':
      case 'h2':
      case 'h3':
      case 'h4':
      case 'h5':
      case 'h6':
        return await processHeading(node, mediaManager);
      case 'strong':
      case 'b':
        return await processInlineStyle(node, { bold: true }, mediaManager);
      case 'em':
      case 'i':
        return await processInlineStyle(node, { italic: true }, mediaManager);
      case 'u':
        return await processInlineStyle(node, { underline: true }, mediaManager);
      case 'strike':
      case 's':
        return await processInlineStyle(node, { strikethrough: true }, mediaManager);
      case 'br':
        return await createParagraph('', {});
      case 'span':
        return await processSpan(node, mediaManager);
      case 'div':
        return await processDiv(node, mediaManager);
      case 'ul':
      case 'ol':
        return await processList(node, mediaManager);
      case 'li':
        return await processListItem(node, mediaManager);
      case 'img':
        const imageOOXML = await processImage(node, mediaManager);
        return imageOOXML ? `<w:p>${imageOOXML}</w:p>` : '';
      case 'page-break':
        return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      default:
        // Process children for unknown tags
        let content = '';
        for (const child of node.childNodes) {
          content += await processNode(child, mediaManager);
        }
        return content;
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
export async function processParagraph(element, mediaManager = null) {
  const styles = parseStyles(element);
  const text = getTextContent(element);

  return await createParagraph(text, styles, element, async (el, parentStyles) => await processChildNodes(el, parentStyles, mediaManager));
}

/**
 * Process heading element
 * @param {Element} element - Heading element
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML paragraph with heading style
 */
export async function processHeading(element, mediaManager = null) {
  const level = parseInt(element.tagName.substring(1));
  const styles = parseStyles(element);
  styles.heading = level;
  const text = getTextContent(element);

  return await createParagraph(text, styles, element, async (el, parentStyles) => await processChildNodes(el, parentStyles, mediaManager));
}

/**
 * Process inline styled element
 * @param {Element} element - Inline element
 * @param {Object} styleOverrides - Style overrides
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML content
 */
export async function processInlineStyle(element, styleOverrides, mediaManager = null) {
  const styles = { ...parseStyles(element), ...styleOverrides };
  const text = getTextContent(element);
  return await createParagraph(text, styles, element, async (el, parentStyles) => await processChildNodes(el, parentStyles, mediaManager));
}

/**
 * Process span element
 * @param {Element} element - Span element
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML content
 */
export async function processSpan(element, mediaManager = null) {
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
export async function processDiv(element, mediaManager = null) {
  let content = '';
  for (const child of element.childNodes) {
    content += await processNode(child, mediaManager);
  }
  return content;
}

/**
 * Process list element
 * @param {Element} element - List element (ul or ol)
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {Promise<string>} - OOXML content
 */
export async function processList(element, mediaManager = null) {
  const listType = element.tagName.toLowerCase();
  let content = '';
  
  // Process each list item
  for (const child of element.children) {
    if (child.tagName.toLowerCase() === 'li') {
      content += await processListItem(child, listType, mediaManager);
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
export async function processListItem(element, listType, mediaManager = null) {
  const styles = parseStyles(element);
  styles.listType = listType;
  styles.listItem = true;
  const text = getTextContent(element);
  return await createParagraph(text, styles, element, async (el, parentStyles) => await processChildNodes(el, parentStyles, mediaManager));
}

/**
 * Process image element
 * @param {Element} element - Image element
 * @param {MediaManager} mediaManager - Media manager instance
 * @returns {string} - OOXML image content
 */
export async function processImage(element, mediaManager = null) {
  if (!mediaManager) {
    return ''; // Skip images if no media manager
  }

  const src = element.getAttribute('src');
  const alt = element.getAttribute('alt') || '';
  const width = parseInt(element.getAttribute('width')) || null;
  const height = parseInt(element.getAttribute('height')) || null;

  if (!src) {
    return '';
  }

  try {
    const mediaFile = await mediaManager.addImage(src, alt);
    if (mediaFile) {
      return createImageOOXML(mediaFile, { width, height });
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
export async function processChildNodes(element, parentStyles, mediaManager = null) {
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
          runs += await processChildNodes(node, childStyles, mediaManager);
          break;
        case 'em':
        case 'i':
          childStyles.italic = true;
          runs += await processChildNodes(node, childStyles, mediaManager);
          break;
        case 'u':
          childStyles.underline = true;
          runs += await processChildNodes(node, childStyles, mediaManager);
          break;
        case 'strike':
        case 's':
          childStyles.strikethrough = true;
          runs += await processChildNodes(node, childStyles, mediaManager);
          break;
        case 'span':
          runs += await processChildNodes(node, childStyles, mediaManager);
          break;
        case 'br':
          runs += '<w:r><w:br/></w:r>';
          break;
        case 'img':
          runs += await processImage(node, mediaManager);
          break;
        default:
          runs += await processChildNodes(node, childStyles, mediaManager);
      }
    }
  }

  return runs;
}
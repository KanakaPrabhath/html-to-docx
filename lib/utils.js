/**
 * Utility functions for HTML to DOCX conversion
 */

/**
 * Escape XML special characters
 * @param {string} text - Text to escape
 * @returns {string} - Escaped text
 */
export function escapeXml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Get OOXML alignment value
 * @param {string} alignment - CSS alignment value
 * @returns {string} - OOXML alignment value
 */
export function getAlignment(alignment) {
  const alignmentMap = {
    'left': 'left',
    'center': 'center',
    'right': 'right',
    'justify': 'both'
  };
  return alignmentMap[alignment] || 'left';
}

/**
 * Get text content from element, handling inline styles
 * @param {Element} element - DOM element
 * @returns {string} - Text content
 */
export function getTextContent(element) {
  let text = '';
  for (const node of element.childNodes) {
    if (node.nodeType === 3) {
      text += node.textContent;
    } else if (node.nodeType === 1) {
      text += getTextContent(node);
    }
  }
  return text;
}
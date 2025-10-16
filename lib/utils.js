/**
 * Utility functions for HTML to DOCX conversion
 */

/**
 * Clean and sanitize HTML content before processing
 * @param {string} html - Raw HTML content
 * @returns {string} - Cleaned HTML content
 */
export function cleanHtml(html) {
  if (!html || typeof html !== 'string') {
    return '<p></p>';
  }

  let cleaned = html;

  // Remove invalid Unicode characters (control characters except \t, \n, \r)
  cleaned = cleaned.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]/g, '');

  // Remove script and style tags
  cleaned = cleaned.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
  cleaned = cleaned.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

  // Remove potentially dangerous tags
  cleaned = cleaned.replace(/<\/?(iframe|object|embed|form|input|button|select|textarea)[^>]*>/gi, '');

  // Simplify complex CSS in style attributes
  cleaned = cleaned.replace(/style="[^"]*"/gi, (match) => {
    const styles = match.slice(7, -1).split(';').filter(s => s.trim());
    const simpleStyles = styles.filter(style => {
      const [property] = style.split(':').map(s => s.trim());
      // Keep only simple, safe properties
      return ['text-align', 'font-weight', 'font-style', 'text-decoration', 'color', 'background-color', 'font-size', 'font-family'].includes(property);
    });
    return simpleStyles.length ? `style="${simpleStyles.join(';')}"` : '';
  });

  // Fix common malformed HTML
  // Remove unclosed tags (basic fix)
  cleaned = cleaned.replace(/<[^>]*$/g, ''); // Remove incomplete tags at end

  // Ensure we have a body
  if (!cleaned.includes('<body') && !cleaned.includes('<html')) {
    cleaned = `<html><body>${cleaned}</body></html>`;
  }

  return cleaned;
}

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
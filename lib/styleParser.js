/**
 * Style parsing utilities for HTML to DOCX conversion
 */

/**
 * Parse inline styles from element
 * @param {Element} element - DOM element
 * @returns {Object} - Parsed styles
 */
export function parseStyles(element) {
  const styles = {};
  const styleAttr = element.getAttribute('style');

  if (styleAttr) {
    const styleRules = styleAttr.split(';').filter(s => s.trim());

    for (const rule of styleRules) {
      const [property, value] = rule.split(':').map(s => s.trim());

      switch (property) {
        case 'text-align':
          styles.alignment = value;
          break;
        case 'font-weight':
          if (value === 'bold' || parseInt(value) >= 600) {
            styles.bold = true;
          }
          break;
        case 'font-style':
          if (value === 'italic') {
            styles.italic = true;
          }
          break;
        case 'text-decoration':
          if (value.includes('underline')) {
            styles.underline = true;
          }
          if (value.includes('line-through')) {
            styles.strikethrough = true;
          }
          break;
        case 'color':
          styles.color = parseColor(value);
          break;
        case 'background-color':
          styles.backgroundColor = parseColor(value);
          break;
        case 'font-size':
          styles.fontSize = parseFontSize(value);
          break;
        case 'font-family':
          styles.fontFamily = value.replace(/['"]/g, '').split(',')[0].trim();
          break;
        case 'margin-left':
          styles.indent = parseIndent(value);
          break;
        case 'border-collapse':
          styles.borderCollapse = value;
          break;
        case 'width':
          styles.width = value;
          break;
        case 'border':
          styles.border = parseBorder(value);
          break;
        case 'padding':
          styles.padding = value;
          break;
        case 'page-break-before':
          if (value === 'always') {
            styles.pageBreakBefore = true;
          }
          break;
        case 'page-break-after':
          if (value === 'always') {
            styles.pageBreakAfter = true;
          }
          break;
        case 'page-break-inside':
          if (value === 'avoid') {
            styles.pageBreakInside = 'avoid';
          }
          break;
      }
    }
  }

  // Handle data-indent-level attribute
  const indentLevel = element.getAttribute('data-indent-level');
  if (indentLevel) {
    const level = parseInt(indentLevel);
    styles.indent = level * 480; // 480 twentieths of a point = 0.5 inch per level
  }

  return styles;
}

/**
 * Parse color value to hex
 * @param {string} color - Color value (hex, rgb, named)
 * @returns {string} - Hex color (without #)
 */
export function parseColor(color) {
  if (color.startsWith('#')) {
    return color.substring(1).toUpperCase();
  }
  // For simplicity, return black for other formats
  // You can add RGB parsing here if needed
  return '000000';
}

/**
 * Parse font size to half-points
 * @param {string} size - Font size (e.g., '12px', '14pt')
 * @returns {number} - Font size in half-points
 */
export function parseFontSize(size) {
  const value = parseFloat(size);
  if (size.includes('px')) {
    return Math.round(value * 1.5); // Convert px to half-points
  } else if (size.includes('pt')) {
    return value * 2; // Convert pt to half-points
  }
  return value * 2;
}

/**
 * Parse indent value to twentieths of a point
 * @param {string} value - Indent value (e.g., '32px', '1in')
 * @returns {number} - Indent in twentieths of a point
 */
export function parseIndent(value) {
  const numValue = parseFloat(value);
  if (value.includes('px')) {
    return Math.round(numValue * 15); // Convert px to twentieths of a point
  } else if (value.includes('in')) {
    return Math.round(numValue * 1440); // Convert inches to twentieths of a point
  } else if (value.includes('cm')) {
    return Math.round(numValue * 567); // Convert cm to twentieths of a point
  } else if (value.includes('mm')) {
    return Math.round(numValue * 56.7); // Convert mm to twentieths of a point
  } else if (value.includes('pt')) {
    return Math.round(numValue * 20); // Convert pt to twentieths of a point
  }
  return Math.round(numValue * 20); // Default assume pt
}

/**
 * Parse border value
 * @param {string} border - Border value (e.g., '1px solid #000000')
 * @returns {Object} - Parsed border properties
 */
export function parseBorder(border) {
  const parts = border.split(' ').filter(p => p.trim());
  const borderObj = {};

  if (parts.length >= 3) {
    borderObj.width = parts[0];
    borderObj.style = parts[1];
    borderObj.color = parseColor(parts[2]);
  }

  return borderObj;
}
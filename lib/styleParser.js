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
      }
    }
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
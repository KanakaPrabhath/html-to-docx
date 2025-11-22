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
        case 'background':
          // Parse background property (supports linear-gradient)
          const gradientInfo = parseGradient(value);
          if (gradientInfo) {
            styles.gradient = gradientInfo;
          } else {
            // Fallback to solid color
            styles.backgroundColor = parseColor(value);
          }
          break;
        case 'font-size':
          styles.fontSize = parseFontSize(value);
          break;
        case 'font-family':
          styles.fontFamily = value.replace(/['"]/g, '').split(',')[0].trim();
          break;
        case 'margin':
          // Handle shorthand margin property (e.g., "0px" or "10px 20px" or "10px 20px 30px 40px")
          parseMarginShorthand(value, styles);
          break;
        case 'margin-left':
          styles.indent = parseIndent(value);
          break;
        case 'margin-top':
          styles.spacingBefore = parseIndent(value);
          break;
        case 'margin-bottom':
          styles.spacingAfter = parseIndent(value);
          break;
        case 'margin-right':
          // For now, ignore margin-right as it's not directly supported in OOXML paragraphs
          break;
        case 'border-collapse':
          styles.borderCollapse = value;
          break;
        case 'width':
          styles.width = value;
          break;
        case 'height':
          styles.height = value;
          break;
        case 'border':
          styles.border = parseBorder(value);
          break;
        case 'padding':
          styles.padding = value;
          break;
        case 'border-radius':
          styles.borderRadius = value;
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
        case 'line-height':
          styles.lineHeight = parseFloat(value);
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

  // Handle no-spacing attribute
  if (element.hasAttribute('data-no-spacing')) {
    styles.noSpacing = true;
  }

  return styles;
}

/**
 * Parse color value to hex
 * @param {string} color - Color value (hex, rgb, named)
 * @returns {string} - Hex color (RRGGBB without #)
 */
export function parseColor(color) {
  if (!color) return '000000';
  
  color = color.trim();
  
  if (color.startsWith('#')) {
    const hex = color.substring(1).toLowerCase();
    
    // Handle different hex formats
    if (hex.length === 3) {
      // #RGB -> #RRGGBB
      return (hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2]).toUpperCase();
    } else if (hex.length === 4) {
      // #RGBA -> #RRGGBB (ignore alpha)
      return (hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2]).toUpperCase();
    } else if (hex.length === 6) {
      // #RRGGBB
      return hex.toUpperCase();
    } else if (hex.length === 8) {
      // #RRGGBBAA -> #RRGGBB (ignore alpha)
      return hex.substring(0, 6).toUpperCase();
    } else {
      // Invalid, return black
      return '000000';
    }
  }
  
  // Handle named colors (basic ones)
  const namedColors = {
    'black': '000000',
    'white': 'FFFFFF',
    'red': 'FF0000',
    'green': '00FF00',
    'blue': '0000FF',
    'yellow': 'FFFF00',
    'cyan': '00FFFF',
    'magenta': 'FF00FF',
    'gray': '808080',
    'grey': '808080'
  };
  
  if (namedColors[color.toLowerCase()]) {
    return namedColors[color.toLowerCase()];
  }
  
  // For rgb/rgba, return black for simplicity
  // Could add parsing later
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

/**
 * Parse shorthand margin value and apply to styles
 * @param {string} margin - Margin value (e.g., '0px', '10px 20px', '10px 20px 30px 40px')
 * @param {Object} styles - Styles object to update
 */
function parseMarginShorthand(margin, styles) {
  const parts = margin.split(' ').filter(p => p.trim());
  
  if (parts.length === 1) {
    // margin: 10px - all sides
    const value = parseIndent(parts[0]);
    styles.spacingBefore = value;
    styles.spacingAfter = value;
    styles.indent = value;
  } else if (parts.length === 2) {
    // margin: 10px 20px - vertical horizontal
    const vertical = parseIndent(parts[0]);
    const horizontal = parseIndent(parts[1]);
    styles.spacingBefore = vertical;
    styles.spacingAfter = vertical;
    styles.indent = horizontal;
  } else if (parts.length === 3) {
    // margin: 10px 20px 30px - top horizontal bottom
    styles.spacingBefore = parseIndent(parts[0]);
    styles.indent = parseIndent(parts[1]);
    styles.spacingAfter = parseIndent(parts[2]);
  } else if (parts.length === 4) {
    // margin: 10px 20px 30px 40px - top right bottom left
    styles.spacingBefore = parseIndent(parts[0]);
    styles.spacingAfter = parseIndent(parts[2]);
    styles.indent = parseIndent(parts[3]);
  }
}

/**
 * Parse gradient value from background property
 * @param {string} value - Background value (e.g., 'linear-gradient(to right, #000000ff, #ffffffff)')
 * @returns {Object|null} - Gradient info object or null if not a gradient
 */
export function parseGradient(value) {
  if (!value || !value.includes('linear-gradient')) {
    return null;
  }

  // Extract gradient function content
  const gradientMatch = value.match(/linear-gradient\(([^)]+)\)/);
  if (!gradientMatch) return null;

  const gradientContent = gradientMatch[1];
  const parts = gradientContent.split(',').map(p => p.trim());

  if (parts.length < 2) return null;

  // Parse direction (to right, to bottom, etc.)
  let direction = 'to bottom'; // default
  let colorIndex = 0;
  
  if (parts[0].startsWith('to ')) {
    direction = parts[0];
    colorIndex = 1;
  } else if (parts[0].includes('deg')) {
    direction = parts[0];
    colorIndex = 1;
  }

  // Parse colors
  const colors = [];
  for (let i = colorIndex; i < parts.length; i++) {
    const colorPart = parts[i].trim();
    // Extract just the color value (without position if present)
    const colorValue = colorPart.split(' ')[0];
    colors.push(parseColor(colorValue));
  }

  if (colors.length < 2) return null;

  return {
    type: 'linear',
    direction: direction,
    colors: colors
  };
}
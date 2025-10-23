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
  cleaned = cleaned.replace(/<[^>]+style="[^"]*"[^>]*>/gi, (match) => {
    const tagMatch = match.match(/^<(\w+)/);
    const tagName = tagMatch ? tagMatch[1].toLowerCase() : '';
    const styleMatch = match.match(/style="([^"]*)"/);
    const styleContent = styleMatch ? styleMatch[1] : '';

    if (!styleContent) return match;

    const styles = styleContent.split(';').filter(s => s.trim());
    let allowedProperties = [];

    if (tagName === 'img') {
      // Allow image-specific properties
      allowedProperties = ['width', 'height', 'float', 'margin', 'display'];
    } else {
      // Keep only simple, safe properties for other elements
      allowedProperties = ['text-align', 'font-weight', 'font-style', 'text-decoration', 'color', 'background-color', 'font-size', 'font-family', 'border-collapse', 'width', 'height', 'border', 'padding', 'border-radius', 'vertical-align'];
    }

    const simpleStyles = styles.filter(style => {
      const [property] = style.split(':').map(s => s.trim());
      return allowedProperties.includes(property);
    });

    if (simpleStyles.length) {
      return match.replace(/style="[^"]*"/, `style="${simpleStyles.join(';')}"`);
    } else {
      return match.replace(/style="[^"]*"/, '');
    }
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

/**
 * Parse dimension value to pixels
 * @param {string} value - Dimension value (e.g., '181px', '50%')
 * @returns {number|null} - Dimension in pixels or null
 */
export function parseDimension(value) {
  if (!value) return null;
  const numValue = parseFloat(value);
  if (isNaN(numValue)) return null;

  if (value.includes('px')) {
    return numValue;
  } else if (value.includes('%')) {
    // For percentage, we'll need context, but for now return as is
    // This might need adjustment based on container width
    return numValue;
  } else if (value.includes('em') || value.includes('rem')) {
    // Convert em/rem to px (assuming 16px base)
    return numValue * 16;
  } else if (value.includes('pt')) {
    // Convert pt to px
    return numValue * 1.333;
  }

  // Default assume px
  return numValue;
}

/**
 * Parse padding value
 * @param {string} padding - Padding value
 * @returns {Object} - Padding object
 */
export function parsePadding(padding) {
  if (!padding) return null;

  const parts = padding.split(' ').filter(p => p.trim());
  const paddings = {};

  if (parts.length === 1) {
    const val = parseFloat(parts[0]) * 20; // Convert to twentieths of a point (supports decimals)
    paddings.top = paddings.right = paddings.bottom = paddings.left = val;
  } else if (parts.length === 2) {
    paddings.top = paddings.bottom = parseFloat(parts[0]) * 20;
    paddings.right = paddings.left = parseFloat(parts[1]) * 20;
  } else if (parts.length === 3) {
    paddings.top = parseFloat(parts[0]) * 20;
    paddings.right = paddings.left = parseFloat(parts[1]) * 20;
    paddings.bottom = parseFloat(parts[2]) * 20;
  } else if (parts.length === 4) {
    paddings.top = parseFloat(parts[0]) * 20;
    paddings.right = parseFloat(parts[1]) * 20;
    paddings.bottom = parseFloat(parts[2]) * 20;
    paddings.left = parseFloat(parts[3]) * 20;
  }

  return paddings;
}

/**
 * Sanitize and validate conversion options
 * @param {Object} options - Raw options object
 * @returns {Object} - Sanitized options object
 */
export function sanitizeOptions(options = {}) {
  const sanitized = { ...options };

  // Validate and sanitize fontSize
  if (typeof sanitized.fontSize !== 'number' || isNaN(sanitized.fontSize) || sanitized.fontSize <= 0) {
    sanitized.fontSize = 11;
  }

  // Validate and sanitize fontFamily
  if (typeof sanitized.fontFamily !== 'string' || sanitized.fontFamily.trim() === '') {
    sanitized.fontFamily = 'Calibri';
  }

  // Validate and sanitize lineHeight
  if (typeof sanitized.lineHeight !== 'number' || isNaN(sanitized.lineHeight) || sanitized.lineHeight <= 0) {
    sanitized.lineHeight = 1.15;
  }

  // Validate and sanitize pageSize
  const validPageSizes = ['A4', 'LETTER', 'LEGAL'];
  if (typeof sanitized.pageSize === 'string') {
    sanitized.pageSize = sanitized.pageSize.toUpperCase();
    if (!validPageSizes.includes(sanitized.pageSize)) {
      // Check if it's a custom object
      if (typeof options.pageSize === 'object' && options.pageSize.width && options.pageSize.height) {
        sanitized.pageSize = options.pageSize;
      } else {
        sanitized.pageSize = 'A4';
      }
    }
  } else if (typeof sanitized.pageSize === 'object' && sanitized.pageSize.width && sanitized.pageSize.height) {
    // Custom size object
    if (typeof sanitized.pageSize.width !== 'number' || typeof sanitized.pageSize.height !== 'number' ||
        isNaN(sanitized.pageSize.width) || isNaN(sanitized.pageSize.height) ||
        sanitized.pageSize.width <= 0 || sanitized.pageSize.height <= 0) {
      sanitized.pageSize = 'A4';
    }
  } else {
    sanitized.pageSize = 'A4';
  }

  // Validate and sanitize margins
  const marginProps = ['marginTop', 'marginRight', 'marginBottom', 'marginLeft', 'marginHeader', 'marginFooter', 'marginGutter'];
  marginProps.forEach(prop => {
    if (sanitized[prop] !== undefined) {
      const val = parseFloat(sanitized[prop]);
      if (isNaN(val) || val < 0) {
        delete sanitized[prop]; // Let it use defaults
      } else {
        sanitized[prop] = val;
      }
    }
  });

  // Validate and sanitize header/footer heights
  ['headerHeight', 'footerHeight'].forEach(prop => {
    if (sanitized[prop] !== undefined) {
      const val = parseFloat(sanitized[prop]);
      if (isNaN(val) || val < 0) {
        delete sanitized[prop];
      } else {
        sanitized[prop] = val;
      }
    }
  });

  return sanitized;
}
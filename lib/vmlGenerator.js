import { parseDimension, parsePadding } from './utils.js';

/**
 * VML generation utilities for HTML to DOCX conversion
 * Reference: https://www.w3.org/TR/NOTE-VML
 */

// Constants for EMU conversions (English Metric Units: 914400 EMUs = 1 inch)
const EMU_PER_INCH = 914400;
const EMU_PER_POINT = EMU_PER_INCH / 72;
const DEFAULT_TEXTBOX_WIDTH_EMU = 9144000; // ~6.35 inches
const DEFAULT_TEXTBOX_HEIGHT_EMU = 914400; // ~0.64 inches
const DEFAULT_PAGE_WIDTH_A4_TWIPS = 11906;
const DEFAULT_PAGE_HEIGHT_A4_TWIPS = 16838;
const TWIPS_PER_INCH = 1440;
const TWIPS_PER_POINT = 20;

/**
 * Create VML text box with content
 * @param {string} content - Text box content (paragraph OOXML)
 * @param {Object} styles - Text box styles
 * @returns {string} - VML text box wrapped in paragraph
 */
export function createTextBox(content, styles) {
  // Parse dimensions
  const width = styles.width ? parseTextBoxDimension(styles.width) : DEFAULT_TEXTBOX_WIDTH_EMU;
  const height = styles.height ? parseTextBoxDimension(styles.height) : DEFAULT_TEXTBOX_HEIGHT_EMU;

  // Convert EMUs to points for style attribute (914400 EMUs = 72 points)
  const widthPt = Math.round((width / EMU_PER_INCH) * 72);
  const heightPt = Math.round((height / EMU_PER_INCH) * 72);

  // Parse padding (convert to twentieths of a point)
  const parsedPadding = parsePadding(styles.padding);
  const paddingObj = parsedPadding || { top: 100, right: 100, bottom: 100, left: 100 };

  // Background color or gradient
  const hasGradient = styles.gradient && styles.gradient.colors && styles.gradient.colors.length >= 2;
  const bgColor = hasGradient ? styles.gradient.colors[0] : (styles.backgroundColor || 'FFFFFF');
  
  // Gradient fill element
  let fillElement = '';
  if (hasGradient) {
    fillElement = generateGradientFill(styles.gradient);
  }

  // Border settings
  const { strokeAttr, strokeElement } = generateBorderAttributes(styles.border);

  // Border radius (rounded corners)
  let arcsize = '';
  if (styles.borderRadius) {
    const radiusPx = parseFloat(styles.borderRadius);
    const minDim = Math.min(widthPt, heightPt);
    const maxRadius = minDim / 2;
    const effectiveRadius = Math.min(radiusPx, maxRadius);
    const arcsizeValue = effectiveRadius / maxRadius;
    if (arcsizeValue > 0) {
      arcsize = ` arcsize="${arcsizeValue}"`;
    }
  }

  // Generate unique ID for the shape
  const shapeId = Math.floor(Math.random() * 1000000) + 1000;

  // Convert padding from twentieths of a point to points
  const paddingLeftPt = paddingObj.left / 20;
  const paddingTopPt = paddingObj.top / 20;
  const paddingRightPt = paddingObj.right / 20;
  const paddingBottomPt = paddingObj.bottom / 20;

  // Build paragraph properties for spacing
  const pPr = generateParagraphProperties(styles);

  // Build fill attribute
  const fillAttr = hasGradient ? 'filled="t"' : `fillcolor="#${bgColor}"`;
  
  // Build the text box VML
  const textBoxXml = `<w:p>
${pPr}
  <w:r>
    <w:pict>
      <v:roundrect id="_x0000_s${shapeId}" style="width:${widthPt}pt;height:${heightPt}pt;" ${fillAttr} ${strokeAttr}${arcsize}>
        ${fillElement}
        ${strokeElement}
        <v:textbox style="mso-fit-shape-to-text:t" inset="${paddingLeftPt}pt,${paddingTopPt}pt,${paddingRightPt}pt,${paddingBottomPt}pt">
          <w:txbxContent>
${content}
          </w:txbxContent>
        </v:textbox>
      </v:roundrect>
    </w:pict>
  </w:r>
</w:p>`;

  return textBoxXml;
}

/**
 * Generate stroke element for border styles
 * @param {string} color - Border color
 * @param {number} size - Border size in points
 * @param {string} style - Border style
 * @returns {string} - Stroke element XML
 */
function generateStrokeElement(color, size, style) {
  let strokeWeightPt = size;
  let strokeElement = `<v:stroke color="#${color}" weight="${strokeWeightPt}pt" insetpen="t"`;

  // Add dashstyle for dotted/dashed lines
  if (style === 'dotted') {
    strokeElement += ` dashstyle="dot"`;
  } else if (style === 'dashed') {
    strokeElement += ` dashstyle="dash"`;
  }

  // Add linestyle for double lines
  if (style === 'double') {
    strokeElement += ` linestyle="thickThin"`;
  } else if (style === 'thick') {
    // For thick, we can use a heavier weight
    strokeWeightPt = Math.max(size * 1.5, size + 2);
    strokeElement = `<v:stroke color="#${color}" weight="${strokeWeightPt}pt" insetpen="t"`;
  }

  strokeElement += '/>';
  return strokeElement;
}

/**
 * Generate VML rounded border shape for page borders
 * @param {Object} options - Shape options
 * @returns {string} - VML shape XML
 */
export function generateRoundedBorderShape(options = {}) {
  const borderConfig = options.pageBorder || {};
  const radius = borderConfig.radius || 0;
  const color = borderConfig.color || '000000';
  const size = borderConfig.size || 1;
  const style = borderConfig.style || 'single';

  // Calculate page size
  const { width: pageWidth, height: pageHeight } = getPageDimensions(options.pageSize);
  // Margins
  const marginTop = options.marginTop !== undefined ? Math.round(options.marginTop * 1440) : 1440;
  const marginRight = options.marginRight !== undefined ? Math.round(options.marginRight * 1440) : 1440;
  const marginBottom = options.marginBottom !== undefined ? Math.round(options.marginBottom * 1440) : 1440;
  const marginLeft = options.marginLeft !== undefined ? Math.round(options.marginLeft * 1440) : 1440;
  const headerHeight = options.headerHeight !== undefined ? Math.round(options.headerHeight * 1440) : 720;
  const footerHeight = options.footerHeight !== undefined ? Math.round(options.footerHeight * 1440) : 720;

  // Convert to points (1 pt = 20 twips)
  const widthPt = (pageWidth - marginLeft - marginRight) / 20;
  const heightPt = (pageHeight - headerHeight - footerHeight) / 20;
  const marginLeftPt = marginLeft / 20;
  const marginTopPt = marginTop / 20;
  const headerHeightPt = headerHeight / 20;
  const topPt = headerHeightPt;

  // Optional expansion beyond margins (default to half of marginLeft)
  let marginOffsetInches = (options.marginLeft !== undefined ? options.marginLeft : 1) / 3;
  if (borderConfig.marginOffset !== undefined) {
    const parsedOffset = parseFloat(borderConfig.marginOffset);
    if (!Number.isNaN(parsedOffset) && parsedOffset >= 0) {
      marginOffsetInches = parsedOffset;
    }
  }
  const offsetTwips = Math.round(marginOffsetInches * 1440);
  const offsetPt = offsetTwips / 20;
  const adjustedLeftPt = Math.max(marginLeftPt - offsetPt, 0);
  const adjustedTopPt = Math.max(topPt - 0, 0);
  const adjustedWidthPt = widthPt + offsetPt * 2;
  const adjustedHeightPt = heightPt + 0 * 2;

  // Border settings - size comes in points (pt)
  const borderWeightPt = size; // Size is already in points
  const strokeAttr = 'stroked="t"';

  // Build stroke element with style support
  const strokeElement = generateStrokeElement(color, borderWeightPt, style);
  
  // Convert radius from points to normalized arcsize (0-1) for VML
  // arcsize represents corner rounding: 0 = square, 1 = maximum rounding
  const normalizedRadius = Math.min(Math.max(radius / Math.min(adjustedWidthPt, adjustedHeightPt), 0), 1);
  const arcsize = ` arcsize="${normalizedRadius}"`;

  // Generate unique ID
  const shapeId = Math.floor(Math.random() * 1000000) + 1000;

  // Adjust position/size so the outer stroke aligns with the page margins
  const shapeXml = `
  <w:p>
    <w:r>
      <w:pict>
        <v:roundrect id="_x0000_s${shapeId}" style="position:absolute;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;left:${adjustedLeftPt}pt;top:${adjustedTopPt}pt;width:${adjustedWidthPt}pt;height:${adjustedHeightPt}pt;" fillcolor="transparent" ${strokeAttr}${arcsize}>
          ${strokeElement}
        </v:roundrect>
      </w:pict>
    </w:r>
  </w:p>`;

  return shapeXml;
}

/**
 * Generate VML gradient fill element
 * @param {Object} gradient - Gradient configuration
 * @returns {string} - VML fill element
 */
function generateGradientFill(gradient) {
  if (!gradient || !gradient.colors || gradient.colors.length < 2) {
    return '';
  }

  const color1 = gradient.colors[0];
  const color2 = gradient.colors[gradient.colors.length - 1];
  
  // Determine angle based on direction
  // VML uses angle differently: 0=left-to-right, 90=top-to-bottom, 180=right-to-left, 270=bottom-to-top
  let angle = -90; // Default: top to bottom (VML angle needs -90 for top-to-bottom)
  if (gradient.direction) {
    const dir = gradient.direction.toLowerCase();
    if (dir.includes('to right')) {
      angle = 0;
    } else if (dir.includes('to left')) {
      angle = 180;
    } else if (dir.includes('to bottom')) {
      angle = -90;
    } else if (dir.includes('to top')) {
      angle = 90;
    } else if (dir.includes('deg')) {
      // Extract angle value and convert CSS angle to VML angle
      const angleMatch = dir.match(/([\d.]+)deg/);
      if (angleMatch) {
        // CSS angles: 0deg=to top, 90deg=to right, 180deg=to bottom, 270deg=to left
        // VML angles: 0=left-to-right, 90=bottom-to-top, 180=right-to-left, 270=top-to-bottom
        const cssAngle = parseFloat(angleMatch[1]);
        angle = cssAngle - 90;
      }
    }
  }

  // VML fill element with gradient
  // Note: In VML, color is the starting color and color2 is the ending color
  // Using method="linear" for linear gradient
  return `<v:fill type="gradient" method="linear" color="#${color1}" color2="#${color2}" angle="${angle}"/>`;
}

/**
 * Generate stroke attributes and elements for borders
 * @param {Object} border - Border configuration
 * @returns {Object} - { strokeAttr, strokeElement }
 */
function generateBorderAttributes(border) {
  if (!border) {
    return { strokeAttr: 'stroked="f"', strokeElement: '' };
  }

  const borderColor = border.color || '000000';
  const borderWidth = parseFloat(border.width) || 1;
  const borderWeightPt = borderWidth * 0.75; // Convert px to pt (approximate)

  return {
    strokeAttr: 'stroked="t"',
    strokeElement: `<v:stroke color="#${borderColor}" weight="${borderWeightPt}pt"/>`
  };
}

/**
 * Generate paragraph properties XML for spacing
 * @param {Object} styles - Styles object
 * @returns {string} - Paragraph properties XML
 */
function generateParagraphProperties(styles) {
  if (!styles.spacingBefore && !styles.spacingAfter) return '';

  let spacingAttrs = '';
  if (styles.spacingBefore) spacingAttrs += ` w:before="${styles.spacingBefore}"`;
  if (styles.spacingAfter) spacingAttrs += ` w:after="${styles.spacingAfter}"`;

  return `<w:pPr><w:spacing${spacingAttrs}/></w:pPr>`;
}

/**
 * Get page dimensions in twips based on page size
 * @param {string|Object} pageSize - Page size specification
 * @returns {Object} - { width, height } in twips
 */
function getPageDimensions(pageSize) {
  const defaultSize = { width: DEFAULT_PAGE_WIDTH_A4_TWIPS, height: DEFAULT_PAGE_HEIGHT_A4_TWIPS };

  if (typeof pageSize === 'object' && pageSize.width && pageSize.height) {
    return {
      width: Math.round(pageSize.width * TWIPS_PER_INCH),
      height: Math.round(pageSize.height * TWIPS_PER_INCH)
    };
  }

  const sizes = {
    A4: { width: DEFAULT_PAGE_WIDTH_A4_TWIPS, height: DEFAULT_PAGE_HEIGHT_A4_TWIPS },
    LETTER: { width: 12240, height: 15840 },
    LEGAL: { width: 12240, height: 20160 }
  };

  return sizes[pageSize?.toUpperCase()] || defaultSize;
}

/**
 * Parse dimension for text box (convert to EMUs - English Metric Units)
 * @param {string} value - Dimension value (e.g., '100px', '50%', '2in')
 * @returns {number} - Dimension in EMUs (914400 EMUs = 1 inch)
 */
function parseTextBoxDimension(value) {
  if (!value) return EMU_PER_INCH; // Default 1 inch

  const numValue = parseFloat(value);
  if (isNaN(numValue)) return EMU_PER_INCH;

  // Unit conversion factors to EMUs
  const unitFactors = {
    px: EMU_PER_INCH / 96, // Assuming 96 DPI
    '%': 5943600 / 100, // Relative to page width (~6.5 inches)
    in: EMU_PER_INCH,
    cm: EMU_PER_INCH / 2.54,
    mm: EMU_PER_INCH / 25.4,
    pt: EMU_PER_POINT
  };

  for (const [unit, factor] of Object.entries(unitFactors)) {
    if (value.includes(unit)) {
      return Math.round(numValue * factor);
    }
  }

  // Default assume px
  return Math.round(numValue * (EMU_PER_INCH / 96));
}
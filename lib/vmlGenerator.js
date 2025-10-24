import { parseDimension, parsePadding } from './utils.js';

/**
 * VML generation utilities for HTML to DOCX conversion
 * Reference: https://www.w3.org/TR/NOTE-VML
 */

/**
 * Create VML text box with content
 * @param {string} content - Text box content (paragraph OOXML)
 * @param {Object} styles - Text box styles
 * @returns {string} - VML text box wrapped in paragraph
 */
export function createTextBox(content, styles) {
  // Parse dimensions
  const width = styles.width ? parseTextBoxDimension(styles.width) : 9144000; // Default ~6.35 inches in EMUs
  const height = styles.height ? parseTextBoxDimension(styles.height) : 914400; // Default ~0.64 inches in EMUs

  // Convert EMUs to points for style attribute (914400 EMUs = 72 points)
  const widthPt = Math.round((width / 914400) * 72);
  const heightPt = Math.round((height / 914400) * 72);

  // Parse padding (convert to twentieths of a point)
  const parsedPadding = parsePadding(styles.padding);
  const paddingObj = parsedPadding || { top: 100, right: 100, bottom: 100, left: 100 };

  // Background color
  const bgColor = styles.backgroundColor || 'FFFFFF';

  // Border settings
  let strokeAttr = 'stroked="f"'; // Default: no border
  let strokeElement = '';

  if (styles.border) {
    const borderColor = styles.border.color || '000000';
    const borderWidth = styles.border.width ? parseFloat(styles.border.width) : 1;
    const borderWeightPt = borderWidth * 0.75; // Convert px to pt (approximate)

    strokeAttr = 'stroked="t"';
    strokeElement = `<v:stroke color="#${borderColor}" weight="${borderWeightPt}pt"/>`;
  }

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
  let pPr = '';
  if (styles.spacingBefore || styles.spacingAfter) {
    let spacingAttrs = '';
    if (styles.spacingBefore) spacingAttrs += ` w:before="${styles.spacingBefore}"`;
    if (styles.spacingAfter) spacingAttrs += ` w:after="${styles.spacingAfter}"`;
    pPr = `<w:pPr><w:spacing${spacingAttrs}/></w:pPr>`;
  }

  // Build the text box VML
  const textBoxXml = `<w:p>
${pPr}
  <w:r>
    <w:pict>
      <v:roundrect id="_x0000_s${shapeId}" style="width:${widthPt}pt;height:${heightPt}pt;" fillcolor="#${bgColor}" ${strokeAttr}${arcsize}>
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
 * Parse dimension for text box (convert to EMUs - English Metric Units)
 * @param {string} value - Dimension value (e.g., '100px', '50%', '2in')
 * @returns {number} - Dimension in EMUs (914400 EMUs = 1 inch)
 */
function parseTextBoxDimension(value) {
  if (!value) return 914400; // Default 1 inch

  const numValue = parseFloat(value);

  if (value.includes('px')) {
    // Convert px to EMUs (assuming 96 DPI: 1 inch = 96px = 914400 EMUs)
    return Math.round((numValue / 96) * 914400);
  } else if (value.includes('%')) {
    // For percentage, assume relative to page width (~6.5 inches for A4)
    const pageWidthEMU = 5943600; // ~6.5 inches
    return Math.round((numValue / 100) * pageWidthEMU);
  } else if (value.includes('in')) {
    // Convert inches to EMUs
    return Math.round(numValue * 914400);
  } else if (value.includes('cm')) {
    // Convert cm to EMUs (1 inch = 2.54 cm)
    return Math.round((numValue / 2.54) * 914400);
  } else if (value.includes('mm')) {
    // Convert mm to EMUs
    return Math.round((numValue / 25.4) * 914400);
  } else if (value.includes('pt')) {
    // Convert pt to EMUs (72 pt = 1 inch)
    return Math.round((numValue / 72) * 914400);
  }

  // Default assume px
  return Math.round((numValue / 96) * 914400);
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
  const pageSize = options.pageSize || 'A4';
  let pageWidth, pageHeight;
  switch (pageSize.toUpperCase()) {
    case 'A4':
      pageWidth = 11906;
      pageHeight = 16838;
      break;
    case 'LETTER':
      pageWidth = 12240;
      pageHeight = 15840;
      break;
    case 'LEGAL':
      pageWidth = 12240;
      pageHeight = 20160;
      break;
    default:
      if (typeof pageSize === 'object' && pageSize.width && pageSize.height) {
        pageWidth = Math.round(pageSize.width * 1440);
        pageHeight = Math.round(pageSize.height * 1440);
      } else {
        pageWidth = 11906;
        pageHeight = 16838;
      }
  }
  // Margins
  const marginTop = options.marginTop !== undefined ? Math.round(options.marginTop * 1440) : 1440;
  const marginRight = options.marginRight !== undefined ? Math.round(options.marginRight * 1440) : 1440;
  const marginBottom = options.marginBottom !== undefined ? Math.round(options.marginBottom * 1440) : 1440;
  const marginLeft = options.marginLeft !== undefined ? Math.round(options.marginLeft * 1440) : 1440;
  const marginHeader = options.marginHeader !== undefined ? Math.round(options.marginHeader * 1440) : 720;

  // Convert to points (1 pt = 20 twips)
  const widthPt = (pageWidth - marginLeft - marginRight) / 20;
  const heightPt = (pageHeight - marginTop - marginBottom) / 20;
  const marginLeftPt = marginLeft / 20;
  const marginTopPt = marginTop / 20;
  const marginHeaderPt = marginHeader / 20;
  const topPt = Math.max(marginTopPt - marginHeaderPt, 0);

  // Optional expansion beyond margins (default to half of marginLeft)
  let marginOffsetInches = (options.marginLeft !== undefined ? options.marginLeft : 1) / 2;
  if (borderConfig.marginOffset !== undefined) {
    const parsedOffset = parseFloat(borderConfig.marginOffset);
    if (!Number.isNaN(parsedOffset) && parsedOffset >= 0) {
      marginOffsetInches = parsedOffset;
    }
  }
  const offsetTwips = Math.round(marginOffsetInches * 1440);
  const offsetPt = offsetTwips / 20;
  const adjustedLeftPt = Math.max(marginLeftPt - offsetPt, 0);
  const adjustedTopPt = topPt;
  const adjustedWidthPt = widthPt + offsetPt * 2;
  const adjustedHeightPt = heightPt + offsetPt * 2;

  // Border settings - size comes in points (pt)
  const borderWeightPt = size; // Size is already in points
  const strokeAttr = 'stroked="t"';
  
  // Build stroke element with style support
  let strokeWeightPt = borderWeightPt;
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
    strokeWeightPt = Math.max(borderWeightPt * 1.5, borderWeightPt + 2);
    strokeElement = `<v:stroke color="#${color}" weight="${strokeWeightPt}pt" insetpen="t"`;
  }
  
  strokeElement += '/>';
  
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
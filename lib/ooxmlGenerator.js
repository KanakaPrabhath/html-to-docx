import { escapeXml, getAlignment, parseDimension } from './utils.js';

/**
 * OOXML generation utilities for HTML to DOCX conversion
 */

/**
 * Create OOXML paragraph
 * @param {string} text - Paragraph text
 * @param {Object} styles - Paragraph styles
 * @param {Element} element - Original DOM element (optional)
 * @param {Function} processChildNodes - Function to process child nodes
 * @returns {Promise<string>} - OOXML paragraph
 */
export async function createParagraph(text, styles, element = null, processChildNodes = null) {
  let runs = '';

  if (element && processChildNodes) {
    // Process child nodes for complex formatting
    runs = await processChildNodes(element, styles);
  } else {
    // Simple text
    runs = createRun(text, styles);
  }

  const pPr = createParagraphProperties(styles);

  if (runs || pPr) {
    return `<w:p>${pPr}${runs}</w:p>`;
  }

  return '';
}

/**
 * Create OOXML run (text with formatting)
 * @param {string} text - Run text
 * @param {Object} styles - Run styles
 * @returns {string} - OOXML run
 */
export function createRun(text, styles) {
  if (!text) return '';

  const rPr = createRunProperties(styles);
  const escapedText = escapeXml(text);

  return `<w:r>${rPr}<w:t xml:space="preserve">${escapedText}</w:t></w:r>`;
}

/**
 * Create paragraph properties OOXML
 * @param {Object} styles - Paragraph styles
 * @returns {string} - OOXML paragraph properties
 */
export function createParagraphProperties(styles) {
  let pPr = '';

  if (styles.alignment) {
    const alignment = getAlignment(styles.alignment);
    pPr += `<w:jc w:val="${alignment}"/>`;
  }

  if (styles.heading) {
    pPr += `<w:pStyle w:val="Heading${styles.heading}"/>`;
  }
  
  // Handle spacing - for content inside textbox, override with explicit values
  if (styles.insideTextBox) {
    // Force spacing to 0 for textbox content, and use exact line height
    const beforeVal = styles.spacingBefore !== undefined ? styles.spacingBefore : 0;
    const afterVal = styles.spacingAfter !== undefined ? styles.spacingAfter : 0;
    pPr += `<w:spacing w:before="${beforeVal}" w:after="${afterVal}" w:line="240" w:lineRule="exact"/>`;
  } else if (styles.noSpacing) {
    // No spacing for paragraphs with no-spacing class
    pPr += '<w:spacing w:before="0" w:after="0"/>';
  } else if (styles.spacingBefore !== undefined || styles.spacingAfter !== undefined) {
    // Normal spacing for non-textbox content
    let spacingAttrs = '';
    if (styles.spacingBefore !== undefined) spacingAttrs += ` w:before="${styles.spacingBefore}"`;
    if (styles.spacingAfter !== undefined) spacingAttrs += ` w:after="${styles.spacingAfter}"`;
    pPr += `<w:spacing${spacingAttrs}/>`;
  }

  if (styles.listItem) {
    const numId = styles.listType === 'ol' ? '2' : '1';
    pPr += `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`;
  }

  if (styles.indent) {
    pPr += `<w:ind w:left="${styles.indent}"/>`;
  }

  if (pPr) {
    return `<w:pPr>${pPr}</w:pPr>`;
  }

  return '';
}

/**
 * Create run properties OOXML
 * @param {Object} styles - Run styles
 * @returns {string} - OOXML run properties
 */
export function createRunProperties(styles) {
  let rPr = '';

  if (styles.bold) {
    rPr += '<w:b/>';
  }

  if (styles.italic) {
    rPr += '<w:i/>';
  }

  if (styles.underline) {
    rPr += '<w:u w:val="single"/>';
  }

  if (styles.strikethrough) {
    rPr += '<w:strike/>';
  }

  if (styles.color) {
    rPr += `<w:color w:val="${styles.color}"/>`;
  }

  if (styles.backgroundColor) {
    rPr += `<w:shd w:val="clear" w:color="auto" w:fill="${styles.backgroundColor}"/>`;
  }

  if (styles.fontSize) {
    rPr += `<w:sz w:val="${styles.fontSize}"/>`;
    rPr += `<w:szCs w:val="${styles.fontSize}"/>`;
  }

  if (styles.fontFamily) {
    rPr += `<w:rFonts w:ascii="${styles.fontFamily}" w:hAnsi="${styles.fontFamily}"/>`;
  }

  if (rPr) {
    return `<w:rPr>${rPr}</w:rPr>`;
  }

  return '';
}

/**
 * Create OOXML table
 * @param {string} content - Table content (rows)
 * @param {Object} styles - Table styles
 * @returns {string} - OOXML table
 */
export function createTable(content, styles) {
  const tblPr = createTableProperties(styles);
  const tblGrid = createTableGrid(styles);

  return `<w:tbl>${tblPr}${tblGrid}${content}</w:tbl>`;
}

/**
 * Create OOXML table row
 * @param {string} content - Row content (cells)
 * @param {Object} styles - Row styles
 * @returns {string} - OOXML table row
 */
export function createTableRow(content, styles) {
  const trPr = createTableRowProperties(styles);

  return `<w:tr>${trPr}${content}</w:tr>`;
}

/**
 * Create OOXML table cell
 * @param {string} content - Cell content
 * @param {Object} styles - Cell styles
 * @param {boolean} isHeader - Whether this is a header cell
 * @returns {string} - OOXML table cell
 */
export function createTableCell(content, styles, isHeader = false) {
  const tcPr = createTableCellProperties(styles, isHeader);

  return `<w:tc>${tcPr}${content}</w:tc>`;
}

/**
 * Create table properties OOXML
 * @param {Object} styles - Table styles
 * @returns {string} - OOXML table properties
 */
export function createTableProperties(styles) {
  let tblPr = '';

  // Table width
  if (styles.width) {
    const width = parseTableWidth(styles.width);
    tblPr += `<w:tblW w:w="${width}" w:type="pct"/>`;
  }

  // Border collapse and borders
  if (styles.borderCollapse === 'collapse' || styles.border) {
    tblPr += '<w:tblBorders>';
    const borderColor = styles.border && styles.border.color ? styles.border.color : 'auto';
    const borderSize = styles.border && styles.border.width ? parseBorderWidth(styles.border.width) : 4;
    
    tblPr += `<w:top w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
    tblPr += `<w:left w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
    tblPr += `<w:bottom w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
    tblPr += `<w:right w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
    tblPr += `<w:insideH w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
    tblPr += `<w:insideV w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
    tblPr += '</w:tblBorders>';
  }

  if (tblPr) {
    return `<w:tblPr>${tblPr}</w:tblPr>`;
  }

  return '';
}

/**
 * Create table grid OOXML
 * @param {Object} styles - Table styles
 * @returns {string} - OOXML table grid
 */
export function createTableGrid(styles) {
  // For simplicity, create a basic grid. In a full implementation,
  // this should be based on the actual column widths
  return '<w:tblGrid><w:gridCol w:w="4788"/><w:gridCol w:w="4788"/></w:tblGrid>';
}

/**
 * Create table row properties OOXML
 * @param {Object} styles - Row styles
 * @returns {string} - OOXML table row properties
 */
export function createTableRowProperties(styles) {
  let trPr = '';

  // Row height
  if (styles.height) {
    const height = parseDimension(styles.height);
    trPr += `<w:trHeight w:val="${height}" w:hRule="atLeast"/>`;
  }

  if (trPr) {
    return `<w:trPr>${trPr}</w:trPr>`;
  }

  return '';
}

/**
 * Create table cell properties OOXML
 * @param {Object} styles - Cell styles
 * @param {boolean} isHeader - Whether this is a header cell
 * @returns {string} - OOXML table cell properties
 */
export function createTableCellProperties(styles, isHeader = false) {
  let tcPr = '';

  // Cell width
  if (styles.width) {
    const width = parseTableWidth(styles.width);
    tcPr += `<w:tcW w:w="${width}" w:type="pct"/>`;
  }

  // Background color
  if (styles.backgroundColor) {
    tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${styles.backgroundColor}"/>`;
  }

  // Borders
  if (styles.border) {
    tcPr += '<w:tcBorders>';
    const borderColor = styles.border.color || 'auto';
    const borderSize = parseBorderWidth(styles.border.width) || 4;
    
    tcPr += `<w:top w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
    tcPr += `<w:left w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
    tcPr += `<w:bottom w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
    tcPr += `<w:right w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
    tcPr += '</w:tcBorders>';
  }

  // Padding
  if (styles.padding) {
    const padding = parsePadding(styles.padding);
    tcPr += `<w:tcMar><w:top w:w="${padding.top || 0}" w:type="dxa"/><w:left w:w="${padding.left || 0}" w:type="dxa"/><w:bottom w:w="${padding.bottom || 0}" w:type="dxa"/><w:right w:w="${padding.right || 0}" w:type="dxa"/></w:tcMar>`;
  }

  // Vertical alignment
  if (styles.verticalAlign) {
    const vAlign = getVerticalAlignment(styles.verticalAlign);
    tcPr += `<w:vAlign w:val="${vAlign}"/>`;
  }

  if (tcPr) {
    return `<w:tcPr>${tcPr}</w:tcPr>`;
  }

  return '';
}

/**
 * Parse table width
 * @param {string} width - Width value
 * @returns {number} - Width in fiftieths of a percent
 */
function parseTableWidth(width) {
  if (!width) return 5000; // Default 100%

  if (width.includes('%')) {
    return parseInt(width) * 50; // Convert percentage to fiftieths
  } else if (width.includes('px')) {
    // Convert px to fiftieths (rough approximation)
    return parseInt(width) * 50;
  }

  return 5000;
}

/**
 * Parse padding value
 * @param {string} padding - Padding value
 * @returns {Object} - Padding object
 */
function parsePadding(padding) {
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
 * Get vertical alignment value
 * @param {string} align - Vertical alignment
 * @returns {string} - OOXML vertical alignment value
 */
function getVerticalAlignment(align) {
  switch (align) {
    case 'top':
      return 'top';
    case 'middle':
      return 'center';
    case 'bottom':
      return 'bottom';
    default:
      return 'top';
  }
}

/**
 * Parse border width to OOXML border size (eighths of a point)
 * @param {string} width - Border width (e.g., '1px', '2px')
 * @returns {number} - Border size in eighths of a point
 */
function parseBorderWidth(width) {
  if (!width) return 4; // Default border size
  
  const numValue = parseFloat(width);
  if (width.includes('px')) {
    // Convert px to eighths of a point (rough approximation)
    // 1px ≈ 0.75pt, 0.75pt * 8 = 6 eighths of a point
    return Math.max(2, Math.min(96, Math.round(numValue * 6)));
  }
  
  return 4; // Default
}

/**
 * Create OOXML text box with content
 * @param {string} content - Text box content (paragraph OOXML)
 * @param {Object} styles - Text box styles
 * @returns {string} - OOXML text box wrapped in paragraph
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
  
  // Build the text box OOXML using VML
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
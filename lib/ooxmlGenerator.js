import { escapeXml, getAlignment, parseDimension, parsePadding } from './utils.js';

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
    // Simple text - ensure we always have content
    const textContent = text || '';
    runs = createRun(textContent, styles);
  }

  // Ensure we have at least an empty run if no runs were created
  if (!runs) {
    runs = createRun('', styles);
  }

  const pPr = createParagraphProperties(styles);

  return `<w:p>${pPr}${runs}</w:p>`;
}

/**
 * Create OOXML run (text with formatting)
 * @param {string} text - Run text
 * @param {Object} styles - Run styles
 * @returns {string} - OOXML run
 */
export function createRun(text, styles) {
  // Handle empty or whitespace-only text
  const textContent = text || '';
  
  const rPr = createRunProperties(styles);
  const escapedText = escapeXml(textContent);

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

  // Handle spacing
  const spacingXml = createSpacingProperties(styles);
  if (spacingXml) {
    pPr += spacingXml;
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
 * Create spacing properties OOXML
 * @param {Object} styles - Paragraph styles
 * @returns {string} - OOXML spacing properties
 */
function createSpacingProperties(styles) {
  // Handle spacing - for content inside textbox, override with explicit values
  if (styles.insideTextBox) {
    // Force spacing to 0 for textbox content, and use exact line height
    const beforeVal = styles.spacingBefore !== undefined ? styles.spacingBefore : 0;
    const afterVal = styles.spacingAfter !== undefined ? styles.spacingAfter : 0;
    return `<w:spacing w:before="${beforeVal}" w:after="${afterVal}" w:line="240" w:lineRule="exact"/>`;
  }

  if (styles.noSpacing) {
    // No spacing for paragraphs with no-spacing attribute - single line spacing with no before/after
    return '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>';
  }

  if (styles.lineHeight !== undefined) {
    // Custom line height
    const lineSpacing = Math.round(styles.lineHeight * 240);
    return `<w:spacing w:line="${lineSpacing}" w:lineRule="auto"/>`;
  }

  if (styles.spacingBefore !== undefined || styles.spacingAfter !== undefined) {
    // Normal spacing for non-textbox content
    let spacingAttrs = '';
    if (styles.spacingBefore !== undefined) spacingAttrs += ` w:before="${styles.spacingBefore}"`;
    if (styles.spacingAfter !== undefined) spacingAttrs += ` w:after="${styles.spacingAfter}"`;
    return `<w:spacing${spacingAttrs}/>`;
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
 * @param {Element} element - Table element (optional)
 * @returns {string} - OOXML table
 */
export function createTable(content, styles, element = null) {
  const tblPr = createTableProperties(styles);
  const tblGrid = createTableGrid(styles, element);

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
    const widthInfo = parseTableWidth(styles.width);
    tblPr += `<w:tblW w:w="${widthInfo.value}" w:type="${widthInfo.type}"/>`;
  } else {
    // Default to auto width
    tblPr += '<w:tblW w:w="0" w:type="auto"/>';
  }

  // Border collapse and borders
  if (styles.borderCollapse === 'collapse' || styles.border) {
    tblPr += '<w:tblBorders>';
    tblPr += createBorderElements(styles.border, true);
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
 * @param {Element} element - Table element (optional)
 * @returns {string} - OOXML table grid
 */
export function createTableGrid(styles, element = null) {
  let gridCols = '';
  
  if (element) {
    // Get the first row to determine column count and widths
    const firstRow = element.querySelector('tr');
    if (firstRow) {
      const cells = firstRow.querySelectorAll('td, th');
      
      // Extract widths from cells
      cells.forEach((cell) => {
        const cellStyle = cell.getAttribute('style') || '';
        const widthMatch = cellStyle.match(/width:\s*([\d.]+)(px|%)/);
        
        let width = 2000; // Default width in DXA (twips)
        if (widthMatch) {
          const value = parseFloat(widthMatch[1]);
          const unit = widthMatch[2];
          
          if (unit === 'px') {
            // Convert px to DXA (twips): 1 px ≈ 15 DXA
            width = Math.round(value * 15);
          } else if (unit === '%') {
            // For percentage, use proportional width
            width = Math.round(value * 100);
          }
        }
        
        gridCols += `<w:gridCol w:w="${width}"/>`;
      });
    }
  }
  
  // Fallback if no columns detected
  if (!gridCols) {
    gridCols = '<w:gridCol w:w="4788"/><w:gridCol w:w="4788"/>';
  }
  
  return `<w:tblGrid>${gridCols}</w:tblGrid>`;
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
    const heightPx = parseDimension(styles.height);
    if (heightPx) {
      // Convert pixels to DXA (twips): 1 px ≈ 15 DXA
      const height = Math.round(heightPx * 15);
      trPr += `<w:trHeight w:val="${height}" w:hRule="atLeast"/>`;
    }
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
    const widthInfo = parseTableWidth(styles.width);
    tcPr += `<w:tcW w:w="${widthInfo.value}" w:type="${widthInfo.type}"/>`;
  }

  // Background color
  if (styles.backgroundColor) {
    tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${styles.backgroundColor}"/>`;
  }

  // Borders
  if (styles.border) {
    tcPr += '<w:tcBorders>';
    tcPr += createBorderElements(styles.border, false);
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
 * @returns {Object} - Object with value and type
 */
function parseTableWidth(width) {
  if (!width) return { value: 5000, type: 'pct' }; // Default 100%

  if (width.includes('%')) {
    return { 
      value: parseInt(width) * 50, 
      type: 'pct' 
    }; // Convert percentage to fiftieths
  } else if (width.includes('px')) {
    // Convert px to DXA (twips): 1 px ≈ 15 DXA
    return { 
      value: Math.round(parseInt(width) * 15), 
      type: 'dxa' 
    };
  }

  return { value: 5000, type: 'pct' };
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
 * Create border elements OOXML
 * @param {Object} border - Border configuration
 * @param {boolean} includeInside - Whether to include inside borders (for tables)
 * @returns {string} - OOXML border elements
 */
function createBorderElements(border, includeInside = false) {
  if (!border) return '';

  const borderColor = border.color || 'auto';
  const borderSize = parseBorderWidth(border.width);

  let borders = `<w:top w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>` +
                `<w:left w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>` +
                `<w:bottom w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>` +
                `<w:right w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;

  if (includeInside) {
    borders += `<w:insideH w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>` +
               `<w:insideV w:val="single" w:sz="${borderSize}" w:space="0" w:color="${borderColor}"/>`;
  }

  return borders;
}
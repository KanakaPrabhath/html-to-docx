/**
 * OOXML Templates for DOCX generation
 */

/**
 * Generate [Content_Types].xml
 */
function generateContentTypesXml(options = {}) {
  let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`;

  if (options.header) {
    contentTypes += `
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`;
  }

  if (options.footer) {
    contentTypes += `
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`;
  }

  contentTypes += `
</Types>`;

  return contentTypes;
}

/**
 * Generate _rels/.rels
 */
function generateRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function generateDocumentRelsXml(options = {}) {
  let relationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`;

  if (options.header && options.enableHeader !== false) {
    relationships += `
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`;
  }

  if (options.footer && options.enableFooter !== false) {
    const nextId = (options.header && options.enableHeader !== false) ? 'rId4' : 'rId3';
    relationships += `
  <Relationship Id="${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`;
  }

  relationships += `
</Relationships>`;

  return relationships;
}

/**
 * Generate word/document.xml
 */
function generateDocumentXml(bodyContent, options = {}) {
  const sectPr = generateSectPr(options);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:o="urn:schemas-microsoft-com:office:office">
  <w:body>
    ${bodyContent}
    ${sectPr}
  </w:body>
</w:document>`;
}

/**
 * Generate word/styles.xml
 */
function generateStylesXml(options) {
  const fontSize = options.fontSize || 11;
  const fontFamily = options.fontFamily || 'Calibri';
  const lineHeight = options.lineHeight || 1.15;
  const fontSizeHalfPoints = fontSize * 2;
  const lineSpacing = Math.round(lineHeight * 240);

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:pPr>
      <w:spacing w:line="${lineSpacing}" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}"/>
      <w:sz w:val="${fontSizeHalfPoints}"/>
      <w:szCs w:val="${fontSizeHalfPoints}"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="240" w:after="120"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}"/>
      <w:b/>
      <w:sz w:val="32"/>
      <w:szCs w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="240" w:after="120"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}"/>
      <w:b/>
      <w:sz w:val="26"/>
      <w:szCs w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="Heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="240" w:after="120"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}"/>
      <w:b/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
</w:styles>`;
}

/**
 * Generate word/fontTable.xml
 */
function generateFontTableXml(fontFamily = 'Calibri') {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="${fontFamily}">
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
  </w:font>
</w:fonts>`;
}

/**
 * Generate word/settings.xml
 */
function generateSettingsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>`;
}

/**
 * Generate word/webSettings.xml
 */
function generateWebSettingsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:optimizeForBrowser/>
  <w:allowPNG/>
</w:webSettings>`;
}

/**
 * Generate word/numbering.xml
 */
function generateNumberingXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:nsid w:val="1B2C3D4E"/>
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="◦"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="▪"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="2160" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:nsid w:val="2B3C4D5E"/>
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%3."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="2160" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>`;
}

/**
 * Generate docProps/app.xml
 */
function generateAppXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>TuteMaker HTML to DOCX</Application>
  <AppVersion>16.0000</AppVersion>
</Properties>`;
}

/**
 * Generate word/header1.xml
 */
function generateHeaderXml(headerContent) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:v="urn:schemas-microsoft-com:vml">
  ${headerContent}
</w:hdr>`;
}

/**
 * Generate word/footer1.xml
 */
function generateFooterXml(footerContent) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
       xmlns:wps="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingShape"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:v="urn:schemas-microsoft-com:vml">
  ${footerContent}
</w:ftr>`;
}

/**
 * Generate page number OOXML for footer as text box
 */
function generatePageNumberOOXML(alignment = 'right') {
  const alignmentMap = {
    left: 'left',
    center: 'center',
    right: 'right'
  };

  const jcVal = alignmentMap[alignment] || 'right';

  return `<w:p>
    <w:pPr>
      <w:jc w:val="${jcVal}"/>
    </w:pPr>
    <w:r>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:instrText xml:space="preserve"> PAGE \* MERGEFORMAT </w:instrText>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <w:t>1</w:t>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>`;
}

/**
 * Generate docProps/core.xml
 */
function generateCoreXml() {
  const now = new Date().toISOString().split('T')[0] + 'T00:00:00Z';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
                   xmlns:dc="http://purl.org/dc/elements/1.1/" 
                   xmlns:dcterms="http://purl.org/dc/terms/" 
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>TuteMaker</dc:creator>
  <cp:lastModifiedBy>TuteMaker</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

/**
 * Generate section properties (sectPr) with header/footer references
 */
function generateSectPr(options = {}) {
  let sectPrContent = '';

  // Page size options - default to A4
  const pageSize = options.pageSize || 'A4';
  let pageWidth, pageHeight;

  switch (pageSize.toUpperCase()) {
    case 'A4':
      pageWidth = 11906; // 8.27 inches in twips
      pageHeight = 16838; // 11.69 inches in twips
      break;
    case 'LETTER':
      pageWidth = 12240; // 8.5 inches in twips
      pageHeight = 15840; // 11 inches in twips
      break;
    case 'LEGAL':
      pageWidth = 12240; // 8.5 inches in twips
      pageHeight = 20160; // 14 inches in twips
      break;
    default:
      // Custom size: expect { width: number, height: number } in inches
      if (typeof pageSize === 'object' && pageSize.width && pageSize.height) {
        pageWidth = Math.round(pageSize.width * 1440); // Convert inches to twips
        pageHeight = Math.round(pageSize.height * 1440);
      } else {
        // Default to A4
        pageWidth = 11906;
        pageHeight = 16838;
      }
  }

  sectPrContent += `
    <w:pgSz w:w="${pageWidth}" w:h="${pageHeight}"/>`;

  // Page margins - default to 1 inch (1440 twips)
  const marginTop = options.marginTop !== undefined ? Math.round(options.marginTop * 1440) : 1440;
  const marginRight = options.marginRight !== undefined ? Math.round(options.marginRight * 1440) : 1440;
  const marginBottom = options.marginBottom !== undefined ? Math.round(options.marginBottom * 1440) : 1440;
  const marginLeft = options.marginLeft !== undefined ? Math.round(options.marginLeft * 1440) : 1440;
  const marginHeader = options.marginHeader !== undefined ? Math.round(options.marginHeader * 1440) : 720; // 0.5 inch default for header
  const marginFooter = options.marginFooter !== undefined ? Math.round(options.marginFooter * 1440) : 720; // 0.5 inch default for footer
  const marginGutter = options.marginGutter !== undefined ? Math.round(options.marginGutter * 1440) : 0;

  sectPrContent += `
    <w:pgMar w:top="${marginTop}" w:right="${marginRight}" w:bottom="${marginBottom}" w:left="${marginLeft}" w:header="${marginHeader}" w:footer="${marginFooter}" w:gutter="${marginGutter}"/>`;

  // Page borders
  if (options.enablePageBorder || options.pageBorder) {
    if (options.pageBorder && options.pageBorder.radius) {
      // Use shape instead of pgBorders for rounded corners
    } else {
      const borderConfig = typeof options.pageBorder === 'object' ? options.pageBorder : {};
      const borderStyle = borderConfig.style || 'single';
      const borderColor = borderConfig.color || '000000';
      const borderSize = borderConfig.size || 4;
      
      sectPrContent += `
    <w:pgBorders>
      <w:top w:val="${borderStyle}" w:sz="${borderSize}" w:space="${marginTop}" w:color="${borderColor}"/>
      <w:left w:val="${borderStyle}" w:sz="${borderSize}" w:space="${marginLeft}" w:color="${borderColor}"/>
      <w:bottom w:val="${borderStyle}" w:sz="${borderSize}" w:space="${marginBottom}" w:color="${borderColor}"/>
      <w:right w:val="${borderStyle}" w:sz="${borderSize}" w:space="${marginRight}" w:color="${borderColor}"/>
    </w:pgBorders>`;
    }
  }

  let nextId = 3;

  if (options.header && options.enableHeader !== false) {
    sectPrContent += `
    <w:headerReference w:type="default" r:id="rId${nextId}"/>`;
    nextId++;
  }

  if (options.footer && options.enableFooter !== false) {
    sectPrContent += `
    <w:footerReference w:type="default" r:id="rId${nextId}"/>`;
  }

  return `<w:sectPr>${sectPrContent}
  </w:sectPr>`;
}

export {
  generateContentTypesXml,
  generateRelsXml,
  generateDocumentRelsXml,
  generateDocumentXml,
  generateStylesXml,
  generateFontTableXml,
  generateSettingsXml,
  generateWebSettingsXml,
  generateNumberingXml,
  generateAppXml,
  generateCoreXml,
  generateHeaderXml,
  generateFooterXml,
  generatePageNumberOOXML,
  generateSectPr
};

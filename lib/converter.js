import JSZip from 'jszip';
import fs from 'fs/promises';
import { parseHtml } from './htmlParser.js';
import { MediaManager, createFullWidthImageOOXML } from './mediaHandler.js';
import {
  generateDocumentXml,
  generateRelsXml,
  generateAppXml,
  generateCoreXml,
  generateStylesXml,
  generateFontTableXml,
  generateSettingsXml,
  generateWebSettingsXml,
  generateNumberingXml,
  generateHeaderXml,
  generateFooterXml,
  generatePageNumberOOXML
} from './templates.js';
import { generateRoundedBorderShape } from './vmlGenerator.js';
import { sanitizeOptions } from './utils.js';

/**
 * Main conversion logic for HTML to DOCX
 */

/**
 * Convert HTML string to DOCX buffer
 * @param {string} html - HTML content to convert
 * @param {Object} options - Conversion options
 * @returns {Promise<Buffer>} - DOCX file buffer
 */
export async function convertHtmlToDocx(html, options = {}) {
  const sanitizedOptions = sanitizeOptions(options);
  const zip = new JSZip();
  const mediaManager = new MediaManager();

  // Parse HTML
  const bodyContent = await parseHtml(html, mediaManager, sanitizedOptions);

  // Parse header and footer if provided
  let headerContent = '';
  let footerContent = '';

  if (sanitizedOptions.header && sanitizedOptions.enableHeader !== false) {
    if (sanitizedOptions.header.startsWith('data:image/')) {
      // Handle base64 image header
      headerContent = await createFullWidthImageOOXML(sanitizedOptions.header, mediaManager, 'header', sanitizedOptions);
    } else {
      // Handle HTML header
      headerContent = await parseHtml(sanitizedOptions.header, mediaManager, sanitizedOptions, 'header');
    }
  }

  if (sanitizedOptions.footer && sanitizedOptions.enableFooter !== false) {
    if (sanitizedOptions.footer.startsWith('data:image/')) {
      // Handle base64 image footer
      footerContent = await createFullWidthImageOOXML(sanitizedOptions.footer, mediaManager, 'footer', sanitizedOptions);
    } else {
      // Handle HTML footer
      footerContent = await parseHtml(sanitizedOptions.footer, mediaManager, sanitizedOptions, 'footer');
    }
  }

  // Add rounded border shape to header if page border is enabled (always use VML shape)
  if (sanitizedOptions.enablePageBorder || sanitizedOptions.pageBorder) {
    const shapeXml = generateRoundedBorderShape(sanitizedOptions);
    if (headerContent) {
      headerContent = shapeXml + headerContent;
    } else {
      headerContent = shapeXml;
      sanitizedOptions.enableHeader = true;
    }
  }

  // Add page numbers to footer if enabled
  if (sanitizedOptions.enablePageNumbers) {
    const pageNumberXml = generatePageNumberOOXML(sanitizedOptions.pageNumberAlignment || 'center');
    if (footerContent) {
      footerContent += pageNumberXml;
    } else {
      footerContent = pageNumberXml;
      sanitizedOptions.enableFooter = true;
    }
  }

  // Add required files to zip
  zip.file('[Content_Types].xml', mediaManager.generateContentTypesXml(sanitizedOptions));

  // _rels folder
  zip.file('_rels/.rels', generateRelsXml());

  // word folder
  zip.file('word/document.xml', generateDocumentXml(bodyContent, sanitizedOptions));
  zip.file('word/_rels/document.xml.rels', mediaManager.generateDocumentRelsXml(sanitizedOptions));
  zip.file('word/styles.xml', generateStylesXml(sanitizedOptions));
  zip.file('word/numbering.xml', generateNumberingXml());

  // Add header and footer files if provided
  if (headerContent) {
    zip.file('word/header1.xml', generateHeaderXml(headerContent));
    const headerRels = mediaManager.generateHeaderRelsXml();
    if (headerRels) {
      zip.file('word/_rels/header1.xml.rels', headerRels);
    }
  }

  if (footerContent) {
    zip.file('word/footer1.xml', generateFooterXml(footerContent));
    const footerRels = mediaManager.generateFooterRelsXml();
    if (footerRels) {
      zip.file('word/_rels/footer1.xml.rels', footerRels);
    }
  }

  // Add media files
  const mediaFiles = mediaManager.getMediaFiles();
  if (mediaFiles.length > 0) {
    mediaFiles.forEach(mediaFile => {
      zip.file(`word/media/${mediaFile.filename}`, mediaFile.buffer);
    });
  }

  // docProps folder
  zip.file('docProps/app.xml', generateAppXml());
  zip.file('docProps/core.xml', generateCoreXml());

  // Generate zip buffer
  const buffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 }
  });

  return buffer;
}

/**
 * Convert HTML string to DOCX file
 * @param {string} html - HTML content to convert
 * @param {string} outputPath - Path where to save the DOCX file
 * @param {Object} options - Conversion options
 * @returns {Promise<void>}
 */
export async function convertHtmlToDocxFile(html, outputPath, options = {}) {
  const buffer = await convertHtmlToDocx(html, options);
  await fs.writeFile(outputPath, buffer);
}
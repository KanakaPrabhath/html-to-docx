import JSZip from 'jszip';
import fs from 'fs/promises';
import { parseHtml } from './htmlParser.js';
import { MediaManager } from './mediaHandler.js';
import {
  generateDocumentXml,
  generateContentTypesXml,
  generateRelsXml,
  generateDocumentRelsXml,
  generateAppXml,
  generateCoreXml,
  generateStylesXml,
  generateFontTableXml,
  generateSettingsXml,
  generateWebSettingsXml,
  generateNumberingXml
} from './templates.js';

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
  const zip = new JSZip();
  const mediaManager = new MediaManager();

  // Parse HTML
  const bodyContent = await parseHtml(html, mediaManager);

  // Add required files to zip
  zip.file('[Content_Types].xml', mediaManager.generateContentTypesXml());

  // _rels folder
  zip.file('_rels/.rels', generateRelsXml());

  // word folder
  zip.file('word/document.xml', generateDocumentXml(bodyContent));
  zip.file('word/_rels/document.xml.rels', mediaManager.generateDocumentRelsXml());
  zip.file('word/styles.xml', generateStylesXml(options));
  zip.file('word/numbering.xml', generateNumberingXml());

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
import JSZip from 'jszip';
import fs from 'fs/promises';
import { parseHtml } from './htmlParser.js';
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

  // Parse HTML
  const bodyContent = parseHtml(html);

  // Add required files to zip
  zip.file('[Content_Types].xml', generateContentTypesXml());

  // _rels folder
  zip.file('_rels/.rels', generateRelsXml());

  // word folder
  zip.file('word/document.xml', generateDocumentXml(bodyContent));
  zip.file('word/_rels/document.xml.rels', generateDocumentRelsXml());
  zip.file('word/styles.xml', generateStylesXml(options));
  zip.file('word/numbering.xml', generateNumberingXml());

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
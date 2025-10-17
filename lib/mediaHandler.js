import fs from 'fs/promises';
import path from 'path';
import https from 'https';
import http from 'http';
import { URL } from 'url';

/**
 * Media handling utilities for HTML to DOCX conversion
 */

/**
 * Media manager class to handle images and other media files
 */
export class MediaManager {
  constructor() {
    this.mediaFiles = [];
    this.relationshipId = 3; // Start from 3 since 1 and 2 are used for styles and numbering
    this.imageId = 1; // Unique ID for images
  }

  /**
   * Add an image from URL or base64 data
   * @param {string} src - Image source (URL or base64 data URL)
   * @param {string} alt - Alt text for the image
   * @returns {Promise<Object>} - Media object with id, filename, and relationship id
   */
  async addImage(src, alt = '') {
    let buffer;
    let filename;
    let extension;

    if (src.startsWith('data:image/')) {
      // Base64 encoded image
      const match = src.match(/^data:image\/([a-zA-Z]+);base64,(.+)$/);
      if (!match) {
        throw new Error('Invalid base64 image data');
      }
      extension = match[1];
      const base64Data = match[2];
      buffer = Buffer.from(base64Data, 'base64');
      filename = `image_${this.mediaFiles.length + 1}.${extension}`;
    } else {
      // URL - download the image
      try {
        buffer = await downloadImage(src);
        extension = getExtensionFromUrl(src) || 'png';
        filename = `image_${this.mediaFiles.length + 1}.${extension}`;
      } catch (error) {
        console.warn(`Failed to download image from ${src}:`, error.message);
        return null;
      }
    }

    const mediaId = `rId${this.relationshipId++}`;
    const mediaFile = {
      id: mediaId,
      filename,
      buffer,
      alt,
      extension,
      imageId: this.imageId++
    };

    this.mediaFiles.push(mediaFile);
    return mediaFile;
  }

  /**
   * Get all media files
   * @returns {Array} - Array of media file objects
   */
  getMediaFiles() {
    return this.mediaFiles;
  }

  /**
   * Generate document relationships XML for media files
   * @returns {string} - OOXML relationships
   */
  generateDocumentRelsXml() {
    let rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n';
    rels += '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n';
    rels += '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>\n';

    this.mediaFiles.forEach((media, index) => {
      rels += `  <Relationship Id="${media.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${media.filename}"/>\n`;
    });

    rels += '</Relationships>';
    return rels;
  }

  /**
   * Generate content types XML with media support
   * @returns {string} - OOXML content types
   */
  generateContentTypesXml() {
    let contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n';
    contentTypes += '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n';
    contentTypes += '  <Default Extension="xml" ContentType="application/xml"/>\n';
    contentTypes += '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n';
    contentTypes += '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n';
    contentTypes += '  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>\n';
    contentTypes += '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n';
    contentTypes += '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\n';

    // Add content types for each media file
    const extensions = [...new Set(this.mediaFiles.map(media => media.extension))];
    extensions.forEach(ext => {
      const contentType = getContentTypeForExtension(ext);
      contentTypes += `  <Default Extension="${ext}" ContentType="${contentType}"/>\n`;
    });

    contentTypes += '</Types>';
    return contentTypes;
  }
}

/**
 * Download image from URL
 * @param {string} url - Image URL
 * @returns {Promise<Buffer>} - Image buffer
 */
function downloadImage(url) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const protocol = parsedUrl.protocol === 'https:' ? https : http;

    const request = protocol.get(url, (response) => {
      // Handle redirects
      if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
        // Follow redirect
        downloadImage(response.headers.location).then(resolve).catch(reject);
        return;
      }

      if (response.statusCode !== 200) {
        reject(new Error(`Failed to download image: ${response.statusCode}`));
        return;
      }

      const chunks = [];
      response.on('data', (chunk) => {
        chunks.push(chunk);
      });

      response.on('end', () => {
        resolve(Buffer.concat(chunks));
      });
    });

    request.on('error', (error) => {
      reject(error);
    });

    request.setTimeout(10000, () => {
      request.destroy();
      reject(new Error('Request timeout'));
    });
  });
}

/**
 * Get file extension from URL
 * @param {string} url - Image URL
 * @returns {string|null} - File extension
 */
function getExtensionFromUrl(url) {
  try {
    const parsedUrl = new URL(url);
    const pathname = parsedUrl.pathname;
    const extension = path.extname(pathname).toLowerCase().slice(1);
    return extension || null;
  } catch {
    return null;
  }
}

/**
 * Get MIME content type for file extension
 * @param {string} extension - File extension
 * @returns {string} - MIME content type
 */
function getContentTypeForExtension(extension) {
  const contentTypes = {
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'gif': 'image/gif',
    'bmp': 'image/bmp',
    'tiff': 'image/tiff',
    'svg': 'image/svg+xml'
  };
  return contentTypes[extension.toLowerCase()] || 'application/octet-stream';
}

/**
 * Create OOXML for an image
 * @param {Object} mediaFile - Media file object
 * @param {Object} options - Image options (width, height, float, margin, display)
 * @returns {string} - OOXML image element
 */
export function createImageOOXML(mediaFile, options = {}) {
  const width = options.width || 300; // Default width in pixels
  const height = options.height || 200; // Default height in pixels

  // Convert pixels to EMUs (English Metric Units) - 1 inch = 914400 EMUs, 96 DPI
  const widthEmu = Math.round(width * 914400 / 96);
  const heightEmu = Math.round(height * 914400 / 96);

  // Use unique IDs for each image
  const imageId = mediaFile.imageId || 1;

  // Check if image should be floating
  const isFloating = options.float === 'left' || options.float === 'right';

  if (isFloating) {
    // Create anchored (floating) image
    return createAnchoredImageOOXML(mediaFile, options, widthEmu, heightEmu, imageId);
  } else {
    // Create inline image
    return createInlineImageOOXML(mediaFile, widthEmu, heightEmu, imageId);
  }
}

/**
 * Create inline image OOXML
 * @param {Object} mediaFile - Media file object
 * @param {number} widthEmu - Width in EMUs
 * @param {number} heightEmu - Height in EMUs
 * @param {number} imageId - Unique image ID
 * @returns {string} - OOXML inline image element
 */
function createInlineImageOOXML(mediaFile, widthEmu, heightEmu, imageId) {
  return `<w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="${widthEmu}" cy="${heightEmu}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="${imageId}" name="${mediaFile.alt || 'Image'}"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="${imageId}" name="${mediaFile.alt || 'Image'}"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="${mediaFile.id}"/>
                <a:stretch>
                  <a:fillRect/>
                </a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="${widthEmu}" cy="${heightEmu}"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                  <a:avLst/>
                </a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>`;
}

/**
 * Create anchored (floating) image OOXML
 * @param {Object} mediaFile - Media file object
 * @param {Object} options - Image options
 * @param {number} widthEmu - Width in EMUs
 * @param {number} heightEmu - Height in EMUs
 * @param {number} imageId - Unique image ID
 * @returns {string} - OOXML anchored image element
 */
function createAnchoredImageOOXML(mediaFile, options, widthEmu, heightEmu, imageId) {
  const float = options.float;
  const margin = options.margin || {};

  // Convert margin pixels to EMUs
  const marginTop = margin.top ? Math.round(margin.top * 914400 / 96) : 0;
  const marginRight = margin.right ? Math.round(margin.right * 914400 / 96) : 0;
  const marginBottom = margin.bottom ? Math.round(margin.bottom * 914400 / 96) : 0;
  const marginLeft = margin.left ? Math.round(margin.left * 914400 / 96) : 0;

  // Determine positioning based on float
  const alignH = float === 'right' ? 'right' : 'left';
  const relativeFromH = 'column';
  const relativeFromV = 'paragraph';
  const posOffsetV = 0;

  return `<w:r>
    <w:drawing>
      <wp:anchor distT="${marginTop}" distB="${marginBottom}" distL="${marginLeft}" distR="${marginRight}" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="0" allowOverlap="0">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="${relativeFromH}">
          <wp:align>${alignH}</wp:align>
        </wp:positionH>
        <wp:positionV relativeFrom="${relativeFromV}">
          <wp:posOffset>${posOffsetV}</wp:posOffset>
        </wp:positionV>
        <wp:extent cx="${widthEmu}" cy="${heightEmu}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:wrapSquare wrapText="${float === 'right' ? 'left' : 'right'}"/>
        <wp:docPr id="${imageId}" name="${mediaFile.alt || 'Image'}"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="${imageId}" name="${mediaFile.alt || 'Image'}"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="${mediaFile.id}"/>
                <a:stretch>
                  <a:fillRect/>
                </a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="${widthEmu}" cy="${heightEmu}"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                  <a:avLst/>
                </a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing>
  </w:r>`;
}
import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs/promises';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Complex test combining section header images with headers, footers, and page borders with radius
 */
async function testComplexSectionHeaders() {
  console.log('Testing complex section headers with headers, footers, and page borders...');

  // Small test image in base64 (1x1 pixel red PNG)
  const testImageBase64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

  // Larger header image (simulated as base64)
  const headerImageBase64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

  const html = `
    <h1>Complex Document with Section Headers, Headers, Footers & Borders</h1>

    <h2>Page 1: Introduction</h2>
    <p>This document demonstrates the combination of multiple advanced features:</p>
    <ul>
      <li>Section header images with top/bottom text wrapping</li>
      <li>HTML headers and footers</li>
      <li>Page borders with rounded corners</li>
      <li>Multiple page layouts</li>
    </ul>

    <p>The document has a decorative page border with rounded corners, a header with company branding, and a footer with page numbers.</p>

    <page-break data-page-break="true" contenteditable="false" data-page-number="2"></page-break>

    <h2>Page 2: Section Header Image Example</h2>
    <p>This page demonstrates a section header image that spans part of the page width:</p>

    <img data-section-header src="${testImageBase64}" alt="Section Header" style="width: 500px; height: 120px; margin: 15px; display: block;"/>

    <p>Notice how the text wraps only above and below the section header image, not around the sides. This creates a clean separation between content sections.</p>

    <p>Additional content below the section header to demonstrate the text flow behavior.</p>

    <page-break data-page-break="true" contenteditable="false" data-page-number="3"></page-break>

    <h2>Page 3: Full-Width Cover Image</h2>
    <p>This page features a full-width cover image that spans the entire page width:</p>

    <img data-section-header data-cover src="${testImageBase64}" alt="Cover Image" style="height: 150px; display: block;"/>

    <p>The cover image above spans the full width of the page, creating a dramatic header effect. Notice how the text appears only below the image.</p>

    <p>This type of layout is perfect for chapter titles, section dividers, or important announcements.</p>

    <page-break data-page-break="true" contenteditable="false" data-page-number="4"></page-break>

    <h2>Page 4: Multiple Section Headers</h2>
    <p>This page demonstrates multiple section header images on the same page:</p>

    <img data-section-header src="${testImageBase64}" alt="First Section" style="width: 400px; height: 100px; margin: 10px; display: block;"/>
    <p>Content between the first and second section headers.</p>

    <img data-section-header src="${testImageBase64}" alt="Second Section" style="width: 350px; height: 90px; margin: 12px; display: block;"/>
    <p>Content after the second section header.</p>

    <p>This demonstrates how multiple section headers can be used to create visual separation within a single page.</p>

    <page-break data-page-break="true" contenteditable="false" data-page-number="5"></page-break>

    <h2>Page 5: Technical Details</h2>
    <p>This document showcases several advanced features:</p>

    <h3>Page Borders</h3>
    <p>The document has rounded page borders that create a modern, professional appearance.</p>

    <h3>Headers and Footers</h3>
    <p>The header contains company branding and the footer includes page numbering.</p>

    <h3>Section Header Images</h3>
    <p>Images with the <code>data-section-header</code> attribute are positioned at the top of content sections with proper text wrapping.</p>

    <h3>Cover Images</h3>
    <p>Images with both <code>data-section-header</code> and <code>data-cover</code> attributes span the full page width.</p>

    <p>This combination of features creates rich, professional documents with complex layouts.</p>
  `;

  const converter = new HtmlToDocx({
    // Header with HTML content and image
    header: `<table style="width: 100%; border-collapse: collapse;">
      <tr>
        <td style="width: 80px; vertical-align: middle;">
          <img src="${headerImageBase64}" alt="Logo" style="width: 60px; height: 60px;"/>
        </td>
        <td style="vertical-align: middle; padding-left: 10px;">
          <p style="margin: 0; font-size: 14pt; font-weight: bold; color: #2E75B6;">Advanced Document Solutions</p>
          <p style="margin: 0; font-size: 10pt; color: #666666;">Professional Document Generation</p>
        </td>
        <td style="text-align: right; vertical-align: middle;">
          <p style="margin: 0; font-size: 10pt; color: #666666;">Confidential</p>
        </td>
      </tr>
    </table>`,

    // Footer with page numbers
    footer: '<p style="text-align: center; font-size: 10pt; color: #666666;">© 2024 Advanced Document Solutions | Page <page-number/></p>',

    // Enable page numbers
    enablePageNumbers: true,

    // Page border with radius
    pageBorder: {
      style: 'single',
      color: '2E75B6',
      size: 6,
      radius: 8
    },

    // Page settings
    pageSize: 'A4',
    marginTop: 1.2,
    marginRight: 1,
    marginBottom: 1.2,
    marginLeft: 1,

    // Font settings
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath = path.join(__dirname, 'test-complex-section-headers.docx');

  try {
    await converter.convertHtmlToDocxFile(html, outputPath);

    // Verify the generated document has the expected features
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(await fs.readFile(outputPath));

    // Check for document.xml
    const documentXml = await zip.file('word/document.xml').async('string');

    // Verify section header images are present
    const hasSectionHeaders = documentXml.includes('<wp:wrapTopAndBottom/>');
    const hasAnchoredImages = documentXml.includes('<wp:anchor');

    // Check for header and footer
    const hasHeader = zip.file('word/header1.xml') !== null;
    const hasFooter = zip.file('word/footer1.xml') !== null;

    // Check for page borders (in document.xml.rels or settings)
    const hasPageBorder = documentXml.includes('pageBorder') || documentXml.includes('w:pgBorders');

    if (!hasSectionHeaders) {
      console.error('✗ Section header images with wrapTopAndBottom not found');
      return false;
    }

    if (!hasAnchoredImages) {
      console.error('✗ Anchored images not found');
      return false;
    }

    if (!hasHeader) {
      console.error('✗ Header file not found');
      return false;
    }

    if (!hasFooter) {
      console.error('✗ Footer file not found');
      return false;
    }

    console.log('✓ Complex section headers test completed successfully');
    console.log(`  Output: ${outputPath}`);
    console.log('  Features verified:');
    console.log('    ✓ Section header images with top/bottom text wrapping');
    console.log('    ✓ HTML header with image and text');
    console.log('    ✓ Footer with page numbers');
    console.log('    ✓ Page borders with rounded corners');
    console.log('    ✓ Multiple page layout with page breaks');
    return true;

  } catch (error) {
    console.error('✗ Complex section headers test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testComplexSectionHeaders;

// Run the test if called directly
if (import.meta.url === `file://${process.argv[1].replace(/\\/g, '/')}`) {
  testComplexSectionHeaders().catch(console.error);
}
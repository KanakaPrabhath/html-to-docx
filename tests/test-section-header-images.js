import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs/promises';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test section header images (data-section-header and data-cover)
 */
async function testSectionHeaderImages() {
  console.log('Testing section header images...');

  // Small test image in base64 (1x1 pixel red PNG)
  const testImageBase64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

  const html = `
    <h1>Section Header Image Tests</h1>

    <h2>Regular Section Header Image</h2>
    <p>This section tests a regular section header image with custom width and margins:</p>
    <img data-section-header src="${testImageBase64}" alt="Section Header" style="width: 400px; height: 150px; margin: 20px; display: block;"/>

    <p>Content after the section header image. This should appear below the image.</p>

    <page-break data-page-break="true" contenteditable="false" data-page-number="2"></page-break>

    <h2>Cover Image (Full Width)</h2>
    <p>This section tests a cover image that spans the full page width:</p>
    <img data-section-header data-cover src="${testImageBase64}" alt="Cover Image" style="height: 200px; display: block;"/>

    <p>Content after the cover image. The cover image should span the full width of the page.</p>

    <page-break data-page-break="true" contenteditable="false" data-page-number="3"></page-break>

    <h2>Multiple Section Headers</h2>
    <p>Testing multiple section header images on the same page:</p>

    <img data-section-header src="${testImageBase64}" alt="First Header" style="width: 300px; height: 100px; margin: 10px; display: block;"/>
    <p>Content between headers.</p>

    <img data-section-header src="${testImageBase64}" alt="Second Header" style="width: 500px; height: 120px; margin: 15px; display: block;"/>
    <p>Content after second header.</p>

    <page-break data-page-break="true" contenteditable="false" data-page-number="4"></page-break>

    <h2>Section Header with Different Page Sizes</h2>
    <p>This tests how section headers work with different page configurations.</p>
    <img data-section-header data-cover src="${testImageBase64}" alt="Full Width Header" style="height: 180px; display: block;"/>

    <p>Final content section.</p>
  `;

  const converter = new HtmlToDocx({
    fontSize: 12,
    fontFamily: 'Arial',
    lineHeight: 1.5,
    pageSize: 'A4',
    marginTop: 1,
    marginRight: 1,
    marginBottom: 1,
    marginLeft: 1
  });

  const outputPath = path.join(__dirname, 'test-section-header-images.docx');

  try {
    await converter.convertHtmlToDocxFile(html, outputPath);
    // Verify generated OOXML for margin support on non-cover section header
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(await fs.readFile(outputPath));
    const documentXml = await zip.file('word/document.xml').async('string');

    // The first non-cover section header in the test uses margin: 20px
    // Convert 20px -> EMU (1 inch = 914400 EMUs, 96 DPI)
    const EMU_PER_INCH = 914400;
    const PIXELS_PER_INCH = 96;
    const expectedImageMarginPx = 20;
    const expectedImageMarginEmu = Math.round(expectedImageMarginPx * EMU_PER_INCH / PIXELS_PER_INCH);
    
    // Page margins are 1 inch each
    const pageMarginInch = 1;
    const expectedPageMarginEmu = Math.round(pageMarginInch * EMU_PER_INCH);
    
    // Expected position offsets: page margin + image margin
    const expectedPosOffsetH = expectedPageMarginEmu + expectedImageMarginEmu;
    const expectedPosOffsetV = expectedPageMarginEmu + expectedImageMarginEmu;

    // Find the first anchor element and check posOffset values
    const firstAnchorIndex = documentXml.indexOf('<wp:anchor');
    let positionCheck = false;
    if (firstAnchorIndex !== -1) {
      const anchorSnippet = documentXml.substring(firstAnchorIndex, firstAnchorIndex + 600);
      const expectedPosH = `<wp:posOffset>${expectedPosOffsetH}</wp:posOffset>`;
      const expectedPosV = `<wp:posOffset>${expectedPosOffsetV}</wp:posOffset>`;
      positionCheck = anchorSnippet.includes(expectedPosH) && anchorSnippet.includes(expectedPosV);
    }

    if (!positionCheck) {
      console.error('✗ Section header margin check failed: expected position offsets not found in document.xml');
      console.error(`  Expected image margin EMU for 20px: ${expectedImageMarginEmu}`);
      console.error(`  Expected page margin EMU for 1 inch: ${expectedPageMarginEmu}`);
      console.error(`  Expected posOffsetH: ${expectedPosOffsetH}`);
      console.error(`  Expected posOffsetV: ${expectedPosOffsetV}`);
      console.error('  You can open the generated DOCX to inspect word/document.xml for <wp:anchor> position offsets.');
      return false;
    }

    console.log('✓ Section header images test completed successfully');
    console.log(`  Output: ${outputPath}`);
    console.log('  Note: Section header images should be positioned at the top of pages');
    console.log('  Note: Cover images should span full page width');
    return true;
  } catch (error) {
    console.error('✗ Section header images test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testSectionHeaderImages;

// Run the test
testSectionHeaderImages().catch(console.error);
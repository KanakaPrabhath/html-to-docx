import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test headers, footers, and page numbers
 */
async function testHeadersFooters() {
  console.log('Testing headers and footers...');

  // Test 1: HTML headers and footers
  const html1 = `
    <h1>Document with HTML Header and Footer</h1>
    <p>This document demonstrates HTML headers and footers.</p>
    <p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>
    <p>Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.</p>
    <p>Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.</p>
    <p>Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>
    <h2>Section Break</h2>
    <p>More content to ensure the document spans multiple pages if needed.</p>
    <p>Additional content for testing page breaks and header/footer repetition.</p>
    <p>Final paragraph to complete the document.</p>
  `;

  const converter1 = new HtmlToDocx({
    header: '<p style="text-align: center; font-size: 10pt; color: #666666;">Company Confidential Document - Header</p>',
    footer: '<p style="text-align: center; font-size: 10pt; color: #666666;">Page Footer - © 2024 Company Name</p>',
    enablePageNumbers: true,
    pageNumberAlignment: 'center',
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath1 = path.join(__dirname, 'test-headers-footers-html.docx');

  try {
    await converter1.convertHtmlToDocxFile(html1, outputPath1);
    console.log('✓ HTML headers and footers test completed successfully');
    console.log(`  Output: ${outputPath1}`);
  } catch (error) {
    console.error('✗ HTML headers and footers test failed:', error.message);
  }

  // Test 2: Image headers and footers
  const smallImageBase64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

  const html2 = `
    <h1>Document with Image Header and Footer</h1>
    <p>This document has image-based headers and footers.</p>
    <p>The header contains a full-width image, and the footer also contains an image.</p>
    <p>Page numbers are enabled and centered.</p>
    <p>More content to fill the page.</p>
    <p>Lorem ipsum dolor sit amet, consectetur adipiscing elit.</p>
    <p>Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>
  `;

  const converter2 = new HtmlToDocx({
    header: smallImageBase64, // Base64 image for header
    footer: smallImageBase64, // Base64 image for footer
    enablePageNumbers: true,
    pageNumberAlignment: 'right',
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath2 = path.join(__dirname, 'test-headers-footers-images.docx');

  try {
    await converter2.convertHtmlToDocxFile(html2, outputPath2);
    console.log('✓ Image headers and footers test completed successfully');
    console.log(`  Output: ${outputPath2}`);
  } catch (error) {
    console.error('✗ Image headers and footers test failed:', error.message);
  }

  // Test 3: Only page numbers (no header/footer content)
  const html3 = `
    <h1>Document with Page Numbers Only</h1>
    <p>This document has only page numbers in the footer.</p>
    <p>No custom header or footer content is specified.</p>
    <p>Page numbers are left-aligned.</p>
    <p>Content continues...</p>
    <p>Lorem ipsum dolor sit amet, consectetur adipiscing elit.</p>
    <p>Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>
    <p>Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.</p>
  `;

  const converter3 = new HtmlToDocx({
    enablePageNumbers: true,
    pageNumberAlignment: 'left',
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath3 = path.join(__dirname, 'test-page-numbers-only.docx');

  try {
    await converter3.convertHtmlToDocxFile(html3, outputPath3);
    console.log('✓ Page numbers only test completed successfully');
    console.log(`  Output: ${outputPath3}`);
  } catch (error) {
    console.error('✗ Page numbers only test failed:', error.message);
  }

  // Test 4: Complex header with formatting
  const html4 = `
    <h1>Document with Complex Header</h1>
    <p>This document has a complex HTML header with formatting.</p>
    <p>The header includes bold text, colors, and alignment.</p>
    <p>Content continues here...</p>
  `;

  const converter4 = new HtmlToDocx({
    header: '<div style="text-align: center; font-size: 12pt;"><strong style="color: #FF0000;">CONFIDENTIAL</strong> - <span style="color: #666666;">Internal Document</span> - <em>Page</em></div>',
    footer: '<p style="text-align: right; font-size: 10pt;">Generated on ' + new Date().toLocaleDateString() + '</p>',
    enablePageNumbers: false, // No page numbers since footer has custom content
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath4 = path.join(__dirname, 'test-complex-header-footer.docx');

  try {
    await converter4.convertHtmlToDocxFile(html4, outputPath4);
    console.log('✓ Complex header and footer test completed successfully');
    console.log(`  Output: ${outputPath4}`);
    return true;
  } catch (error) {
    console.error('✗ Complex header and footer test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testHeadersFooters;

// Run the test
testHeadersFooters().catch(console.error);
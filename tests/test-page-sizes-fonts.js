import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test different page sizes and font customization
 */
async function testPageSizesAndFonts() {
  console.log('Testing page sizes and font customization...');

  // Test 1: Font customization
  const html1 = `
    <h1>Font Customization Test</h1>
    <p>This document demonstrates different font settings.</p>
    <h2>Default Font Settings</h2>
    <p>This paragraph uses the default font family and size.</p>
    <p><strong>Bold text</strong> and <em>italic text</em> should inherit the font settings.</p>

    <h2>Different Font Sizes</h2>
    <p><span style="font-size: 10px;">10px text</span></p>
    <p><span style="font-size: 12px;">12px text</span></p>
    <p><span style="font-size: 14px;">14px text</span></p>
    <p><span style="font-size: 16px;">16px text</span></p>
    <p><span style="font-size: 18px;">18px text</span></p>

    <h2>Different Font Families</h2>
    <p><span style="font-family: Arial;">Arial font</span></p>
    <p><span style="font-family: 'Times New Roman';">Times New Roman font</span></p>
    <p><span style="font-family: Calibri;">Calibri font</span></p>
    <p><span style="font-family: Verdana;">Verdana font</span></p>
  `;

  const fontSettings = [
    { name: 'default', fontSize: 11, fontFamily: 'Calibri', lineHeight: 1.15 },
    { name: 'large', fontSize: 14, fontFamily: 'Arial', lineHeight: 1.2 },
    { name: 'small', fontSize: 9, fontFamily: 'Times New Roman', lineHeight: 1.1 },
    { name: 'compact', fontSize: 10, fontFamily: 'Verdana', lineHeight: 1.0 }
  ];

  for (let i = 0; i < fontSettings.length; i++) {
    const fonts = fontSettings[i];
    const converter = new HtmlToDocx({
      fontSize: fonts.fontSize,
      fontFamily: fonts.fontFamily,
      lineHeight: fonts.lineHeight
    });

    const outputPath = path.join(__dirname, `test-fonts-${fonts.name}.docx`);

    try {
      await converter.convertHtmlToDocxFile(html1, outputPath);
      console.log(`✓ Font settings (${fonts.name}) test completed successfully`);
      console.log(`  Output: ${outputPath}`);
    } catch (error) {
      console.error(`✗ Font settings (${fonts.name}) test failed:`, error.message);
    }
  }

  // Test 2: Page size variations with content
  const html2 = `
    <h1>Page Size Variations</h1>
    <p>This document tests how content flows on different page sizes.</p>

    <h2>Content Layout Test</h2>
    <p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>
    <p>Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.</p>
    <p>Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.</p>
    <p>Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>

    <h2>Table on Different Page Sizes</h2>
    <table style="border-collapse: collapse; width: 100%;">
      <tr>
        <th style="border: 1px solid #000; padding: 8px;">Column 1</th>
        <th style="border: 1px solid #000; padding: 8px;">Column 2</th>
        <th style="border: 1px solid #000; padding: 8px;">Column 3</th>
      </tr>
      <tr>
        <td style="border: 1px solid #000; padding: 8px;">Data 1</td>
        <td style="border: 1px solid #000; padding: 8px;">Data 2</td>
        <td style="border: 1px solid #000; padding: 8px;">Data 3</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000; padding: 8px;">Longer content that should wrap</td>
        <td style="border: 1px solid #000; padding: 8px;">More content</td>
        <td style="border: 1px solid #000; padding: 8px;">Even more content</td>
      </tr>
    </table>

    <h2>List Layout</h2>
    <ul>
      <li>First item in the list</li>
      <li>Second item with longer text that should demonstrate how lists flow on different page sizes</li>
      <li>Third item</li>
      <li>Fourth item with even more text to ensure proper wrapping and layout</li>
    </ul>
  `;

  const pageSizeSettings = [
    { name: 'A4-Portrait', pageSize: 'A4', orientation: 'portrait' },
    { name: 'Letter-Portrait', pageSize: 'Letter', orientation: 'portrait' },
    { name: 'A4-Landscape', pageSize: { width: 11.69, height: 8.27 }, orientation: 'landscape' },
    { name: 'Custom-6x9', pageSize: { width: 6, height: 9 }, orientation: 'custom' }
  ];

  for (let i = 0; i < pageSizeSettings.length; i++) {
    const pageSetting = pageSizeSettings[i];
    const converter = new HtmlToDocx({
      pageSize: pageSetting.pageSize,
      fontSize: 11,
      fontFamily: 'Calibri',
      lineHeight: 1.15
    });

    const outputPath = path.join(__dirname, `test-page-size-${pageSetting.name}.docx`);

    try {
      await converter.convertHtmlToDocxFile(html2, outputPath);
      console.log(`✓ Page size (${pageSetting.name}) test completed successfully`);
      console.log(`  Output: ${outputPath}`);
    } catch (error) {
      console.error(`✗ Page size (${pageSetting.name}) test failed:`, error.message);
    }
  }

  // Test 3: Font and page size combinations
  const html3 = `
    <h1>Font and Page Size Combinations</h1>
    <p>This document combines different font settings with page sizes.</p>
    <p>The layout should adapt to both the font choices and page dimensions.</p>

    <h2>Content Sections</h2>
    <p><strong>Section 1:</strong> Normal paragraph text.</p>
    <p><em>Section 2:</em> <span style="color: #FF0000;">Colored text</span> with emphasis.</p>
    <p><u>Section 3:</u> Underlined text for emphasis.</p>

    <h2>Sample Table</h2>
    <table style="border-collapse: collapse;">
      <tr>
        <th style="border: 1px solid #000; padding: 5px;">Feature</th>
        <th style="border: 1px solid #000; padding: 5px;">Description</th>
      </tr>
      <tr>
        <td style="border: 1px solid #000; padding: 5px;">Fonts</td>
        <td style="border: 1px solid #000; padding: 5px;">Customizable font families and sizes</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000; padding: 5px;">Layout</td>
        <td style="border: 1px solid #000; padding: 5px;">Responsive to page dimensions</td>
      </tr>
    </table>
  `;

  const combinations = [
    { name: 'A4-LargeFont', pageSize: 'A4', fontSize: 14, fontFamily: 'Arial' },
    { name: 'Letter-SmallFont', pageSize: 'Letter', fontSize: 9, fontFamily: 'Times New Roman' },
    { name: 'Custom-MediumFont', pageSize: { width: 7, height: 10 }, fontSize: 12, fontFamily: 'Calibri' }
  ];

  for (let i = 0; i < combinations.length; i++) {
    const combo = combinations[i];
    const converter = new HtmlToDocx({
      pageSize: combo.pageSize,
      fontSize: combo.fontSize,
      fontFamily: combo.fontFamily,
      lineHeight: 1.15
    });

    const outputPath = path.join(__dirname, `test-combo-${combo.name}.docx`);

    try {
      await converter.convertHtmlToDocxFile(html3, outputPath);
      console.log(`✓ Font/page combination (${combo.name}) test completed successfully`);
      console.log(`  Output: ${outputPath}`);
    } catch (error) {
      console.error(`✗ Font/page combination (${combo.name}) test failed:`, error.message);
    }
  }

  return true;
}

// Export the test function
export default testPageSizesAndFonts;

// Run the test
testPageSizesAndFonts().catch(console.error);
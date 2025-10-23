import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test basic text formatting features
 * Covers: bold, italic, underline, strikethrough, colors, alignment
 */
async function testBasicTextFormatting() {
  console.log('Testing basic text formatting...');

  const html = `
    <h1>Basic Text Formatting Test</h1>

    <h2>Text Alignment</h2>
    <p style="text-align: left;">This text is left-aligned.</p>
    <p style="text-align: center;">This text is center-aligned.</p>
    <p style="text-align: right;">This text is right-aligned.</p>
    <p style="text-align: justify;">This text is justified. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>

    <h2>Font Styling</h2>
    <p>This is <strong>bold text</strong>.</p>
    <p>This is <em>italic text</em>.</p>
    <p>This is <u>underlined text</u>.</p>
    <p>This is <strike>strikethrough text</strike>.</p>
    <p>This is <strong><em><u>bold, italic, and underlined</u></em></strong> text.</p>

    <h2>Text Colors</h2>
    <p>This is <span style="color: #FF0000;">red text</span>.</p>
    <p>This is <span style="color: #00FF00;">green text</span>.</p>
    <p>This is <span style="color: #0000FF;">blue text</span>.</p>
    <p>This is <span style="color: #FFFF00;">yellow text</span>.</p>
    <p>This is <span style="color: #FF00FF;">magenta text</span>.</p>
    <p>This is <span style="color: #00FFFF;">cyan text</span>.</p>

    <h2>Background Colors</h2>
    <p>This is <span style="background-color: #FFFF00;">text with yellow background</span>.</p>
    <p>This is <span style="background-color: #FF0000; color: #FFFFFF;">white text on red background</span>.</p>

    <h2>Font Sizes</h2>
    <p>This is <span style="font-size: 12px;">12px text</span>.</p>
    <p>This is <span style="font-size: 16px;">16px text</span>.</p>
    <p>This is <span style="font-size: 20px;">20px text</span>.</p>

    <h2>Font Families</h2>
    <p>This is <span style="font-family: Arial;">Arial font</span>.</p>
    <p>This is <span style="font-family: 'Times New Roman';">Times New Roman font</span>.</p>
    <p>This is <span style="font-family: Calibri;">Calibri font</span>.</p>

    <h2>Line Breaks and Spacing</h2>
    <p>Line 1<br>Line 2<br>Line 3</p>
    <p>Normal paragraph.</p>
    <p><br></p>
    <p>Paragraph after empty line.</p>

    <h2>Combined Formatting</h2>
    <p style="text-align: center; font-size: 18px; color: #0000FF;">
      <strong>Centered, large, blue, bold text</strong>
    </p>
  `;

  const converter = new HtmlToDocx({
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath = path.join(__dirname, 'test-basic-formatting.docx');

  try {
    await converter.convertHtmlToDocxFile(html, outputPath);
    console.log('✓ Basic text formatting test completed successfully');
    console.log(`  Output: ${outputPath}`);
    return true;
  } catch (error) {
    console.error('✗ Basic text formatting test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testBasicTextFormatting;

// Run the test
testBasicTextFormatting().catch(console.error);
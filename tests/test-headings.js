import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test headings (H1-H6) with and without custom replacements
 */
async function testHeadings() {
  console.log('Testing headings...');

  // Test 1: Standard headings
  const html1 = `
    <h1>Heading 1 - Main Title</h1>
    <p>This is content under H1.</p>

    <h2>Heading 2 - Section Title</h2>
    <p>This is content under H2.</p>

    <h3>Heading 3 - Subsection Title</h3>
    <p>This is content under H3.</p>

    <h4>Heading 4 - Sub-subsection Title</h4>
    <p>This is content under H4.</p>

    <h5>Heading 5 - Minor Title</h5>
    <p>This is content under H5.</p>

    <h6>Heading 6 - Smallest Title</h6>
    <p>This is content under H6.</p>
  `;

  const converter1 = new HtmlToDocx();
  const outputPath1 = path.join(__dirname, 'test-headings-standard.docx');

  try {
    await converter1.convertHtmlToDocxFile(html1, outputPath1);
    console.log('✓ Standard headings test completed successfully');
    console.log(`  Output: ${outputPath1}`);
  } catch (error) {
    console.error('✗ Standard headings test failed:', error.message);
  }

  // Test 2: Headings with custom replacements
  const html2 = `
    <h1>Custom Styled Heading 1</h1>
    <p>This heading should have a custom black background with white text.</p>

    <h2>Custom Styled Heading 2</h2>
    <p>This heading should have a custom red background with white text.</p>

    <h3>Custom Styled Heading 3</h3>
    <p>This heading should have a custom blue background with white text.</p>
  `;

  const converter2 = new HtmlToDocx({
    headingReplacements: [
      // H1 replacement - black background
      `<div data-h1 class="textbox" style="border: 1px solid #000000ff; border-radius: 5px; padding: 0px 5px 0px 5px; background-color: #000000ff; width:100%;">
<p data-no-spacing style="color: #ffffffff; margin: 0; font-size: 21px; font-weight: bold;">HEADING_TEXT</p>
</div>`,
      // H2 replacement - red background
      `<div data-h2 class="textbox" style="border: 1px solid #ff0000ff; border-radius: 5px; padding: 0px 5px 0px 5px; background-color: #ff0000ff; width:100%;">
<p data-no-spacing style="color: #ffffffff; margin: 0; font-size: 19px; font-weight: bold;">HEADING_TEXT</p>
</div>`,
      // H3 replacement - blue background
      `<div data-h3 class="textbox" style="border: 1px solid #0000ffff; border-radius: 5px; padding: 0px 5px 0px 5px; background-color: #0000ffff; width:100%;">
<p data-no-spacing style="color: #ffffffff; margin: 0; font-size: 16px; font-weight: bold;">HEADING_TEXT</p>
</div>`
    ]
  });

  const outputPath2 = path.join(__dirname, 'test-headings-custom.docx');

  try {
    await converter2.convertHtmlToDocxFile(html2, outputPath2);
    console.log('✓ Custom headings test completed successfully');
    console.log(`  Output: ${outputPath2}`);
  } catch (error) {
    console.error('✗ Custom headings test failed:', error.message);
  }

  // Test 3: Headings with inline formatting
  const html3 = `
    <h1>Heading with <em>Italic</em> and <strong>Bold</strong> Text</h1>
    <p>Normal paragraph.</p>

    <h2>Heading with <span style="color: #FF0000;">Red</span> Text</h2>
    <p>Normal paragraph.</p>

    <h3>Heading with <u>Underlined</u> Text</h3>
    <p>Normal paragraph.</p>
  `;

  const converter3 = new HtmlToDocx();
  const outputPath3 = path.join(__dirname, 'test-headings-inline.docx');

  try {
    await converter3.convertHtmlToDocxFile(html3, outputPath3);
    console.log('✓ Inline formatted headings test completed successfully');
    console.log(`  Output: ${outputPath3}`);
    return true;
  } catch (error) {
    console.error('✗ Inline formatted headings test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testHeadings;

// Run the test
testHeadings().catch(console.error);
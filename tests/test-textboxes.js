import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test text boxes (divs with background, border, padding)
 */
async function testTextBoxes() {
  console.log('Testing text boxes...');

  const html = `
    <h1>Text Box Test</h1>

    <h2>Basic Text Boxes</h2>
    <div style="background-color: #FFFF00; padding: 10px; border: 1px solid #000;">
      <p>This is a basic text box with yellow background, padding, and border.</p>
    </div>

    <h2>Text Boxes with Different Backgrounds</h2>
    <div style="background-color: #FF0000; color: #FFFFFF; padding: 15px; margin: 10px 0;">
      <p>Red background with white text.</p>
      <p>Multiple paragraphs in the same text box.</p>
    </div>

    <div style="background-color: #00FF00; padding: 10px; border-radius: 5px;">
      <p>Green background with rounded corners (border-radius).</p>
    </div>

    <div style="background-color: #0000FF; color: #FFFFFF; padding: 20px;">
      <p>Blue background with lots of padding.</p>
    </div>

    <h2>Text Boxes with Borders</h2>
    <div style="border: 2px solid #FF0000; padding: 10px; margin: 10px 0;">
      <p>Thick red border.</p>
    </div>

    <div style="border: 1px dashed #00FF00; padding: 10px; margin: 10px 0;">
      <p>Dashed green border.</p>
    </div>

    <div style="border: 3px double #0000FF; padding: 10px; margin: 10px 0;">
      <p>Double blue border.</p>
    </div>

    <h2>Text Boxes with Width</h2>
    <div style="background-color: #F0F0F0; padding: 10px; width: 50%; margin: 10px auto;">
      <p>50% width text box, centered.</p>
    </div>

    <div style="background-color: #E0E0E0; padding: 10px; width: 300px; margin: 10px 0;">
      <p>Fixed width (300px) text box.</p>
    </div>

    <h2>Complex Text Boxes</h2>
    <div style="background-color: #FFFACD; border: 2px solid #DAA520; border-radius: 10px; padding: 15px; margin: 15px 0; width: 80%;">
      <h3 style="margin-top: 0;">Nested Content</h3>
      <p>This text box contains:</p>
      <ul>
        <li>Headings</li>
        <li>Lists</li>
        <li><strong>Formatted text</strong></li>
      </ul>
      <p>And multiple paragraphs with <em>emphasis</em> and <span style="color: #FF0000;">colors</span>.</p>
    </div>

    <h2>Text Boxes vs Regular Divs</h2>
    <p>Regular div (no background/border/padding - should render as normal text):</p>
    <div>
      <p>This is regular content in a div without text box styling.</p>
    </div>

    <p>Text box div (has background - should render as text box):</p>
    <div style="background-color: #FFE4E1; padding: 5px;">
      <p>This is content in a text box div.</p>
    </div>

    <h2>Text Boxes with Images</h2>
    <div style="background-color: #F5F5DC; border: 1px solid #8B4513; padding: 10px; margin: 10px 0;">
      <p>Text box containing an image:</p>
      <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==" alt="Small image in text box" style="width: 50px; height: 50px; float: right;"/>
      <p>This text box has an image floated to the right. The text should flow around it.</p>
      <p>More content in the text box to demonstrate the layout.</p>
    </div>

    <h2>Edge Cases</h2>
    <div style="background-color: #FFFFFF; border: 1px solid #000000; padding: 0px;">
      <p>Text box with zero padding.</p>
    </div>

    <div style="background-color: #000000; color: #FFFFFF; padding: 10px;">
      <p>Empty text box content:</p>
    </div>

    <div style="background-color: #CCCCCC; padding: 10px;">
      <p>Text box with only background color (no border).</p>
    </div>
  `;

  const converter = new HtmlToDocx({
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath = path.join(__dirname, 'test-textboxes.docx');

  try {
    await converter.convertHtmlToDocxFile(html, outputPath);
    console.log('✓ Text boxes test completed successfully');
    console.log(`  Output: ${outputPath}`);
    return true;
  } catch (error) {
    console.error('✗ Text boxes test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testTextBoxes;

// Run the test
testTextBoxes().catch(console.error);
import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test border-radius property for text boxes
 */
async function testBorderRadius() {
  console.log('Testing border-radius...');

  const html = `
    <h1>Border Radius Test</h1>

    <h2>Basic Border Radius Values</h2>
    <div style="background-color: #FFFF00; padding: 10px; border-radius: 0px;">
      <p>No border radius (square corners).</p>
    </div>

    <div style="background-color: #FF0000; color: #FFFFFF; padding: 10px; border-radius: 5px;">
      <p>5px border radius.</p>
    </div>

    <div style="background-color: #00FF00; padding: 10px; border-radius: 10px;">
      <p>10px border radius.</p>
    </div>

    <div style="background-color: #0000FF; color: #FFFFFF; padding: 10px; border-radius: 20px;">
      <p>20px border radius.</p>
    </div>

    <div style="background-color: #FF00FF; padding: 10px; border-radius: 50px;">
      <p>50px border radius (very rounded).</p>
    </div>

    <h2>Border Radius with Borders</h2>
    <div style="background-color: #FFFFFF; border: 2px solid #000000; padding: 10px; border-radius: 5px;">
      <p>White background with black border and 5px radius.</p>
    </div>

    <div style="background-color: #FFE4E1; border: 1px solid #FF0000; padding: 15px; border-radius: 10px;">
      <p>Misty rose background with red border and 10px radius.</p>
    </div>

    <div style="background-color: #E0FFFF; border: 3px solid #008080; padding: 10px; border-radius: 15px;">
      <p>Light cyan background with teal border and 15px radius.</p>
    </div>

    <h2>Border Radius with Different Units</h2>
    <div style="background-color: #FFFACD; padding: 10px; border-radius: 0.5em;">
      <p>0.5em border radius (should be treated as px).</p>
    </div>

    <div style="background-color: #F0E68C; padding: 10px; border-radius: 1rem;">
      <p>1rem border radius (should be treated as px).</p>
    </div>

    <h2>Border Radius Limits</h2>
    <div style="background-color: #DDA0DD; padding: 10px; border-radius: 100px; width: 200px; height: 50px;">
      <p>Very large radius (100px) on small element - should be clamped.</p>
    </div>

    <div style="background-color: #98FB98; padding: 10px; border-radius: 25%; width: 200px;">
      <p>Percentage radius (25%) - should be treated as px.</p>
    </div>

    <h2>Border Radius with Width Constraints</h2>
    <div style="background-color: #F5DEB3; padding: 10px; border-radius: 10px; width: 50%;">
      <p>50% width with 10px radius.</p>
    </div>

    <div style="background-color: #DEB887; padding: 10px; border-radius: 15px; width: 300px;">
      <p>Fixed width (300px) with 15px radius.</p>
    </div>

    <div style="background-color: #D2B48C; padding: 10px; border-radius: 20px; width: 100%;">
      <p>Full width (100%) with 20px radius.</p>
    </div>

    <h2>Complex Border Radius Combinations</h2>
    <div style="background-color: #FFB6C1; border: 2px solid #DC143C; padding: 20px; border-radius: 25px; margin: 15px 0; width: 80%;">
      <h3 style="margin-top: 0;">Complex Example</h3>
      <p>This text box combines:</p>
      <ul>
        <li>Large border radius (25px)</li>
        <li>Thick border (2px solid)</li>
        <li>Large padding (20px)</li>
        <li>Width constraint (80%)</li>
        <li>Margin spacing</li>
      </ul>
      <p>Multiple paragraphs with <strong>bold text</strong> and <em>italic text</em>.</p>
    </div>

    <h2>Border Radius Edge Cases</h2>
    <div style="background-color: #FFFFE0; padding: 5px; border-radius: 2px;">
      <p>Very small radius (2px) with minimal padding.</p>
    </div>

    <div style="background-color: #F0FFF0; padding: 30px; border-radius: 30px;">
      <p>Large padding (30px) matching large radius (30px).</p>
    </div>

    <div style="background-color: #FAF0E6; border: 1px solid #A0522D; padding: 0px; border-radius: 5px;">
      <p>Zero padding with border and radius.</p>
    </div>
  `;

  const converter = new HtmlToDocx({
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath = path.join(__dirname, 'test-border-radius.docx');

  try {
    await converter.convertHtmlToDocxFile(html, outputPath);
    console.log('✓ Border radius test completed successfully');
    console.log(`  Output: ${outputPath}`);
    return true;
  } catch (error) {
    console.error('✗ Border radius test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testBorderRadius;

// Run the test
testBorderRadius().catch(console.error);
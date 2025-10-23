import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test tables with styling
 */
async function testTables() {
  console.log('Testing tables...');

  const html = `
    <h1>Table Formatting Test</h1>

    <h2>Basic Table</h2>
    <table>
      <tr>
        <th>Header 1</th>
        <th>Header 2</th>
        <th>Header 3</th>
      </tr>
      <tr>
        <td>Data 1</td>
        <td>Data 2</td>
        <td>Data 3</td>
      </tr>
      <tr>
        <td>Data 4</td>
        <td>Data 5</td>
        <td>Data 6</td>
      </tr>
    </table>

    <h2>Table with Borders and Styling</h2>
    <table style="border-collapse: collapse; width: 100%;">
      <tr style="background-color: #f0f0f0;">
        <th style="border: 1px solid #000000; padding: 8px; text-align: left;">Name</th>
        <th style="border: 1px solid #000000; padding: 8px; text-align: left;">Age</th>
        <th style="border: 1px solid #000000; padding: 8px; text-align: left;">City</th>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;">John Doe</td>
        <td style="border: 1px solid #000000; padding: 8px;">30</td>
        <td style="border: 1px solid #000000; padding: 8px;">New York</td>
      </tr>
      <tr style="background-color: #f9f9f9;">
        <td style="border: 1px solid #000000; padding: 8px;">Jane Smith</td>
        <td style="border: 1px solid #000000; padding: 8px;">25</td>
        <td style="border: 1px solid #000000; padding: 8px;">Los Angeles</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;">Bob Johnson</td>
        <td style="border: 1px solid #000000; padding: 8px;">35</td>
        <td style="border: 1px solid #000000; padding: 8px;">Chicago</td>
      </tr>
    </table>

    <h2>Table with Cell Formatting</h2>
    <table style="border-collapse: collapse;">
      <tr>
        <th style="border: 1px solid #000000; padding: 8px;">Type</th>
        <th style="border: 1px solid #000000; padding: 8px;">Example</th>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;"><strong>Bold Text</strong></td>
        <td style="border: 1px solid #000000; padding: 8px;"><strong>This is bold</strong></td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;"><em>Italic Text</em></td>
        <td style="border: 1px solid #000000; padding: 8px;"><em>This is italic</em></td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;"><span style="color: #FF0000;">Colored Text</span></td>
        <td style="border: 1px solid #000000; padding: 8px;"><span style="color: #FF0000;">This is red</span></td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;">Aligned Right</td>
        <td style="border: 1px solid #000000; padding: 8px; text-align: right;">Right aligned content</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;">Centered</td>
        <td style="border: 1px solid #000000; padding: 8px; text-align: center;">Centered content</td>
      </tr>
    </table>

    <h2>Complex Table with Rowspan/Colspan</h2>
    <p>Note: Rowspan and colspan are not fully supported in DOCX tables, but basic structure is preserved.</p>
    <table style="border-collapse: collapse;">
      <tr>
        <th style="border: 1px solid #000000; padding: 8px;">A</th>
        <th style="border: 1px solid #000000; padding: 8px;">B</th>
        <th style="border: 1px solid #000000; padding: 8px;">C</th>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;">1</td>
        <td style="border: 1px solid #000000; padding: 8px;">2</td>
        <td style="border: 1px solid #000000; padding: 8px;">3</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;">4</td>
        <td style="border: 1px solid #000000; padding: 8px;">5</td>
        <td style="border: 1px solid #000000; padding: 8px;">6</td>
      </tr>
    </table>

    <h2>Table with Long Content</h2>
    <table style="border-collapse: collapse;">
      <tr>
        <th style="border: 1px solid #000000; padding: 8px;">Column 1</th>
        <th style="border: 1px solid #000000; padding: 8px;">Column 2</th>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;">This is a long cell content that should wrap properly within the table cell. Lorem ipsum dolor sit amet, consectetur adipiscing elit.</td>
        <td style="border: 1px solid #000000; padding: 8px;">Another long cell with lots of text that demonstrates text wrapping in table cells. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</td>
      </tr>
    </table>

    <h2>Empty Cells and Edge Cases</h2>
    <table style="border-collapse: collapse;">
      <tr>
        <th style="border: 1px solid #000000; padding: 8px;">Header 1</th>
        <th style="border: 1px solid #000000; padding: 8px;">Header 2</th>
        <th style="border: 1px solid #000000; padding: 8px;">Header 3</th>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;">Normal</td>
        <td style="border: 1px solid #000000; padding: 8px;"></td>
        <td style="border: 1px solid #000000; padding: 8px;">Normal</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000000; padding: 8px;"></td>
        <td style="border: 1px solid #000000; padding: 8px;"></td>
        <td style="border: 1px solid #000000; padding: 8px;"></td>
      </tr>
    </table>
  `;

  const converter = new HtmlToDocx({
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath = path.join(__dirname, 'test-tables.docx');

  try {
    await converter.convertHtmlToDocxFile(html, outputPath);
    console.log('✓ Tables test completed successfully');
    console.log(`  Output: ${outputPath}`);
    return true;
  } catch (error) {
    console.error('✗ Tables test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testTables;

// Run the test
testTables().catch(console.error);
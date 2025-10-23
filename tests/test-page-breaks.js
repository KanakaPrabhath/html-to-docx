import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test page breaks and advanced layout
 */
async function testPageBreaks() {
  console.log('Testing page breaks and advanced layout...');

  const html = `
    <h1>Page Breaks and Advanced Layout Test</h1>

    <h2>Manual Page Breaks</h2>
    <p>This paragraph is on the first page.</p>
    <p>More content on the first page.</p>

    <page-break></page-break>

    <h2>Content After First Page Break</h2>
    <p>This content starts on a new page due to the page break element.</p>
    <p>Page breaks can be inserted anywhere in the document.</p>

    <h2>CSS Page Breaks</h2>
    <p>This paragraph has page-break-after: always.</p>
    <p style="page-break-after: always;">This paragraph forces a page break after it.</p>
    <p>This content appears after the CSS page break.</p>

    <p style="page-break-before: always;">This paragraph has page-break-before: always and appears on a new page.</p>
    <p>Content following the paragraph with page-break-before.</p>

    <h2>Page Breaks in Different Contexts</h2>
    <ul>
      <li>List item 1</li>
      <li>List item 2 <page-break></page-break> List item 2 continued on new page</li>
      <li>List item 3</li>
    </ul>

    <h3>Page Break in Table</h3>
    <table style="border-collapse: collapse;">
      <tr>
        <th style="border: 1px solid #000; padding: 8px;">Column 1</th>
        <th style="border: 1px solid #000; padding: 8px;">Column 2</th>
      </tr>
      <tr>
        <td style="border: 1px solid #000; padding: 8px;">Data 1</td>
        <td style="border: 1px solid #000; padding: 8px;">Data 2</td>
      </tr>
      <tr style="page-break-after: always;">
        <td style="border: 1px solid #000; padding: 8px;">Data 3</td>
        <td style="border: 1px solid #000; padding: 8px;">Data 4</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000; padding: 8px;">Data 5</td>
        <td style="border: 1px solid #000; padding: 8px;">Data 6</td>
      </tr>
    </table>

    <h2>Page Breaks with Headings</h2>
    <h3 style="page-break-before: always;">Heading with Page Break Before</h3>
    <p>Content under heading that starts on new page.</p>

    <h3>Normal Heading</h3>
    <p>Normal content.</p>

    <h3 style="page-break-after: always;">Heading with Page Break After</h3>
    <p>This content follows the heading but the next heading will be on a new page.</p>

    <h3>Next Heading (should be on new page)</h3>
    <p>This heading should appear on a new page.</p>

    <h2>Page Breaks in Text Boxes</h2>
    <div style="background-color: #FFFFE0; padding: 10px; border: 1px solid #000; page-break-after: always;">
      <p>This text box has a page break after it.</p>
      <p>The next content will be on a new page.</p>
    </div>

    <p>This content appears after the text box page break.</p>

    <h2>Multiple Page Breaks</h2>
    <p>Content before multiple breaks.</p>
    <page-break></page-break>
    <page-break></page-break>
    <p>Content after multiple consecutive page breaks.</p>

    <h2>Page Breaks with Images</h2>
    <p>Image before page break:</p>
    <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==" alt="Small image" style="width: 50px; height: 50px;"/>
    <page-break></page-break>
    <p>Content after image and page break.</p>

    <h2>Edge Cases</h2>
    <p>Page break at the beginning:</p>
    <page-break></page-break>
    <p>Content after initial page break.</p>

    <p>Page break at the end:</p>
    <page-break></page-break>

    <h2>Final Section</h2>
    <p>This is the final content of the document.</p>
    <p>No more page breaks follow.</p>
  `;

  const converter = new HtmlToDocx({
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15,
    enablePageNumbers: true,
    pageNumberAlignment: 'center'
  });

  const outputPath = path.join(__dirname, 'test-page-breaks.docx');

  try {
    await converter.convertHtmlToDocxFile(html, outputPath);
    console.log('✓ Page breaks and advanced layout test completed successfully');
    console.log(`  Output: ${outputPath}`);
    console.log('  Note: Page breaks may not be visible in all PDF viewers when converted from DOCX');
    return true;
  } catch (error) {
    console.error('✗ Page breaks test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testPageBreaks;

// Run the test
testPageBreaks().catch(console.error);
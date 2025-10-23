import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test lists (ordered and unordered, nested)
 */
async function testLists() {
  console.log('Testing lists...');

  const html = `
    <h1>List Formatting Test</h1>

    <h2>Unordered Lists</h2>
    <ul>
      <li>First item</li>
      <li>Second item</li>
      <li>Third item</li>
    </ul>

    <h2>Ordered Lists</h2>
    <ol>
      <li>First numbered item</li>
      <li>Second numbered item</li>
      <li>Third numbered item</li>
    </ol>

    <h2>Nested Lists</h2>
    <ul>
      <li>Top level item 1
        <ul>
          <li>Nested item 1.1</li>
          <li>Nested item 1.2</li>
        </ul>
      </li>
      <li>Top level item 2
        <ul>
          <li>Nested item 2.1</li>
          <li>Nested item 2.2
            <ul>
              <li>Deep nested item 2.2.1</li>
              <li>Deep nested item 2.2.2</li>
            </ul>
          </li>
        </ul>
      </li>
    </ul>

    <h2>Mixed Ordered and Unordered Lists</h2>
    <ol>
      <li>Numbered item 1
        <ul>
          <li>Bullet under number 1</li>
          <li>Bullet under number 2</li>
        </ul>
      </li>
      <li>Numbered item 2
        <ol>
          <li>Sub-number 2.1</li>
          <li>Sub-number 2.2</li>
        </ol>
      </li>
    </ol>

    <h2>Lists with Formatting</h2>
    <ul>
      <li><strong>Bold item</strong></li>
      <li><em>Italic item</em></li>
      <li><u>Underlined item</u></li>
      <li><span style="color: #FF0000;">Red item</span></li>
    </ul>

    <h2>Lists with Multiple Paragraphs</h2>
    <ul>
      <li>
        <p>First paragraph in list item.</p>
        <p>Second paragraph in same list item.</p>
      </li>
      <li>
        <p>Another item with multiple paragraphs.</p>
        <p>This is the second paragraph.</p>
        <p>And a third paragraph.</p>
      </li>
    </ul>

    <h2>Empty and Edge Case Lists</h2>
    <ul>
      <li>Normal item</li>
      <li></li>
      <li>Item after empty</li>
      <li>   </li>
      <li>Last item</li>
    </ul>

    <h2>Very Long List Items</h2>
    <ul>
      <li>This is a very long list item that should wrap to multiple lines in the document. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.</li>
      <li>Another long item with lots of text that should demonstrate proper text wrapping and formatting within list items. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.</li>
    </ul>
  `;

  const converter = new HtmlToDocx({
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath = path.join(__dirname, 'test-lists.docx');

  try {
    await converter.convertHtmlToDocxFile(html, outputPath);
    console.log('✓ Lists test completed successfully');
    console.log(`  Output: ${outputPath}`);
    return true;
  } catch (error) {
    console.error('✗ Lists test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testLists;

// Run the test
testLists().catch(console.error);
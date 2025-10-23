import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test edge cases and error handling
 */
async function testEdgeCases() {
  console.log('Testing edge cases and error handling...');

  // Test 1: Empty and minimal content
  const html1 = `
    <h1>Empty Content Test</h1>
  `;

  const converter1 = new HtmlToDocx();
  const outputPath1 = path.join(__dirname, 'test-edge-empty.docx');

  try {
    await converter1.convertHtmlToDocxFile(html1, outputPath1);
    console.log('✓ Empty content test completed successfully');
    console.log(`  Output: ${outputPath1}`);
  } catch (error) {
    console.error('✗ Empty content test failed:', error.message);
  }

  // Test 2: Very large content
  let largeHtml = '<h1>Large Content Test</h1>';
  for (let i = 0; i < 100; i++) {
    largeHtml += `<h2>Section ${i + 1}</h2>`;
    largeHtml += `<p>This is paragraph ${i + 1}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>`;
    if (i % 10 === 0) {
      largeHtml += '<page-break></page-break>';
    }
  }

  const converter2 = new HtmlToDocx();
  const outputPath2 = path.join(__dirname, 'test-edge-large.docx');

  try {
    await converter2.convertHtmlToDocxFile(largeHtml, outputPath2);
    console.log('✓ Large content test completed successfully');
    console.log(`  Output: ${outputPath2}`);
  } catch (error) {
    console.error('✗ Large content test failed:', error.message);
  }

  // Test 3: Malformed HTML
  const html3 = `
    <h1>Malformed HTML Test</h1>
    <p>Unclosed paragraph
    <strong>Unclosed strong
    <em>Unclosed em
    <p>Another paragraph</p>
    <br>
    <br/>
    <br />
    <img src="test.jpg" alt="test">
    <div>
      <span>Nested content
    </div>
    <ul>
      <li>List item 1
      <li>List item 2
    </ul>
  `;

  const converter3 = new HtmlToDocx();
  const outputPath3 = path.join(__dirname, 'test-edge-malformed.docx');

  try {
    await converter3.convertHtmlToDocxFile(html3, outputPath3);
    console.log('✓ Malformed HTML test completed successfully');
    console.log(`  Output: ${outputPath3}`);
  } catch (error) {
    console.error('✗ Malformed HTML test failed:', error.message);
  }

  // Test 4: Special characters and Unicode
  const html4 = `
    <h1>Special Characters Test</h1>
    <p>Basic ASCII: A B C 1 2 3 ! @ # $ % ^ & * ( )</p>
    <p>Quotes: "double quotes" 'single quotes' \`backticks\`</p>
    <p>Math symbols: ± × ÷ √ ∞ ∑ ∏ ∆ ∇</p>
    <p>Greek: α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ τ υ φ χ ψ ω</p>
    <p>Arrows: ← ↑ → ↓ ↔ ↕ ⇐ ⇑ ⇒ ⇓ ⇔ ⇕</p>
    <p>Currency: $ € £ ¥ ¢ ₹ ₽ ₩ ₦</p>
    <p>Fractions: ½ ⅓ ¼ ¾ ⅛ ⅜ ⅝ ⅞</p>
    <p>Emojis: 😀 👍 ❤️ 🔥 ⭐ 🎉</p>
    <p>Accented: café naïve résumé</p>
  `;

  const converter4 = new HtmlToDocx();
  const outputPath4 = path.join(__dirname, 'test-edge-unicode.docx');

  try {
    await converter4.convertHtmlToDocxFile(html4, outputPath4);
    console.log('✓ Special characters test completed successfully');
    console.log(`  Output: ${outputPath4}`);
  } catch (error) {
    console.error('✗ Special characters test failed:', error.message);
  }

  // Test 5: Buffer output test
  const html5 = `<p>This is a test for buffer output.</p>`;

  const converter5 = new HtmlToDocx();

  try {
    const buffer = await converter5.convertHtmlToDocx(html5);
    console.log('✓ Buffer output test completed successfully');
    console.log(`  Buffer size: ${buffer.length} bytes`);
  } catch (error) {
    console.error('✗ Buffer output test failed:', error.message);
  }

  // Test 6: Invalid options handling
  const html6 = `<h1>Invalid Options Test</h1><p>This should handle invalid options gracefully.</p>`;

  const converter6 = new HtmlToDocx({
    fontSize: 'invalid', // Should be a number
    fontFamily: 123, // Should be a string
    lineHeight: 'invalid', // Should be a number
    pageSize: 'InvalidSize', // Invalid page size
    marginTop: 'invalid' // Should be a number
  });

  const outputPath6 = path.join(__dirname, 'test-edge-invalid-options.docx');

  try {
    await converter6.convertHtmlToDocxFile(html6, outputPath6);
    console.log('✓ Invalid options test completed successfully');
    console.log(`  Output: ${outputPath6}`);
  } catch (error) {
    console.error('✗ Invalid options test failed:', error.message);
  }

  // Test 7: Nested complex structures
  const html7 = `
    <h1>Complex Nested Structures</h1>
    <div style="background-color: #f0f0f0; padding: 10px;">
      <h2>Heading in text box</h2>
      <ul>
        <li>List item with <strong>bold</strong> and <em>italic</em>
          <ul>
            <li>Nested list item</li>
            <li>Another nested item with <span style="color: #FF0000;">color</span></li>
          </ul>
        </li>
        <li>Second top-level item
          <ol>
            <li>Ordered sub-item 1</li>
            <li>Ordered sub-item 2</li>
          </ol>
        </li>
      </ul>
      <table style="border-collapse: collapse; margin: 10px 0;">
        <tr>
          <th style="border: 1px solid #000; padding: 5px;">Table in</th>
          <th style="border: 1px solid #000; padding: 5px;">Text Box</th>
        </tr>
        <tr>
          <td style="border: 1px solid #000; padding: 5px;">Cell 1</td>
          <td style="border: 1px solid #000; padding: 5px;">Cell 2</td>
        </tr>
      </table>
    </div>
  `;

  const converter7 = new HtmlToDocx();
  const outputPath7 = path.join(__dirname, 'test-edge-complex-nested.docx');

  try {
    await converter7.convertHtmlToDocxFile(html7, outputPath7);
    console.log('✓ Complex nested structures test completed successfully');
    console.log(`  Output: ${outputPath7}`);
  } catch (error) {
    console.error('✗ Complex nested structures test failed:', error.message);
  }

  // Test 8: Performance test with many elements
  let html8 = '<h1>Performance Test</h1>';
  for (let i = 0; i < 500; i++) {
    html8 += `<p>Paragraph ${i + 1} with some content to test performance.</p>`;
  }

  const converter8 = new HtmlToDocx();
  const outputPath8 = path.join(__dirname, 'test-edge-performance.docx');

  try {
    const startTime = Date.now();
    await converter8.convertHtmlToDocxFile(html8, outputPath8);
    const endTime = Date.now();
    console.log('✓ Performance test completed successfully');
    console.log(`  Output: ${outputPath8}`);
    console.log(`  Time taken: ${endTime - startTime}ms`);
  } catch (error) {
    console.error('✗ Performance test failed:', error.message);
  }

  return true;
}

// Export the test function
export default testEdgeCases;

// Run the test
testEdgeCases().catch(console.error);
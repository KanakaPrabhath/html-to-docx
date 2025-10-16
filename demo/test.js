import HtmlToDocx from '@kanaka-prabhath/html-to-docx';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function runTests() {
  console.log('Starting HTML to DOCX conversion tests...\n');

  // Test 1: Basic alignment test (from your example)
  console.log('Test 1: Text alignment test');
  const html1 = `<p style="text-align: left;">Test Text </p><p style="text-align: center;">Test Text</p><p style="text-align: right;">Test Text</p><p style="text-align: right;"><br></p>`;
  
  const converter1 = new HtmlToDocx();
  const outputPath1 = path.join(__dirname, 'test-output-alignment.docx');
  
  try {
    await converter1.convertHtmlToDocxFile(html1, outputPath1);
    console.log('✓ Alignment test completed successfully');
    console.log(`  Output: ${outputPath1}\n`);
  } catch (error) {
    console.error('✗ Alignment test failed:', error.message);
  }

  // Test 2: Complex formatting test
  console.log('Test 2: Complex formatting test');
  const html2 = `
    <h1>This is a Heading 1</h1>
    <h2>This is a Heading 2</h2>
    <p style="text-align: left;">This is left-aligned text.</p>
    <p style="text-align: center;"><strong>This is bold and centered</strong></p>
    <p style="text-align: right;"><em>This is italic and right-aligned</em></p>
    <p style="text-align: justify;"><u>This is underlined and justified text that should span multiple lines to demonstrate the justification effect properly.</u></p>
    <p><strong><em><u>Bold, italic, and underlined</u></em></strong></p>
    <p style="color: #FF0000;">Red text</p>
    <p style="background-color: #FFFF00;">Yellow background</p>
    <p style="font-size: 16px;">Larger text (16px)</p>
    <p style="font-family: Arial;">Arial font family</p>
    <p>Line 1<br>Line 2<br>Line 3</p>
    <p><br></p>
    <ul>
      <li>List item 1</li>
      <li>List item 2</li>
      <li>List item 3</li>
    </ul>
  `;
  
  const converter2 = new HtmlToDocx({
    fontSize: 12,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });
  const outputPath2 = path.join(__dirname, 'test-output-complex.docx');
  
  try {
    await converter2.convertHtmlToDocxFile(html2, outputPath2);
    console.log('✓ Complex formatting test completed successfully');
    console.log(`  Output: ${outputPath2}\n`);
  } catch (error) {
    console.error('✗ Complex formatting test failed:', error.message);
  }

  // Test 3: Buffer output test
  console.log('Test 3: Buffer output test');
  const html3 = `<p style="text-align: center;">Test buffer output</p>`;
  
  const converter3 = new HtmlToDocx();
  
  try {
    const buffer = await converter3.convertHtmlToDocx(html3);
    console.log('✓ Buffer output test completed successfully');
    console.log(`  Buffer size: ${buffer.length} bytes\n`);
  } catch (error) {
    console.error('✗ Buffer output test failed:', error.message);
  }

  // Test 4: Edge cases
  console.log('Test 4: Edge cases test');
  const html4 = `
    <p></p>
    <p><br></p>
    <p>   </p>
    <p style="text-align: center;"></p>
    <p>Normal text with <strong>bold</strong> and <em>italic</em> inline.</p>
    <p><span style="color: #0000FF; font-size: 14px;">Blue span with custom size</span></p>
  `;
  
  const converter4 = new HtmlToDocx();
  const outputPath4 = path.join(__dirname, 'test-output-edge-cases.docx');
  
  try {
    await converter4.convertHtmlToDocxFile(html4, outputPath4);
    console.log('✓ Edge cases test completed successfully');
    console.log(`  Output: ${outputPath4}\n`);
  } catch (error) {
    console.error('✗ Edge cases test failed:', error.message);
  }

  // Test 5: HTML cleaning test
  console.log('Test 5: HTML cleaning test');
  const html5 = `
    <p>Normal text</p>
    <script>alert('dangerous');</script>
    <style>body { background: red; }</style>
    <p style="text-align: center; position: absolute; color: red;">Centered text</p>
    <p>Text with invalid char: \x00\x01</p>
    <iframe src="evil.com"></iframe>
  `;
  
  const converter5 = new HtmlToDocx();
  const outputPath5 = path.join(__dirname, 'test-output-cleaning.docx');
  
  try {
    await converter5.convertHtmlToDocxFile(html5, outputPath5);
    console.log('✓ HTML cleaning test completed successfully');
    console.log(`  Output: ${outputPath5}\n`);
  } catch (error) {
    console.error('✗ HTML cleaning test failed:', error.message);
  }
}

// Run tests
runTests().catch(console.error);

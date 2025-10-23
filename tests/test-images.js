import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test images (URL and base64)
 */
async function testImages() {
  console.log('Testing images...');

  // Small 1x1 pixel red PNG in base64 (for testing)
  const smallRedPngBase64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

  const html = `
    <h1>Image Test</h1>

    <h2>Base64 Images</h2>
    <p>This is a small red square (1x1 pixel) embedded as base64:</p>
    <img src="${smallRedPngBase64}" alt="Small red square" style="width: 50px; height: 50px; border: 1px solid #000;"/>

    <p>Another base64 image with different styling:</p>
    <img src="${smallRedPngBase64}" alt="Another small image" style="width: 100px; height: 100px; float: right; margin: 0px 0px 10px 10px;"/>

    <h2>URL Images</h2>
    <p>The following image should be downloaded from a URL (placeholder - will fail gracefully if URL doesn't exist):</p>
    <img src="https://via.placeholder.com/200x100/ff0000/ffffff?text=Test+Image" alt="Placeholder image" style="width: 200px; height: 100px;"/>

    <h2>Images with Alt Text</h2>
    <p>Image with descriptive alt text:</p>
    <img src="${smallRedPngBase64}" alt="A small red test image for demonstration purposes" style="width: 150px; height: 150px; display: block; margin: 10px auto;"/>

    <h2>Images in Different Contexts</h2>
    <p>Image inline with text: <img src="${smallRedPngBase64}" alt="inline" style="width: 20px; height: 20px; vertical-align: middle;"/> This text has an image embedded within it.</p>

    <div style="text-align: center; border: 1px solid #ccc; padding: 10px; margin: 10px 0;">
      <p>Centered image in a bordered box:</p>
      <img src="${smallRedPngBase64}" alt="Centered image" style="width: 80px; height: 80px;"/>
    </div>

    <h2>Large Images</h2>
    <p>A larger version of the same image:</p>
    <img src="${smallRedPngBase64}" alt="Large version" style="width: 300px; height: 300px; display: block; margin: 20px auto;"/>

    <h2>Images in Lists</h2>
    <ul>
      <li>Item with image: <img src="${smallRedPngBase64}" alt="list image" style="width: 30px; height: 30px;"/></li>
      <li>Another item</li>
    </ul>

    <h2>Invalid Images</h2>
    <p>Invalid base64 (should be skipped gracefully):</p>
    <img src="data:image/png;base64,invalid" alt="Invalid base64" style="width: 50px; height: 50px;"/>

    <p>Invalid URL (should be skipped gracefully):</p>
    <img src="invalid-url" alt="Invalid URL" style="width: 50px; height: 50px;"/>

    <p>Missing src attribute (should be skipped):</p>
    <img alt="No src" style="width: 50px; height: 50px;"/>
  `;

  const converter = new HtmlToDocx({
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath = path.join(__dirname, 'test-images.docx');

  try {
    await converter.convertHtmlToDocxFile(html, outputPath);
    console.log('✓ Images test completed successfully');
    console.log(`  Output: ${outputPath}`);
    console.log('  Note: URL images may fail if the URL is not accessible');
    return true;
  } catch (error) {
    console.error('✗ Images test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testImages;

// Run the test
testImages().catch(console.error);
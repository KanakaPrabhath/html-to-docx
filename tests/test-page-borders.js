import HtmlToDocx from '../lib/index.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test page borders and margins
 */
async function testPageBorders() {
  console.log('Testing page borders and margins...');

  // Test 1: Different page border styles
  const html1 = `
    <h1>Page Border Styles Test</h1>
    <p>This document demonstrates different page border styles.</p>
    <p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>
    <p>Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.</p>
    <p>Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.</p>
  `;

  const borderStyles = [
    { style: 'single', color: 'FF0000', size: 4 },
    { style: 'double', color: '00FF00', size: 8 },
    { style: 'thick', color: '0000FF', size: 12 },
    { style: 'dotted', color: 'FF00FF', size: 2 },
    { style: 'dashed', color: 'FFFF00', size: 3 }
  ];

  for (let i = 0; i < borderStyles.length; i++) {
    const border = borderStyles[i];
    const converter = new HtmlToDocx({
      enablePageBorder: true,
      pageBorder: border,
      fontSize: 11,
      fontFamily: 'Calibri',
      lineHeight: 1.15
    });

    const outputPath = path.join(__dirname, `test-page-border-${border.style}.docx`);

    try {
      await converter.convertHtmlToDocxFile(html1, outputPath);
      console.log(`✓ Page border (${border.style}) test completed successfully`);
      console.log(`  Output: ${outputPath}`);
    } catch (error) {
      console.error(`✗ Page border (${border.style}) test failed:`, error.message);
    }
  }

  // Test 1.5: Page border radius variations
  const borderRadii = [
    { radius: 0, name: 'sharp' },
    { radius: 2, name: 'small-radius' },
    { radius: 5, name: 'medium-radius' },
    { radius: 8, name: 'large-radius' },
    { radius: 15, name: 'extra-large-radius' },
    { radius: 25, name: 'extreme-radius' }
  ];

  for (let i = 0; i < borderRadii.length; i++) {
    const borderRadius = borderRadii[i];
    const converter = new HtmlToDocx({
      enablePageBorder: true,
      pageBorder: {
        style: 'single',
        color: '0000FF',
        size: 2,
        radius: borderRadius.radius
      },
      fontSize: 11,
      fontFamily: 'Calibri',
      lineHeight: 1.15
    });

    const outputPath = path.join(__dirname, `test-page-border-radius-${borderRadius.name}.docx`);

    try {
      await converter.convertHtmlToDocxFile(html1, outputPath);
      console.log(`✓ Page border radius (${borderRadius.name}) test completed successfully`);
      console.log(`  Output: ${outputPath}`);
    } catch (error) {
      console.error(`✗ Page border radius (${borderRadius.name}) test failed:`, error.message);
    }
  }

  // Test 1.6: Border radius with different border styles
  const borderRadiusStyles = [
    { style: 'single', radius: 5, name: 'single-rounded' },
    { style: 'double', radius: 5, name: 'double-rounded' },
    { style: 'thick', radius: 5, name: 'thick-rounded' },
    { style: 'dotted', radius: 3, name: 'dotted-rounded' },
    { style: 'dashed', radius: 3, name: 'dashed-rounded' }
  ];

  for (let i = 0; i < borderRadiusStyles.length; i++) {
    const borderConfig = borderRadiusStyles[i];
    const converter = new HtmlToDocx({
      enablePageBorder: true,
      pageBorder: {
        style: borderConfig.style,
        color: 'FF0000',
        size: 4,
        radius: borderConfig.radius
      },
      fontSize: 11,
      fontFamily: 'Calibri',
      lineHeight: 1.15
    });

    const outputPath = path.join(__dirname, `test-page-border-radius-${borderConfig.name}.docx`);

    try {
      await converter.convertHtmlToDocxFile(html1, outputPath);
      console.log(`✓ Page border radius with ${borderConfig.name} style test completed successfully`);
      console.log(`  Output: ${outputPath}`);
    } catch (error) {
      console.error(`✗ Page border radius with ${borderConfig.name} style test failed:`, error.message);
    }
  }

  // Test 1.7: Border radius with different colors and sizes
  const borderRadiusColors = [
    { color: 'FF0000', size: 2, radius: 4, name: 'red-small-rounded' },
    { color: '00FF00', size: 6, radius: 8, name: 'green-medium-rounded' },
    { color: '0000FF', size: 10, radius: 12, name: 'blue-large-rounded' },
    { color: 'FFFF00', size: 1, radius: 2, name: 'yellow-thin-rounded' },
    { color: 'FF00FF', size: 8, radius: 6, name: 'magenta-thick-rounded' }
  ];

  for (let i = 0; i < borderRadiusColors.length; i++) {
    const borderConfig = borderRadiusColors[i];
    const converter = new HtmlToDocx({
      enablePageBorder: true,
      pageBorder: {
        style: 'single',
        color: borderConfig.color,
        size: borderConfig.size,
        radius: borderConfig.radius
      },
      fontSize: 11,
      fontFamily: 'Calibri',
      lineHeight: 1.15
    });

    const outputPath = path.join(__dirname, `test-page-border-radius-${borderConfig.name}.docx`);

    try {
      await converter.convertHtmlToDocxFile(html1, outputPath);
      console.log(`✓ Page border radius (${borderConfig.name}) test completed successfully`);
      console.log(`  Output: ${outputPath}`);
    } catch (error) {
      console.error(`✗ Page border radius (${borderConfig.name}) test failed:`, error.message);
    }
  }

  // Test 2: Page margins
  const html2 = `
    <h1>Page Margins Test</h1>
    <p>This document tests different page margin settings.</p>
    <p>The margins affect the printable area of the document.</p>
    <p>Content should be positioned according to the specified margins.</p>
  `;

  const marginSettings = [
    { name: 'default', top: 1, right: 1, bottom: 1, left: 1 },
    { name: 'narrow', top: 0.5, right: 0.5, bottom: 0.5, left: 0.5 },
    { name: 'wide', top: 2, right: 2, bottom: 2, left: 2 },
    { name: 'asymmetric', top: 1.5, right: 2.5, bottom: 1, left: 1.5 }
  ];

  for (let i = 0; i < marginSettings.length; i++) {
    const margins = marginSettings[i];
    const converter = new HtmlToDocx({
      marginTop: margins.top,
      marginRight: margins.right,
      marginBottom: margins.bottom,
      marginLeft: margins.left,
      fontSize: 11,
      fontFamily: 'Calibri',
      lineHeight: 1.15
    });

    const outputPath = path.join(__dirname, `test-margins-${margins.name}.docx`);

    try {
      await converter.convertHtmlToDocxFile(html2, outputPath);
      console.log(`✓ Page margins (${margins.name}) test completed successfully`);
      console.log(`  Output: ${outputPath}`);
    } catch (error) {
      console.error(`✗ Page margins (${margins.name}) test failed:`, error.message);
    }
  }

  // Test 3: Page size variations
  const html3 = `
    <h1>Page Size Test</h1>
    <p>This document tests different page sizes.</p>
    <p>The content layout should adapt to the page dimensions.</p>
  `;

  const pageSizes = [
    { name: 'A4', size: 'A4' },
    { name: 'Letter', size: 'Letter' },
    { name: 'Legal', size: 'Legal' },
    { name: 'Custom-8x10', size: { width: 8, height: 10 } }
  ];

  for (let i = 0; i < pageSizes.length; i++) {
    const pageSize = pageSizes[i];
    const converter = new HtmlToDocx({
      pageSize: pageSize.size,
      fontSize: 11,
      fontFamily: 'Calibri',
      lineHeight: 1.15
    });

    const outputPath = path.join(__dirname, `test-page-size-${pageSize.name}.docx`);

    try {
      await converter.convertHtmlToDocxFile(html3, outputPath);
      console.log(`✓ Page size (${pageSize.name}) test completed successfully`);
      console.log(`  Output: ${outputPath}`);
    } catch (error) {
      console.error(`✗ Page size (${pageSize.name}) test failed:`, error.message);
    }
  }

  // Test 4: Combined page settings
  const html4 = `
    <h1>Combined Page Settings</h1>
    <p>This document combines page borders, margins, and size.</p>
    <p>All page layout settings are applied together.</p>
    <p>This demonstrates the full page customization capabilities.</p>
  `;

  const converter4 = new HtmlToDocx({
    pageSize: 'A4',
    marginTop: 1.5,
    marginRight: 1.25,
    marginBottom: 1.5,
    marginLeft: 1.25,
    enablePageBorder: true,
    pageBorder: {
      style: 'double',
      color: '000000',
      size: 6
    },
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  const outputPath4 = path.join(__dirname, 'test-combined-page-settings.docx');

  try {
    await converter4.convertHtmlToDocxFile(html4, outputPath4);
    console.log('✓ Combined page settings test completed successfully');
    console.log(`  Output: ${outputPath4}`);
    return true;
  } catch (error) {
    console.error('✗ Combined page settings test failed:', error.message);
    return false;
  }
}

// Export the test function
export default testPageBorders;

// Run the test
testPageBorders().catch(console.error);
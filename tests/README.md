# HTML to DOCX Test Suite

This directory contains comprehensive tests for the HTML to DOCX converter library. The tests cover all major features and edge cases to ensure the library works correctly.

## Test Files

### Core Feature Tests
- **`test-basic-formatting.js`** - Tests text formatting (bold, italic, underline, colors, alignment)
- **`test-headings.js`** - Tests headings (H1-H6) with and without custom replacements
- **`test-lists.js`** - Tests ordered and unordered lists, including nested lists
- **`test-tables.js`** - Tests table creation with styling and borders
- **`test-images.js`** - Tests image embedding (base64 and URL images)
- **`test-textboxes.js`** - Tests text boxes (divs with background/border/padding)

### Layout and Document Structure Tests
- **`test-headers-footers.js`** - Tests headers, footers, and page numbers
- **`test-page-breaks.js`** - Tests page breaks and advanced layout features
- **`test-page-borders.js`** - Tests page borders and margin settings
- **`test-page-sizes-fonts.js`** - Tests different page sizes and font customization

### Edge Cases and Error Handling
- **`test-edge-cases.js`** - Tests edge cases, malformed HTML, and error conditions

### Test Runner
- **`run-tests.js`** - Main test runner that executes all test files

## Running Tests

### Run All Tests
```bash
node tests/run-tests.js
```

### Run Specific Test
```bash
node tests/run-tests.js --run test-basic-formatting.js
```

### List Available Tests
```bash
node tests/run-tests.js --list
```

### Get Help
```bash
node tests/run-tests.js --help
```

## Test Output

Each test generates one or more `.docx` files in the `tests/` directory. You can open these files in Microsoft Word or any DOCX-compatible viewer to verify the results.

### Expected Output Files
- `test-basic-formatting.docx`
- `test-headings-standard.docx`
- `test-headings-custom.docx`
- `test-headings-inline.docx`
- `test-lists.docx`
- `test-tables.docx`
- `test-images.docx`
- `test-textboxes.docx`
- `test-headers-footers-html.docx`
- `test-headers-footers-images.docx`
- `test-page-numbers-only.docx`
- `test-complex-header-footer.docx`
- `test-page-breaks.docx`
- `test-page-border-*.docx` (multiple files for different border styles)
- `test-margins-*.docx` (multiple files for different margin settings)
- `test-page-size-*.docx` (multiple files for different page sizes)
- `test-fonts-*.docx` (multiple files for different font settings)
- `test-combo-*.docx` (multiple files for font/page combinations)
- `test-edge-*.docx` (multiple files for edge cases)

## Test Coverage

The test suite covers:

### Text Formatting
- Font styles: bold, italic, underline, strikethrough
- Text colors and background colors
- Font families and sizes
- Text alignment: left, center, right, justify

### Document Structure
- Headings (H1-H6) with custom styling
- Paragraphs and line breaks
- Lists (ordered, unordered, nested)
- Tables with borders and styling

### Media and Images
- Base64 encoded images
- URL images (with graceful failure handling)
- Image sizing and positioning
- Alt text support

### Layout and Styling
- Text boxes (divs with background/border)
- Page breaks (manual and CSS-based)
- Headers and footers (HTML and image-based)
- Page numbers with alignment options
- Page borders and margins
- Different page sizes (A4, Letter, Legal, custom)

### Advanced Features
- Custom heading replacements
- Font customization (family, size, line height)
- Page size variations
- Complex nested structures
- Unicode and special character support

### Error Handling
- Malformed HTML
- Invalid options
- Missing resources
- Large content handling
- Edge cases and boundary conditions

## Adding New Tests

To add a new test:

1. Create a new JavaScript file in the `tests/` directory
2. Export a default async function that runs the test
3. Return `true` on success, `false` on failure
4. Add the filename to the `testFiles` array in `run-tests.js`
5. Document the test in this README

Example test structure:
```javascript
async function testNewFeature() {
  console.log('Testing new feature...');

  const html = '<p>Test content</p>';
  const converter = new HtmlToDocx();
  const outputPath = path.join(__dirname, 'test-new-feature.docx');

  try {
    await converter.convertHtmlToDocxFile(html, outputPath);
    console.log('✓ New feature test completed successfully');
    return true;
  } catch (error) {
    console.error('✗ New feature test failed:', error.message);
    return false;
  }
}

export default testNewFeature;
```

## Troubleshooting

### Tests Fail to Run
- Ensure Node.js 14+ is installed
- Check that all dependencies are installed: `npm install`
- Verify file permissions for writing DOCX files

### Generated DOCX Files Look Incorrect
- Open the files in Microsoft Word for best compatibility
- Some features may not display correctly in all DOCX viewers
- Check the console output for any warnings or errors

### Performance Issues
- Large content tests may take time to complete
- Some tests generate multiple output files
- Network-dependent tests (URL images) may fail if offline

## Integration with CI/CD

You can integrate these tests into your CI/CD pipeline:

```bash
# Run all tests
node tests/run-tests.js

# Check exit code
if [ $? -eq 0 ]; then
  echo "All tests passed"
else
  echo "Some tests failed"
  exit 1
fi
```

## Contributing

When adding new features to the library:
1. Add corresponding tests to ensure the feature works correctly
2. Update this README with documentation for new tests
3. Ensure all existing tests still pass
4. Test edge cases and error conditions
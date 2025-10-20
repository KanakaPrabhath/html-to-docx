# HTML to DOCX Demo

This demo application showcases the functionality of the `@kanaka-prabhath/html-to-docx` library, which converts HTML content with CSS styles to Microsoft Word DOCX format using OOXML.

## Features

- Convert HTML strings to DOCX buffers
- Save DOCX files directly to disk
- Support for various HTML elements and CSS styling
- Configurable font settings and formatting options

## Setup

### Linking the Local Package for Development

To test with the local development version of the html-to-docx library:

1. **Link the main package globally:**
   ```bash
   cd ../  # Go to the root directory
   npm link
   ```

2. **Link the package in the demo:**
   ```bash
   cd demo
   npm link @kanaka-prabhath/html-to-docx
   ```

This will create a symbolic link so that the demo uses the local version of the package instead of a published one.

### Alternative: Install from npm (for production testing)

If you want to test with the published version:

```bash
npm install @kanaka-prabhath/html-to-docx
```

## Usage

### Running the Demo

```bash
npm start
```

This will convert a sample HTML document to a DOCX file (`demo-output.docx`).

### Running Tests

The demo includes comprehensive tests for both usage patterns:

#### Class-Based Tests (Recommended for multiple conversions)
```bash
node test-class.js
```

This runs tests using the `HtmlToDocx` class with:
- Constructor options
- Runtime options override
- Buffer output with merged options
- Advanced features (headers, footers, page numbers)
- Multiple conversions with the same instance

#### Direct Function Tests (For single conversions)
```bash
node test-direct-functions.js
```

This runs tests using direct function calls with:
- Basic conversions without options
- Conversions with inline options
- Buffer output
- Advanced features (headers, footers, page numbers)
- Heading replacements

#### Original Comprehensive Tests
```bash
npm test
```

This runs the original comprehensive test suite covering various HTML elements and styling scenarios.

All test output files will be generated in the demo directory with descriptive names.

## API Usage

### Basic Conversion to Buffer

```javascript
import { convertHtmlToDocx } from '@kanaka-prabhath/html-to-docx';

const html = '<p>Hello <strong>World</strong>!</p>';
const docxBuffer = await convertHtmlToDocx(html);
// Use the buffer as needed
```

### Direct File Output

```javascript
import HtmlToDocx from '@kanaka-prabhath/html-to-docx';

const converter = new HtmlToDocx({
  fontSize: 12,
  fontFamily: 'Arial',
  lineHeight: 1.15
});

await converter.convertHtmlToDocxFile(html, 'output.docx');
```

### Conversion to Buffer with Options

```javascript
import HtmlToDocx from '@kanaka-prabhath/html-to-docx';

const converter = new HtmlToDocx({
  fontSize: 14,
  fontFamily: 'Calibri'
});

const buffer = await converter.convertHtmlToDocx(html);
```

## Supported HTML Elements

- Headings (h1-h6)
- Paragraphs (p)
- Text formatting (strong, em, u)
- Lists (ul, ol, li)
- Line breaks (br)
- Spans with inline styles
- Basic CSS styling (color, background-color, font-size, font-family, text-align)

## Output

All generated DOCX files can be opened with Microsoft Word or any compatible word processor.
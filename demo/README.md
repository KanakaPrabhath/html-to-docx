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

```bash
npm test
```

This runs comprehensive tests including:
- Text alignment
- Complex formatting (headings, bold, italic, underline)
- Color and background styling
- Font size and family
- Lists and line breaks
- Edge cases

Test output files will be generated in the demo directory.

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
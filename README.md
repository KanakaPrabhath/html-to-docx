# HTML to DOCX Converter

A Node.js library for converting HTML with inline styles to DOCX files using OOXML (Office Open XML).

## Features

- ✅ Convert HTML to DOCX with proper formatting
- ✅ Support for text alignment (left, center, right, justify)
- ✅ Support for text formatting (bold, italic, underline)
- ✅ Support for colors and background colors
- ✅ Support for font family and font size
- ✅ Support for headings (h1-h6)
- ✅ Support for lists (ul, ol)
- ✅ Support for images (URLs and base64)
- ✅ Support for headers and footers (HTML or full-width base64 images)
- ✅ Support for page numbers in footer with alignment options
- ✅ Support for line breaks
- ✅ Generates valid OOXML structure
- ✅ Customizable default styles
- ✅ Automatic HTML cleaning and sanitization

## Installation

```bash
npm install
```

## Usage

### Basic Usage

```javascript
const HtmlToDocx = require('./index');

const converter = new HtmlToDocx();

const html = `
  <p style="text-align: left;">Left aligned text</p>
  <p style="text-align: center;">Center aligned text</p>
  <p style="text-align: right;">Right aligned text</p>
  <p style="text-align: right;"><br></p>
`;

// Convert to buffer
const buffer = await converter.convertHtmlToDocx(html);

// Or save to file
await converter.convertHtmlToDocxFile(html, 'output.docx');
```

### With Headers and Footers

```javascript
const converter = new HtmlToDocx();

const html = `<p>Main document content</p>`;

const options = {
  header: '<p style="text-align: center;">Company Header</p>', // HTML header
  footer: '<p style="text-align: center;">Page Footer</p>'   // HTML footer
};

const buffer = await converter.convertHtmlToDocx(html, options);
```

### With Page Size and Margins

```javascript
const options = {
  pageSize: 'A4', // 'A4', 'Letter', 'Legal', or custom {width: 8.5, height: 11}
  marginTop: 1,    // 1 inch top margin
  marginRight: 1,  // 1 inch right margin
  marginBottom: 1, // 1 inch bottom margin
  marginLeft: 1,   // 1 inch left margin
  marginHeader: 0.5, // 0.5 inch header margin
  marginFooter: 0.5  // 0.5 inch footer margin
};

const buffer = await converter.convertHtmlToDocx(html, options);
```

### With Page Numbers

```javascript
const options = {
  enablePageNumbers: true,        // Enable/disable page numbers (default: false)
  pageNumberAlignment: 'right'    // 'left', 'center', or 'right' (default: 'right')
};

const buffer = await converter.convertHtmlToDocx(html, options);
```

Page numbers appear in the footer area and update automatically in Word.

### Supported HTML Elements

#### Text Alignment
```html
<p style="text-align: left;">Left aligned</p>
<p style="text-align: center;">Center aligned</p>
<p style="text-align: right;">Right aligned</p>
<p style="text-align: justify;">Justified text</p>
```

#### Text Formatting
```html
<p><strong>Bold text</strong></p>
<p><em>Italic text</em></p>
<p><u>Underlined text</u></p>
<p><strong><em>Bold and italic</em></strong></p>
```

#### Colors
```html
<p style="color: #FF0000;">Red text</p>
<p style="background-color: #FFFF00;">Yellow background</p>
```

#### Font Styling
```html
<p style="font-size: 14px;">Larger text</p>
<p style="font-family: Arial;">Arial font</p>
<p style="font-weight: bold; font-style: italic;">Bold italic</p>
```

#### Headings
```html
<h1>Heading 1</h1>
<h2>Heading 2</h2>
<h3>Heading 3</h3>
```

#### Lists
```html
<ul>
  <li>Item 1</li>
  <li>Item 2</li>
</ul>

<ol>
  <li>First</li>
  <li>Second</li>
</ol>
```

#### Images
```html
<!-- URL images (automatically downloaded) -->
<img src="https://example.com/image.png" alt="Description" width="300" height="200">

<!-- Base64 encoded images -->
<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==" alt="Base64 Image">
```

#### Line Breaks
```html
<p>Line 1<br>Line 2</p>
<p><br></p> <!-- Empty paragraph -->
```

## API Reference

### Class: HtmlToDocx

#### Constructor

```javascript
new HtmlToDocx(options)
```

**Options:**
- `fontSize` (number): Default font size in points (default: 11)
- `fontFamily` (string): Default font family (default: 'Calibri')
- `lineHeight` (number): Default line height multiplier (default: 1.15)
- `pageSize` (string|object): Page size - 'A4', 'Letter', 'Legal', or custom {width: number, height: number} in inches (default: 'A4')
- `marginTop` (number): Top margin in inches (default: 1)
- `marginRight` (number): Right margin in inches (default: 1)
- `marginBottom` (number): Bottom margin in inches (default: 1)
- `marginLeft` (number): Left margin in inches (default: 1)
- `marginHeader` (number): Header margin in inches (default: 0.5)
- `marginFooter` (number): Footer margin in inches (default: 0.5)
- `marginGutter` (number): Gutter margin in inches (default: 0)
- `enablePageNumbers` (boolean): Enable/disable page numbers in footer (default: false)
- `pageNumberAlignment` (string): Page number alignment - 'left', 'center', or 'right' (default: 'right')

#### Methods

##### convertHtmlToDocx(html, options)

Convert HTML string to DOCX buffer.

**Parameters:**
- `html` (string): HTML content to convert
- `options` (object): Conversion options
  - `header` (string): HTML content or base64 image data URL for document header
  - `footer` (string): HTML content or base64 image data URL for document footer
  - `enablePageNumbers` (boolean): Enable/disable page numbers in footer (default: false)
  - `pageNumberAlignment` (string): Page number alignment - 'left', 'center', or 'right' (default: 'right')
  - `pageSize` (string|object): Page size - 'A4', 'Letter', 'Legal', or custom {width: number, height: number} in inches (default: 'A4')
  - `marginTop` (number): Top margin in inches (default: 1)
  - `marginRight` (number): Right margin in inches (default: 1)
  - `marginBottom` (number): Bottom margin in inches (default: 1)
  - `marginLeft` (number): Left margin in inches (default: 1)
  - `marginHeader` (number): Header margin in inches (default: 0.5)
  - `marginFooter` (number): Footer margin in inches (default: 0.5)
  - `marginGutter` (number): Gutter margin in inches (default: 0)

**Returns:**
- Promise<Buffer>: DOCX file buffer

##### convertHtmlToDocxFile(html, outputPath, options)

Convert HTML string to DOCX file.

**Parameters:**
- `html` (string): HTML content to convert
- `outputPath` (string): Path where to save the DOCX file
- `options` (object): Conversion options (same as above)

**Returns:**
- Promise<void>

## HTML Cleaning and Sanitization

The library automatically cleans and sanitizes input HTML to ensure compatibility with Microsoft Word:

### Automatic Cleaning Features

- **Removes dangerous elements**: `<script>`, `<style>`, `<iframe>`, `<object>`, `<embed>`, `<form>`, `<input>`, `<button>`, `<select>`, `<textarea>`
- **Strips invalid Unicode characters**: Control characters that could cause corruption
- **Simplifies complex CSS**: Keeps only safe properties like `text-align`, `font-weight`, `color`, etc.
- **Fixes malformed HTML**: Basic cleanup of broken tags
- **Ensures content structure**: Adds minimal content if HTML is empty

### Safe CSS Properties

Only these CSS properties are preserved in the output:

- `text-align`
- `font-weight`
- `font-style`
- `text-decoration`
- `color`
- `background-color`
- `font-size`
- `font-family`

Complex or unsafe CSS (like `position: absolute`, `display: none`) is automatically removed.

### Example of Cleaning

**Input HTML:**
```html
<p>Normal text</p>
<script>alert('dangerous');</script>
<p style="text-align: center; position: absolute;">Centered</p>
```

**After Cleaning:**
```html
<p>Normal text</p>
<p style="text-align: center;">Centered</p>
```

## Architecture

The library uses OOXML (Office Open XML) to generate valid DOCX files:

1. **HTML Parsing**: Uses JSDOM to parse HTML into DOM
2. **Style Extraction**: Parses inline styles from elements
3. **Image Processing**: Downloads URL images or decodes base64 images and embeds them as media files
4. **OOXML Generation**: Converts DOM nodes to OOXML elements
5. **ZIP Packaging**: Uses JSZip to package as .docx file

### File Structure

```
html-to-docx/
├── lib/
│   ├── index.js           # Main converter class
│   ├── converter.js       # Main conversion logic
│   ├── htmlParser.js      # HTML parsing and processing
│   ├── ooxmlGenerator.js  # OOXML element generation
│   ├── mediaHandler.js    # Image processing and media management
│   ├── styleParser.js     # CSS style parsing
│   ├── templates.js       # OOXML templates
│   └── utils.js           # Utility functions
├── demo/
│   ├── index.js           # Demo script
│   └── package.json       # Demo dependencies
├── examples.js            # Additional examples
├── package.json           # Package configuration
└── README.md              # This file
```

## Testing

```bash
node test.js
```

This will create a `test-output.docx` file with various formatting examples.

## Integration with TuteMaker

This library is designed to be used in the TuteMaker Electron app for exporting HTML content from the editor to Word documents.

### Example Integration

```javascript
// In Electron main process
const HtmlToDocx = require('./electron/html-to-docx');

ipcMain.handle('export-to-docx', async (event, htmlContent, outputPath) => {
  try {
    const converter = new HtmlToDocx({
      fontSize: 11,
      fontFamily: 'Calibri'
    });
    
    await converter.convertHtmlToDocxFile(htmlContent, outputPath);
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});
```

## Technical Details

### OOXML Structure

The generated DOCX file contains:

- `[Content_Types].xml` - File type definitions
- `_rels/.rels` - Package relationships
- `word/document.xml` - Main document content
- `word/_rels/document.xml.rels` - Document relationships (includes image references)
- `word/styles.xml` - Document styles
- `word/fontTable.xml` - Font definitions
- `word/settings.xml` - Document settings
- `word/media/` - Embedded images and media files
- `docProps/app.xml` - Application properties
- `docProps/core.xml` - Core properties

### Image Processing

Images in HTML are automatically processed and embedded:

1. **URL Images**: Downloaded via HTTP/HTTPS and cached as media files
2. **Base64 Images**: Decoded and embedded directly
3. **Media Files**: Stored in `word/media/` with unique filenames
4. **Relationships**: Referenced in `document.xml.rels` with proper IDs
5. **OOXML Generation**: Images rendered as inline drawings with proper sizing

### Paragraph Properties (`<w:pPr>`)

- Alignment: `<w:jc w:val="left|center|right|both"/>`
- Spacing: `<w:spacing w:line="..." w:lineRule="auto"/>`
- Style: `<w:pStyle w:val="Heading1"/>`

### Run Properties (`<w:rPr>`)

- Bold: `<w:b/>`
- Italic: `<w:i/>`
- Underline: `<w:u w:val="single"/>`
- Color: `<w:color w:val="RRGGBB"/>`
- Font Size: `<w:sz w:val="halfPoints"/>`
- Font Family: `<w:rFonts w:ascii="fontName"/>`

## License

MIT

## Author

TuteMaker Team

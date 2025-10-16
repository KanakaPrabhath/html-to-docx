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

### With Custom Options

```javascript
const converter = new HtmlToDocx({
  fontSize: 12,
  fontFamily: 'Arial',
  lineHeight: 1.5
});
```

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

#### Methods

##### convertHtmlToDocx(html)

Convert HTML string to DOCX buffer.

**Parameters:**
- `html` (string): HTML content to convert

**Returns:**
- Promise<Buffer>: DOCX file buffer

##### convertHtmlToDocxFile(html, outputPath)

Convert HTML string to DOCX file.

**Parameters:**
- `html` (string): HTML content to convert
- `outputPath` (string): Path where to save the DOCX file

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
3. **OOXML Generation**: Converts DOM nodes to OOXML elements
4. **ZIP Packaging**: Uses JSZip to package as .docx file

### File Structure

```
html-to-docx/
├── index.js           # Main converter class
├── templates.js       # OOXML templates
├── test.js           # Test file
├── package.json      # Package configuration
└── README.md         # This file
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
- `word/styles.xml` - Document styles
- `word/fontTable.xml` - Font definitions
- `word/settings.xml` - Document settings
- `docProps/app.xml` - Application properties
- `docProps/core.xml` - Core properties

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

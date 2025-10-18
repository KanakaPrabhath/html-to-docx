# HTML to DOCX Converter

A powerful Node.js library for converting HTML with inline styles to Microsoft Word DOCX files using OOXML (Office Open XML). Perfect for generating professional documents from web content, rich text editors, or any HTML source.

## ✨ Features

### 📝 Text Formatting
- **Text Alignment**: Left, center, right, and justified alignment
- **Font Styling**: Bold, italic, underline, strikethrough
- **Colors**: Text color and background color support
- **Typography**: Font family, font size, and line height
- **Headings**: H1-H6 with automatic styling

### 📋 Lists & Structure
- **Unordered Lists**: Bulleted lists with proper indentation
- **Ordered Lists**: Numbered lists with automatic numbering
- **Nested Lists**: Support for nested list structures
- **Custom Indentation**: Configurable list indentation levels

### 🖼️ Media & Images
- **URL Images**: Automatic download and embedding of web images
- **Base64 Images**: Direct embedding of base64 encoded images
- **Image Sizing**: Width and height control via attributes or CSS
- **Alt Text**: Accessibility support with alt text preservation

### 📄 Document Layout
- **Headers & Footers**: HTML content or full-width images
- **Page Numbers**: Automatic page numbering with alignment options
- **Page Breaks**: Manual page breaks and section breaks
- **Margins**: Customizable page margins (top, bottom, left, right)
- **Page Size**: Support for A4, Letter, Legal, and custom sizes

### 🎨 Advanced Styling
- **Text Boxes**: Div elements with background colors and borders
- **Tables**: Full table support with headers and cells
- **Spacing**: Custom paragraph spacing and line height
- **Borders**: Border styling for text boxes and tables
- **Background Colors**: Element background color support

### 🔧 Technical Features
- **OOXML Generation**: Creates valid Office Open XML structure
- **HTML Sanitization**: Automatic cleaning of unsafe HTML elements
- **CSS Parsing**: Support for safe CSS properties
- **Media Management**: Efficient handling of embedded images
- **ZIP Packaging**: Proper DOCX file generation using JSZip

## 📦 Installation

```bash
npm install @kanaka-prabhath/html-to-docx
```

## 🚀 Quick Start

### Basic Usage

```javascript
import { HtmlToDocx } from '@kanaka-prabhath/html-to-docx';

const converter = new HtmlToDocx();

const html = `
  <h1>Welcome to DOCX Export</h1>
  <p>This is a <strong>bold</strong> paragraph with <em>italic</em> text.</p>
  <ul>
    <li>Feature 1</li>
    <li>Feature 2</li>
  </ul>
`;

// Convert to buffer
const buffer = await converter.convertHtmlToDocx(html);

// Save to file
await converter.convertHtmlToDocxFile(html, 'document.docx');
```

### Advanced Configuration

```javascript
const converter = new HtmlToDocx({
  fontSize: 12,           // Default font size in points
  fontFamily: 'Arial',    // Default font family
  lineHeight: 1.2,        // Line height multiplier
  pageSize: 'A4',         // Page size: 'A4', 'Letter', 'Legal', or custom object
  marginTop: 1,           // Top margin in inches
  marginRight: 1,         // Right margin in inches
  marginBottom: 1,        // Bottom margin in inches
  marginLeft: 1,          // Left margin in inches
  marginHeader: 0.5,      // Header margin in inches
  marginFooter: 0.5       // Footer margin in inches
});
```

## 📖 API Reference

### Class: HtmlToDocx

#### Constructor

```javascript
new HtmlToDocx(options?)
```

**Parameters:**
- `options` (Object, optional): Configuration options
  - `fontSize` (number): Default font size in points (default: 11)
  - `fontFamily` (string): Default font family (default: 'Calibri')
  - `lineHeight` (number): Default line height multiplier (default: 1.15)
  - `pageSize` (string|object): Page size - 'A4', 'Letter', 'Legal', or custom `{width: number, height: number}` in inches (default: 'A4')
  - `marginTop` (number): Top margin in inches (default: 1)
  - `marginRight` (number): Right margin in inches (default: 1)
  - `marginBottom` (number): Bottom margin in inches (default: 1)
  - `marginLeft` (number): Left margin in inches (default: 1)
  - `marginHeader` (number): Header margin in inches (default: 0.5)
  - `marginFooter` (number): Footer margin in inches (default: 0.5)
  - `marginGutter` (number): Gutter margin in inches (default: 0)

#### Methods

##### convertHtmlToDocx(html, options?)

Convert HTML string to DOCX buffer.

**Parameters:**
- `html` (string): HTML content to convert
- `options` (Object, optional): Conversion options (same as constructor plus additional runtime options)

**Returns:** Promise<Buffer> - DOCX file buffer

##### convertHtmlToDocxFile(html, outputPath, options?)

Convert HTML string and save to DOCX file.

**Parameters:**
- `html` (string): HTML content to convert
- `outputPath` (string): Path where to save the DOCX file
- `options` (Object, optional): Conversion options

**Returns:** Promise<void>

### Runtime Options

Additional options that can be passed to conversion methods:

- `header` (string): HTML content or base64 image data URL for document header
- `footer` (string): HTML content or base64 image data URL for document footer
- `enablePageNumbers` (boolean): Enable/disable page numbers in footer (default: false)
- `pageNumberAlignment` (string): Page number alignment - 'left', 'center', or 'right' (default: 'right')
- `headingReplacements` (Array<string>): Custom HTML templates for headings (H1, H2, H3, etc.)

## 🎯 Supported HTML Elements

### Text Elements
```html
<p style="text-align: center;">Centered paragraph</p>
<strong>Bold text</strong>
<em>Italic text</em>
<u>Underlined text</u>
<strike>Strikethrough text</strike>
<span style="color: #FF0000;">Red text</span>
```

### Headings
```html
<h1>Main Title</h1>
<h2>Section Header</h2>
<h3>Subsection</h3>
```

### Lists
```html
<ul>
  <li>Unordered item</li>
  <li>Another item</li>
</ul>

<ol>
  <li>Ordered item 1</li>
  <li>Ordered item 2</li>
</ol>
```

### Images
```html
<!-- URL images -->
<img src="https://example.com/image.png" alt="Description" width="300" height="200">

<!-- Base64 images -->
<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==" alt="Base64 Image">
```

### Tables
```html
<table>
  <tr>
    <th>Header 1</th>
    <th>Header 2</th>
  </tr>
  <tr>
    <td>Data 1</td>
    <td>Data 2</td>
  </tr>
</table>
```

### Text Boxes
```html
<div style="background-color: #FFFF00; padding: 10px; border: 1px solid #000;">
  <p>Content in a text box</p>
</div>
```

### Page Breaks
```html
<p>Content before break</p>
<page-break></page-break>
<p>Content after break</p>
```

## 🎨 CSS Property Support

### Supported Properties
- `text-align`: left, center, right, justify
- `font-weight`: normal, bold, or numeric values ≥ 600
- `font-style`: normal, italic
- `text-decoration`: underline, line-through
- `color`: Hex colors (#RGB, #RRGGBB), named colors
- `background-color`: Hex colors and named colors
- `font-size`: px, pt units
- `font-family`: Font family names
- `margin`: All margin properties for spacing
- `padding`: Padding for text boxes
- `border`: Border styling for text boxes
- `border-radius`: Border radius for text boxes
- `width`/`height`: Dimensions for images and text boxes
- `page-break-before`/`page-break-after`: always
- `float`: left, right (for images)

### Automatic Sanitization
The library automatically removes or ignores:
- Dangerous elements: `<script>`, `<iframe>`, `<object>`, etc.
- Unsafe CSS: `position: absolute`, complex layouts
- Invalid Unicode characters
- Malformed HTML structures

## 📄 Headers & Footers

### HTML Headers/Footers
```javascript
const options = {
  header: '<p style="text-align: center; font-size: 10pt;">Company Header</p>',
  footer: '<p style="text-align: center; font-size: 10pt;">Page Footer</p>'
};
```

### Image Headers/Footers
```javascript
const options = {
  header: 'data:image/png;base64,...', // Full-width header image
  footer: 'data:image/png;base64,...'  // Full-width footer image
};
```

### Page Numbers
```javascript
const options = {
  enablePageNumbers: true,
  pageNumberAlignment: 'center' // 'left', 'center', or 'right'
};
```

## 🔧 Advanced Usage

### Custom Heading Styles
```javascript
const options = {
  headingReplacements: [
    // H1 replacement
    '<div style="background-color: #E6E6E6; padding: 10px;"><h1>HEADING_TEXT</h1></div>',
    // H2 replacement
    '<div style="border-left: 4px solid #0066CC; padding-left: 10px;"><h2>HEADING_TEXT</h2></div>',
    // H3 replacement
    '<h3 style="color: #0066CC;">HEADING_TEXT</h3>'
  ]
};
```

### Electron Integration
```javascript
// In Electron main process
const HtmlToDocx = require('@kanaka-prabhath/html-to-docx');

ipcMain.handle('export-to-docx', async (event, { html, outputPath, options }) => {
  try {
    const converter = new HtmlToDocx(options);
    await converter.convertHtmlToDocxFile(html, outputPath);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});
```

### Batch Processing
```javascript
const documents = [
  { html: '<h1>Doc 1</h1>', name: 'document1' },
  { html: '<h1>Doc 2</h1>', name: 'document2' }
];

for (const doc of documents) {
  await converter.convertHtmlToDocxFile(doc.html, `${doc.name}.docx`);
}
```

## 🏗️ Architecture

The library processes HTML through several stages:

1. **HTML Parsing**: Uses JSDOM to parse and clean HTML
2. **Style Extraction**: Parses inline CSS properties
3. **Element Processing**: Converts HTML elements to OOXML
4. **Media Handling**: Downloads and embeds images
5. **OOXML Generation**: Creates Office Open XML structure
6. **ZIP Packaging**: Packages everything into a .docx file

### File Structure Inside DOCX
```
document.docx/
├── [Content_Types].xml
├── _rels/.rels
├── word/
│   ├── document.xml          # Main document content
│   ├── styles.xml            # Document styles
│   ├── numbering.xml         # List numbering definitions
│   ├── header1.xml           # Header content (if used)
│   ├── footer1.xml           # Footer content (if used)
│   ├── _rels/                # Relationships
│   └── media/                # Embedded images
└── docProps/
    ├── app.xml               # Application properties
    └── core.xml              # Core properties
```

## 🧪 Testing

```bash
# Run the demo
cd demo
npm install
npm test

# This creates test-output.docx with various formatting examples
```

## 📋 Requirements

- Node.js 14+
- Dependencies: `jsdom`, `jszip`

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Add tests for new features
4. Ensure all tests pass
5. Submit a pull request

## 📄 License

MIT License - see LICENSE file for details

## 👥 Author

Kanaka Prabhath (TuteMaker Team)

## 🙏 Acknowledgments

Built with [JSDOM](https://github.com/jsdom/jsdom) for HTML parsing and [JSZip](https://github.com/Stuk/jszip) for ZIP file generation.

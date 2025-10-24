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
- **Page Breaks**: Manual page breaks, CSS page breaks, and section breaks
- **Page Borders**: Customizable border styles, colors, sizes, and radius
- **Margins**: Customizable page margins (top, bottom, left, right)
- **Page Size**: Support for A4, Letter, Legal, and custom sizes

### 🎨 Advanced Styling
- **Text Boxes**: Div elements with background colors, borders, padding, width, and border-radius
- **Section Header Images**: Images with top/bottom text wrapping using `data-section-header`
- **Tables**: Full table support with headers and cells
- **Spacing**: Custom paragraph spacing and line height
- **Borders**: Border styling for text boxes and tables
- **Background Colors**: Element background color support

### 🔧 Technical Features
- **OOXML Generation**: Creates valid Office Open XML structure
- **HTML Sanitization**: Automatic cleaning of unsafe HTML elements
- **CSS Parsing**: Support for safe CSS properties
- **Media Management**: Efficient handling of embedded images
- **Buffer Output**: Direct buffer output for memory-based processing
- **Edge Case Handling**: Robust processing of malformed HTML and special characters
- **Performance Optimized**: Efficient processing for large documents
- **ZIP Packaging**: Proper DOCX file generation using JSZip

## 📦 Installation

```bash
npm install @kanaka-prabhath/html-to-docx@1.0.2
```

## 🚀 Quick Start

### Basic Usage

#### Using Direct Functions (For single conversions)

```javascript
import { convertHtmlToDocx } from '@kanaka-prabhath/html-to-docx';
import fs from 'fs';

const htmlContent = `
  <h1>Welcome to DOCX Export</h1>
  <p>This is a <strong>bold</strong> paragraph with <em>italic</em> text.</p>
`;

const options = {
  pageSize: 'A4',
  marginTop: 1,
  marginRight: 1,
  marginBottom: 1,
  marginLeft: 1,
  enablePageNumbers: true,
  pageNumberAlignment: 'center'
};

const docxBuffer = await convertHtmlToDocx(htmlContent, options);
fs.writeFileSync('demo-output.docx', docxBuffer);
```

### Advanced Configuration

```javascript
const options = {
  pageSize: 'A4', // Default page size (A4, Letter, Legal, or custom {width: number, height: number} in inches)
  marginTop: 1, // 1 inch top margin
  marginRight: 1, // 1 inch right margin
  marginBottom: 1, // 1 inch bottom margin
  marginLeft: 1, // 1 inch left margin
  marginHeader: 0.5, // 0.5 inch header margin
  marginFooter: 0.5, // 0.5 inch footer margin
  headerHeight: 1, // 1 inch header height
  footerHeight: 1, // 1 inch footer height
  enableHeader: true, // Enable/disable header (default: true if header content provided)
  enableFooter: true, // Enable/disable footer (default: true if footer content provided)
  enablePageNumbers: true, // Enable/disable page numbers
  pageNumberAlignment: 'center', // Page number alignment: 'left', 'center', or 'right'
  header: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9ewAAAABJRU5ErkJggg==', // Blue colored image for header (positioned at top-left 0,0, full width)
  footer: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9ewAAAABJRU5ErkJggg==',  // Blue colored image for footer (positioned at bottom-left 0,bottom, full width)
  headingReplacements: [
    `<div data-h1 class="textbox" style="border: 1px solid #000000ff; border-radius: 5px; padding: 0px 5px 0px 5px; background-color: #000000ff; width:100%;">
<p data-no-spacing style="color: #ffffffff; margin: 0; font-size: 21px; font-weight: bold;">HEADING_TEXT</p>
</div>`,
    `<div data-h2 class="textbox" style="border: 1px solid #00b118ff; border-radius: 5px; padding: 0px; background-color: #a10101ff; width:100%;">
<p data-no-spacing style="color: #ffffffff; margin: 0; font-size: 19px; font-weight: bold;">HEADING_TEXT</p>
</div>`,
    `<div data-h3 class="textbox" style="border: 1px solid #00b118ff; border-radius: 5px; padding: 0px; background-color: #a10101ff; width:100%;">
<p data-no-spacing style="color: #ffffffff; margin: 0; font-size: 16px; font-weight: bold;">HEADING_TEXT</p>
</div>`
  ]
};

const converter = new HtmlToDocx(options);
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
  - `headerHeight` (number): Header height in inches (default: undefined)
  - `footerHeight` (number): Footer height in inches (default: undefined)
  - `marginGutter` (number): Gutter margin in inches (default: 0)
  - `enablePageBorder` (boolean): Enable page borders aligned with margins (default: false)
  - `pageBorder` (object): Page border configuration `{style: 'single', color: '000000', size: 4, radius: 0}` (optional)

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
- `enablePageBorder` (boolean): Enable page borders aligned with margins (default: false)
- `pageBorder` (object): Page border configuration `{style: 'single', color: '000000', size: 4, radius: 0}` (optional)
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

<!-- CSS page breaks -->
<p style="page-break-before: always;">Content with page break before</p>
<p style="page-break-after: always;">Content with page break after</p>
```

### Section Header Images
```html
<!-- Regular section header image with text wrapping -->
<img data-section-header src="image.png" alt="Section Header" style="width: 400px; height: 150px; margin: 20px; display: block;"/>
<p>Text that wraps above and below the section header image.</p>

<!-- Cover image that spans full page width -->
<img data-section-header data-cover src="cover-image.png" alt="Cover Image" style="height: 200px; display: block;"/>
<p>Text that appears below the full-width cover image.</p>
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
- `border`: Border styling for text boxes (width, style, color)
- `border-radius`: Border radius for text boxes
- `width`/`height`: Dimensions for images and text boxes
- `page-break-before`/`page-break-after`: always
- `float`: left, right (for images)

### Special Attributes
- `data-section-header`: Positions image at top of content section with text wrapping
- `data-cover`: Makes section header image span full page width
- `data-no-spacing`: Removes default spacing from paragraphs

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

### Page Borders

Page borders support various styles, colors, sizes, and radius for professional document appearance.

```javascript
const options = {
  enablePageBorder: true,
  pageBorder: {
    style: 'double',    // 'single', 'double', 'thick', 'dotted', 'dashed'
    color: 'FF0000',    // Hex color without #
    size: 8,            // Border thickness in points
    radius: 5           // Border radius in points (rounded corners)
  }
};
```
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

### Advanced Features Examples

#### Section Header Images and Cover Images
```javascript
const html = `
  <h1>Document with Section Headers</h1>
  
  <img data-section-header src="data:image/png;base64,iVBORw0KGgo..." alt="Section Header" style="width: 400px; height: 150px; margin: 20px;"/>
  <p>Content that wraps above and below the section header image.</p>
  
  <page-break></page-break>
  
  <img data-section-header data-cover src="data:image/png;base64,iVBORw0KGgo..." alt="Cover Image" style="height: 200px;"/>
  <p>Content below the full-width cover image.</p>
`;

const options = {
  enablePageBorder: true,
  pageBorder: {
    style: 'double',
    color: '000000',
    size: 6,
    radius: 8
  }
};

await converter.convertHtmlToDocxFile(html, 'advanced-document.docx', options);
```

#### Enhanced Page Breaks and Text Boxes
```javascript
const html = `
  <div style="background-color: #F0F0F0; padding: 15px; border: 2px solid #333; border-radius: 5px;">
    <h2>Text Box with Rounded Borders</h2>
    <p>This content appears in a styled text box.</p>
  </div>
  
  <p style="page-break-after: always;">This paragraph forces a page break after it.</p>
  
  <h2>New Page Content</h2>
  <p>This appears on a new page due to the CSS page break.</p>
`;

await converter.convertHtmlToDocxFile(html, 'enhanced-layout.docx');
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

Kanaka Prabhath 

## 🙏 Acknowledgments

Built with [JSDOM](https://github.com/jsdom/jsdom) for HTML parsing and [JSZip](https://github.com/Stuk/jszip) for ZIP file generation.

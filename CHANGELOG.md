# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.2] - 2025-10-24

### Added
- Page borders with customizable styles (single, double, thick, dotted, dashed), colors, sizes, and border radius
- Text boxes (div elements with background color, borders, padding, width, and border-radius)
- Section header images with `data-section-header` attribute for top/bottom text wrapping
- Manual page breaks using `<page-break>` elements
- CSS page breaks (`page-break-before` and `page-break-after` properties)
- Page breaks support in various contexts (lists, tables, text boxes)
- Buffer output support via `convertHtmlToDocx()` method
- Enhanced edge case handling for malformed HTML, special characters, and Unicode
- Performance improvements for large documents with many elements
- Image support in HTML headers and footers
- Advanced font customization per element (size, family, inline styles)
- Custom page size dimensions (beyond standard presets)
- Complex nested structures support (tables, lists, text boxes within each other)

### Changed
- Improved error handling for invalid configuration options
- Enhanced document generation for edge cases and large content

## [1.0.1] - 2025-10-20

### Added
- Direct function exports (`convertHtmlToDocx`, `convertHtmlToDocxFile`) for single-use conversions
- Runtime options merging in class methods for flexible configuration
- Comprehensive API documentation in README
- TypeScript type definitions (@types/jsdom, @types/jszip) as dev dependencies
- Enhanced demo with separate test scripts for class-based and direct function usage

### Changed
- Improved README with detailed usage examples, API reference, and advanced configuration options
- Updated demo README with better testing instructions
- Enhanced JSDoc comments in lib/index.js for better developer experience
- Updated package.json with version bump and dev dependencies

### Fixed
- Better option merging in HtmlToDocx class methods

## [1.0.0] - 2025-10-XX

### Added
- Initial release of HTML to DOCX converter
- Support for HTML elements: headings, paragraphs, lists, text formatting
- CSS styling support: colors, fonts, alignment, spacing
- Image embedding (URL and base64)
- Table support
- Header and footer functionality
- Page numbering
- Custom page sizes and margins
- OOXML generation for Microsoft Word compatibility
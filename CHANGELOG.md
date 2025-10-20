# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
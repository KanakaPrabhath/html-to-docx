import { convertHtmlToDocx } from '@kanaka-prabhath/html-to-docx';
import fs from 'fs';

// Sample HTML content
const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <title>Welcome to html-to-docx</title>
</head>
<body style="font-family: Arial, sans-serif; margin: 20px; background-color: #f9f9f9;">
    <h1 style="color: #333; text-align: center;">Welcome to the html-to-docx library!</h1>

    <p style="line-height: 1.6; background-color: #ffff99; padding: 5px;">This library allows you to convert HTML content to DOCX format with ease.</p>
    <h2 style="color: #555;">Features</h2>
    <ul style="margin-left: 20px;">
        <li>Supports various HTML elements</li>
        <li>Handles styles and formatting</li>
        <li>Converts images and media</li>
        <li>Generates professional DOCX files</li>
    </ul>
    <page-break data-page-break="true" contenteditable="false" data-page-number="2"></page-break>
    <h2 style="color: #555;">Sample Content</h2>
    <p style="line-height: 1.6;">This is a paragraph with <strong>bold text</strong>, <em>italic text</em>, and <u>underlined text</u>.</p>
    <p style="line-height: 1.6;">Here's a list of benefits:</p>
    <ol>
        <li>Easy to use</li>
        <li>Fast conversion</li>
        <li>High quality output</li>
    </ol>
    <p style="line-height: 1.6;"><img style="max-width: 200px; height: auto; border: 1px solid #ccc;" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChAI/hxaB2tKUAAAAASUVORK5CYII=" alt="Sample Image" /></p>
    <p style="line-height: 1.6;">Thank you for using html-to-docx!</p>
</body>
</html>
`;

async function main() {
    try {
        const options = {
            fontSize: 12,
            fontFamily: 'Arial',
            lineHeight: 1.5,
            pageSize: 'A4', // Default page size (A4, Letter, Legal, or custom {width: number, height: number} in inches)
            marginTop: 0.6, // 1 inch top margin
            marginRight: 0.6, // 1 inch right margin
            marginBottom: 0.6, // 1 inch bottom margin
            marginLeft: 0.6, // 1 inch left margin
            marginHeader: 0, // 0 header from Top
            marginFooter: 0, // 0 footer from Bottom
            headerHeight: 0.6, // 1 inch header height
            footerHeight: 0.6, // 1 inch footer height
            enableHeader: true, // Enable/disable header (default: true if header content provided)
            enableFooter: true, // Enable/disable footer (default: true if footer content provided)
            enablePageNumbers: true, // Enable/disable page numbers
            pageNumberAlignment: 'center', // Page number alignment: 'left', 'center', or 'right'
            enablePageBorder: true, // Enable page border aligned with margins
            pageBorder: { style: 'single', color: '0000FF', size: 1, radius: 5 }, // Custom page border: blue single line of size 4
            header: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9ewAAAABJRU5ErkJggg==', // Blue colored image for header (positioned at top-left 0,0, full width)
            footer: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9ewAAAABJRU5ErkJggg==',  // Blue colored image for footer (positioned at bottom-left 0,bottom, full width)
            headingReplacements: [
                `<div data-h1 class="textbox" style="border: 1px solid #000000ff; border-radius: 5px; padding: 5px 5px 0px 5px; background-color: #000000ff; width:100%;">
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

        const docxBuffer = await convertHtmlToDocx(htmlContent, options);
        fs.writeFileSync('demo-output-test.docx', docxBuffer);
        console.log('DOCX file created: demo-output-test.docx');
    } catch (error) {
        console.error('Error converting HTML to DOCX:', error);
    }
}

main();
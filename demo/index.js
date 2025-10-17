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
        const docxBuffer = await convertHtmlToDocx(htmlContent);
        fs.writeFileSync('demo-output.docx', docxBuffer);
        console.log('DOCX file created: demo-output.docx');
    } catch (error) {
        console.error('Error converting HTML to DOCX:', error);
    }
}

main();
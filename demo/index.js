import { convertHtmlToDocx } from '../lib/index.js';
import fs from 'fs';

// Sample HTML content
const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <title>Sample Document</title>
    <style>
        body { font-family: Arial, sans-serif; }
        h1 { color: blue; }
        p { margin: 10px 0; }
    </style>
</head>
<body>
    <h1>Hello World</h1>
    <p>This is a sample HTML document to test the html-to-docx converter.</p>
    <p>It includes some basic styling and content.</p>
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
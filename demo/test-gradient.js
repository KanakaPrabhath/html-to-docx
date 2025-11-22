import { convertHtmlToDocx } from './lib/index.js';
import fs from 'fs';

// Test HTML with gradient backgrounds
const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <title>Gradient Test</title>
</head>
<body>
    <h1>Heading 1 - Should have black to white gradient</h1>
    <p>This is some content after heading 1.</p>
    
    <h2>Heading 2 - Should have gray to white gradient</h2>
    <p>This is some content after heading 2.</p>
    
    <h3>Heading 3 - Should have solid background</h3>
    <p>This is some content after heading 3.</p>
    
    <div class="textbox" style="border: 1px solid #ff0000; border-radius: 10px; padding: 10px; background: linear-gradient(to right, #ff0000, #0000ff); width:100%;">
        <p data-no-spacing style="color: #ffffff; margin: 0;">Direct gradient test: Red to Blue</p>
    </div>
    
    <div class="textbox" style="border: 2px solid #000000; border-radius: 5px; padding: 15px; background: linear-gradient(to bottom, #00ff00, #ffff00); width:100%;">
        <p data-no-spacing style="color: #000000; margin: 0; font-weight: bold;">Another gradient: Green to Yellow</p>
    </div>
</body>
</html>
`;

async function main() {
    try {
        const options = {
            fontSize: 12,
            fontFamily: 'Arial',
            lineHeight: 1.5,
            pageSize: 'A4',
            marginTop: 1,
            marginRight: 1,
            marginBottom: 1,
            marginLeft: 1,
            headingReplacements: [
                `<div data-h1 class="textbox" style="border: 1px solid #ffffffff; border-radius: 5px; padding: 0px 5px 0px 5px; background: linear-gradient(to right, #000000ff, #ffffffff); width:100%;">
<p data-no-spacing style="color: #ffffffff; margin: 0; font-size: 21px; font-weight: bold;">HEADING_TEXT</p>
</div>`,
                `<div data-h2 class="textbox" style="border: 1px dashed #050505ff; border-radius: 5px; padding: 0px 0px 0px 0px; background: linear-gradient(to bottom, #a1a1a1ff, #ffffffff); width:100%;">
<p data-no-spacing style="color: #000000ff; margin: 0; font-size: 19px; font-weight: bold;">HEADING_TEXT</p>
</div>`,
                `<div data-h3 class="textbox" style="border: 1px solid #00b118ff; border-radius: 5px; padding: 0px; background-color: #a10101ff; width:100%;">
<p data-no-spacing style="color: #ffffffff; margin: 0; font-size: 16px; font-weight: bold;">HEADING_TEXT</p>
</div>`
            ]
        };

        const docxBuffer = await convertHtmlToDocx(htmlContent, options);
        fs.writeFileSync('gradient-test-output.docx', docxBuffer);
        console.log('✓ Gradient test DOCX file created: gradient-test-output.docx');
        console.log('✓ H1: Black to White gradient (to right)');
        console.log('✓ H2: Gray to White gradient (to bottom)');
        console.log('✓ H3: Solid red background');
        console.log('✓ Direct gradient boxes: Red-Blue and Green-Yellow');
    } catch (error) {
        console.error('❌ Error converting HTML to DOCX:', error);
        process.exit(1);
    }
}

main();

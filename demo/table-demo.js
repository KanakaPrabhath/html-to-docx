import { convertHtmlToDocx } from '@kanaka-prabhath/html-to-docx';
import fs from 'fs';

// HTML table content with black solid borders
const htmlContent = `
<!DOCTYPE html>
<html>
<head>
    <title>Table Demo</title>
</head>
<body>
    <table style="border-collapse: collapse; border: 1px solid #000000; width: 100%;">
        <tbody>
            <tr style="height: 32px;">
                <th style="border: 1px solid #000000; padding: 5px; width: 86px;">Header 1</th>
                <th style="border: 1px solid #000000; padding: 5px; width: 86px;">Header 2</th>
                <td style="border: 1px solid #000000; padding: 5px; width: 100px;"></td>
                <td style="border: 1px solid #000000; padding: 5px; width: 295px;"></td>
            </tr>
            <tr style="height: 67px;">
                <td style="border: 1px solid #000000; padding: 5px;"></td>
                <td style="border: 1px solid #000000; padding: 5px;"></td>
                <td style="border: 1px solid #000000; padding: 5px;"></td>
                <td style="border: 1px solid #000000; padding: 5px;"></td>
            </tr>
            <tr style="height: 75px;">
                <td style="border: 1px solid #000000; padding: 5px;">Cell 3</td>
                <td style="border: 1px solid #000000; padding: 5px;">Cell 4</td>
                <td style="border: 1px solid #000000; padding: 5px;"></td>
                <td style="border: 1px solid #000000; padding: 5px;"></td>
            </tr>
            <tr style="height: 118px;">
                <td style="border: 1px solid #000000; padding: 5px;"></td>
                <td style="border: 1px solid #000000; padding: 5px;"></td>
                <td style="border: 1px solid #000000; padding: 5px;"></td>
                <td style="border: 1px solid #000000; padding: 5px;"></td>
            </tr>
        </tbody>
    </table>
    <p><br></p>
</body>
</html>
`;

async function main() {
    try {
        const options = {
            fontSize: 11,
            fontFamily: 'Calibri',
            lineHeight: 1.15,
            pageSize: 'A4',
            marginTop: 1.0,
            marginRight: 1.0,
            marginBottom: 1.0,
            marginLeft: 1.0,
        };

        const docxBuffer = await convertHtmlToDocx(htmlContent, options);
        fs.writeFileSync('table-output.docx', docxBuffer);
        console.log('DOCX file created: table-output.docx');
        console.log('Table with black solid borders and correct dimensions has been generated.');
    } catch (error) {
        console.error('Error converting HTML to DOCX:', error);
    }
}

main();

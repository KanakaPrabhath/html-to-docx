import { convertHtmlToDocx } from '@kanaka-prabhath/html-to-docx';
import fs from 'fs';

// Sample HTML content
const htmlContent = `
  <!DOCTYPE html>
    <html>
    <head>
        <title>Comprehensive Formatting Test</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
        </style>
    </head>
    <body>
        <p><b>It is a long established fact that a </b></p>
        <p><i>reader will be distracted by the readable </i></p>
        <p><u>content of a page when looking at its layout. </u></p>
        <p><strike>The point of using Lorem Ipsum is that it has a</strike></p>
        <p><b><i>more-or-less normal distribution of letters, as opposed to using </i></b></p>
        <p><i><u>'Content here, content here', </u></i></p>
        <p>making it look like readable English. Many </p>
        <p style="text-align: justify;">desktop publishing packages and web page editors now use Lorem Ipsum as their default model text, and a search for 'lorem ipsum' will uncover many web sites still in their infancy. Various versions have evolved over the years, sometimes by accident, sometimes on purpose (injected humour and the like).</p>
        <p style="text-align: justify;"><br></p>
        <p style="text-align: left;">Contrary to popular belief, </p>
        <p style="text-align: center;">Lorem Ipsum is not simply random text. </p>
        <p style="text-align: right;">It has roots in a piece of classical Latin literature from 45 BC,</p>
        <p><span style="font-family: 'Times New Roman';"> making it over 2000 years old. Richard </span></p>
        <p><span style="font-family: Arial;">McClintock, a Latin professor at Hampden-Sydney College in Virginia, </span></p>
        <p><span style="font-family: 'Iskoola Pota';">looked up one of the more obscure Latin words, consectetur, </span></p>
        <p><span style="font-family: 'Iskoola Pota';">සාහිතයය නිර්මාණයක් පරිවර්තනය කිරීදමන් බොදපාදරාත්තු වන්දන් ලියැ ී ඇති එක් භාෂාවකින් කියවා</span></p>
        <p><font face="Iskoola Pota"><br></font><span style="font-size: 27px;"><span style="font-family: 'Iskoola Pota';">අවදබෝධ්‍ කර ගත දනා හැකි පාඨක සමූහයකට මුහුණ නිර්මාණය ආදද්ශකයක් තමන්දේ බසන් ඉිරිපත් කිරීමකි.</span></span><br><br></p>
        <p><span style="font-size: 15px;">සාහිතයය වශදයන් ගත් කෙ පරිවර්තනය ෙ කො කාර්යයකි.</span></p>
        <page-break data-page-break="true" contenteditable="false" data-page-number="2"></page-break>
        <h1>There are many variations of passages </h1>
        <h2>of Lorem Ipsum available, but the majority have suffered </h2>
        <h3>alteration in some form, by injected humour, </h3>
        <h4>or randomised words which don't look even slightly believable. </h4>
        <p></p>
        <ul>
          <li>If you are going to use a passage of Lorem Ipsum, </li>
          <li>you need to be sure there isn't </li>
          <li>anything embarrassing hidden</li>
          <li> in the middle of text. All the Lorem </li>
        </ul>
        <ol>
          <li>Ipsum generators on the Internet </li>
          <li>tend to repeat predefined chunks as necessary, </li>
          <li>making this the first true generator on the Internet. </li>
          <li>It uses a dictionary of over 200 Latin words, </li>
        </ol>
        <p></p>
        <p>combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable.</p>
        <p><br></p>
        <p><img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR42mNkYPhfDwAChAI/hxaB2tKUAAAAASUVORK5CYII=" alt="Test Image" width="100" height="100" /></p>
        <p><br></p>
        <p data-indent-level="3" style="margin-left: 96px;">Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots in a piece of classical Latin literature from 45 BC, </p>
        <p data-indent-level="1" style="margin-left: 32px;">making it over 2000 years old. Richard McClintock, a Latin professor at Hampden-Sydney College in Virginia, looked up one of the more obscure Latin words, consectetur,</p>
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
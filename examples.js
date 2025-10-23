/**
 * Example integration of HTML to DOCX converter in TuteMaker App
 * This file shows how to use the html-to-docx library in the Electron main process
 */

const HtmlToDocx = require('./index');
const path = require('path');

/**
 * Example 1: Basic conversion from HTML editor content
 */
async function exportEditorContent(htmlContent, outputPath) {
  const converter = new HtmlToDocx({
    fontSize: 11,
    fontFamily: 'Calibri',
    lineHeight: 1.15
  });

  try {
    await converter.convertHtmlToDocxFile(htmlContent, outputPath);
    console.log('Document exported successfully to:', outputPath);
    return { success: true, path: outputPath };
  } catch (error) {
    console.error('Export failed:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Example 2: Convert to buffer (for further processing or sending)
 */
async function convertToBuffer(htmlContent) {
  const converter = new HtmlToDocx();

  try {
    const buffer = await converter.convertHtmlToDocx(htmlContent);
    console.log('Document converted to buffer, size:', buffer.length, 'bytes');
    return { success: true, buffer };
  } catch (error) {
    console.error('Conversion failed:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Example 3: IPC Handler for Electron
 * Add this to electron/handlers/documentHandlers.js
 */
function setupDocxExportHandler(ipcMain) {
  ipcMain.handle('export-html-to-docx', async (event, { html, outputPath, options = {} }) => {
    try {
      const converter = new HtmlToDocx({
        fontSize: options.fontSize || 11,
        fontFamily: options.fontFamily || 'Calibri',
        lineHeight: options.lineHeight || 1.15
      });

      await converter.convertHtmlToDocxFile(html, outputPath);

      return {
        success: true,
        message: 'Document exported successfully',
        path: outputPath
      };
    } catch (error) {
      console.error('DOCX export error:', error);
      return {
        success: false,
        error: error.message
      };
    }
  });
}

/**
 * Example 4: Usage from React component
 * Call from your React app (e.g., HtmlEditor component)
 */
async function exportFromReact() {
  // In your React component:
  const handleExport = async () => {
    const htmlContent = editorContent; // Your HTML editor content
    
    // Show save dialog
    const result = await window.electronAPI.showSaveDialog({
      title: 'Export as Word Document',
      defaultPath: 'document.docx',
      filters: [
        { name: 'Word Documents', extensions: ['docx'] }
      ]
    });

    if (result.canceled) return;

    // Export to DOCX
    const exportResult = await window.electronAPI.exportHtmlToDocx({
      html: htmlContent,
      outputPath: result.filePath,
      options: {
        fontSize: 12,
        fontFamily: 'Calibri'
      }
    });

    if (exportResult.success) {
      alert('Document exported successfully!');
    } else {
      alert(`Export failed: ${exportResult.error}`);
    }
  };
}

/**
 * Example 5: Batch conversion
 */
async function batchConvertDocuments(documents, outputDir) {
  const converter = new HtmlToDocx();
  const results = [];

  for (const doc of documents) {
    const outputPath = path.join(outputDir, `${doc.name}.docx`);
    
    try {
      await converter.convertHtmlToDocxFile(doc.html, outputPath);
      results.push({ name: doc.name, success: true, path: outputPath });
    } catch (error) {
      results.push({ name: doc.name, success: false, error: error.message });
    }
  }

  return results;
}

/**
 * Example 6: Custom styling based on template
 */
async function exportWithTemplate(htmlContent, template) {
  const templateStyles = {
    'default': { fontSize: 11, fontFamily: 'Calibri', lineHeight: 1.15 },
    'large-print': { fontSize: 14, fontFamily: 'Arial', lineHeight: 1.5 },
    'compact': { fontSize: 10, fontFamily: 'Times New Roman', lineHeight: 1.0 }
  };

  const styles = templateStyles[template] || templateStyles['default'];
  const converter = new HtmlToDocx(styles);

  return await converter.convertHtmlToDocx(htmlContent);
}

/**
 * Example 7: Page border
 */
async function exportWithPageBorder(htmlContent, outputPath) {
  const converter = new HtmlToDocx({
    enablePageBorder: true,
    pageBorder: {
      style: 'double',
      color: 'FF0000',
      size: 8,
      radius: 5 // Note: Not supported in OOXML page borders
    }
  });

  return await converter.convertHtmlToDocxFile(htmlContent, outputPath);
}

module.exports = {
  exportEditorContent,
  convertToBuffer,
  setupDocxExportHandler,
  batchConvertDocuments,
  exportWithTemplate,
  exportWithPageBorder
};

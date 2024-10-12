const express = require('express');
const serverless = require('serverless-http');
const multer = require('multer');
const { PDFDocument } = require('pdf-lib');
const sharp = require('sharp');
const fs = require('fs').promises;
const path = require('path');
const { fromPath } = require('pdf2pic');
const pdf = require('pdf-parse');
const Excel = require('exceljs');
const { Document, Packer, Paragraph } = require('docx');

const app = express();
const upload = multer({ dest: '/tmp/uploads/' });

// Root route
app.get('/', (req, res) => {
  res.send('PDF Conversion API is running. Use POST /convert to convert files.');
});

app.post('/convert', upload.single('pdf'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('No file uploaded.');
    }

    console.log('File received:', req.file);

    const inputPath = req.file.path;
    const outputPath = path.join('/tmp', `converted_${Date.now()}`);
    const targetFormat = req.body.format || 'pdf'; // Default to PDF if no format specified

    console.log('Converting to format:', targetFormat);

    let result;
    switch (targetFormat.toLowerCase()) {
      case 'pdf':
        result = await convertPdfToPdf(inputPath, `${outputPath}.pdf`);
        break;
      case 'word':
        result = await convertPdfToWord(inputPath, `${outputPath}.docx`);
        break;
      case 'excel':
        result = await convertPdfToExcel(inputPath, `${outputPath}.xlsx`);
        break;
      case 'jpeg':
      case 'jpg':
      case 'png':
        result = await convertPdfToImage(inputPath, `${outputPath}.${targetFormat}`, targetFormat);
        break;
      case 'html':
        result = await convertPdfToHtml(inputPath, `${outputPath}.html`);
        break;
      default:
        return res.status(400).send('Unsupported format');
    }

    console.log('Conversion completed. Result:', result);

    res.download(result, path.basename(result), (err) => {
      if (err) {
        console.error('Error sending file:', err);
        res.status(500).send('Error sending file');
      }
      // Clean up temporary files
      fs.unlink(inputPath).catch(err => console.error('Error deleting input file:', err));
      fs.unlink(result).catch(err => console.error('Error deleting output file:', err));
    });
  } catch (error) {
    console.error('Error in /convert route:', error);
    res.status(500).send(`An error occurred during conversion: ${error.message}`);
  }
});

async function convertPdfToPdf(input, output) {
  console.log('Converting PDF to PDF');
  const pdfDoc = await PDFDocument.load(await fs.readFile(input));
  const pdfBytes = await pdfDoc.save();
  await fs.writeFile(output, pdfBytes);
  console.log('PDF to PDF conversion completed');
  return output;
}

async function convertPdfToWord(input, output) {
  console.log('Converting PDF to Word');
  const data = await pdf(await fs.readFile(input));
  const doc = new Document({
    sections: [{
      properties: {},
      children: [new Paragraph(data.text)]
    }]
  });
  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(output, buffer);
  console.log('PDF to Word conversion completed');
  return output;
}

async function convertPdfToExcel(input, output) {
  console.log('Converting PDF to Excel');
  const data = await pdf(await fs.readFile(input));
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  worksheet.addRow([data.text]);
  await workbook.xlsx.writeFile(output);
  console.log('PDF to Excel conversion completed');
  return output;
}

async function convertPdfToImage(input, output, format) {
  console.log('Converting PDF to Image');
  const options = {
    density: 300,
    saveFilename: path.basename(output),
    savePath: path.dirname(output),
    format: format.toUpperCase(),
    width: 2480,
    height: 3508
  };
  const convert = fromPath(input, options);
  const result = await convert(1); // Convert first page
  console.log('PDF to Image conversion completed');
  return result.path;
}

async function convertPdfToHtml(input, output) {
  console.log('Converting PDF to HTML');
  const data = await pdf(await fs.readFile(input));
  const html = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Converted PDF</title>
</head>
<body>
    <pre>${data.text}</pre>
</body>
</html>`;
  await fs.writeFile(output, html);
  console.log('PDF to HTML conversion completed');
  return output;
}

module.exports.handler = serverless(app);
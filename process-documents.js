const fs = require('fs');
const path = require('path');
const pdfjs = require('pdfjs-dist/legacy/build/pdf');
const { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, VerticalAlign, convertInchesToTwip } = require('docx');

async function extractTextFromPdf(filePath) {
  const pdfBuffer = fs.readFileSync(filePath);
  const pdf = await pdfjs.getDocument({ data: pdfBuffer }).promise;

  let fullText = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map(item => item.str).join('');
    fullText += pageText + '\n';
  }

  return fullText;
}

function replaceName(text) {
  // Replace variations of the name
  let result = text
    .replace(/ИП Чеусова Николая ИГоревича/g, 'ИП Пугачеву Ольгу Николаевну')
    .replace(/ИП Чеусова Николая Игоревича/g, 'ИП Пугачеву Ольгу Николаевну')
    .replace(/Чеусова Николая ИГоревича/g, 'Пугачеву Ольгу Николаевну')
    .replace(/Чеусова Николая Игоревича/g, 'Пугачеву Ольгу Николаевну')
    .replace(/Чеусова/g, 'Пугачева')
    .replace(/Николай Игоревич/g, 'Ольга Николаевна')
    .replace(/Николай ИГоревич/g, 'Ольга Николаевна');

  return result;
}

function createWordDocument(text) {
  // Split text into paragraphs
  const paragraphs = text.split('\n').map(line => {
    return new Paragraph({
      text: line.trim() || ' ',
      spacing: { line: 240, lineRule: 'auto' },
    });
  });

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margins: {
            top: convertInchesToTwip(1),
            right: convertInchesToTwip(1),
            bottom: convertInchesToTwip(1),
            left: convertInchesToTwip(1),
          },
        },
      },
      children: paragraphs
    }]
  });

  return doc;
}

async function main() {
  try {
    const uploadDir = '/root/.claude/uploads/3e540a7b-5ab9-5cc8-9677-e3d80f77fa1e';
    const files = fs.readdirSync(uploadDir).filter(f => f.endsWith('.pdf')).sort();

    if (files.length < 2) {
      console.error('Error: Expected at least 2 PDF files');
      console.error('Found files:', files);
      return;
    }

    console.log('Processing PDF files...');
    console.log('File 1:', files[0]);
    console.log('File 2:', files[1]);

    // Extract text from both PDFs
    const filePath1 = path.join(uploadDir, files[0]);
    const filePath2 = path.join(uploadDir, files[1]);

    console.log('Extracting text from first PDF...');
    const text1 = await extractTextFromPdf(filePath1);
    console.log('✓ Extracted', text1.length, 'characters');

    console.log('Extracting text from second PDF...');
    const text2 = await extractTextFromPdf(filePath2);
    console.log('✓ Extracted', text2.length, 'characters');

    // Replace names
    console.log('Replacing names...');
    const replacedText1 = replaceName(text1);
    const replacedText2 = replaceName(text2);

    // Create Word documents
    console.log('Creating Word documents...');
    const doc1 = createWordDocument(replacedText1);
    const doc2 = createWordDocument(replacedText2);

    // Save documents
    const output1 = '/home/user/marketingskills/Договор_Пугачева.docx';
    const output2 = '/home/user/marketingskills/Реквизиты_Пугачева.docx';

    await Packer.toFile(doc1, output1);
    await Packer.toFile(doc2, output2);

    console.log(`✓ Created: ${output1}`);
    console.log(`✓ Created: ${output2}`);
    console.log('\n✅ Processing complete!');

  } catch (error) {
    console.error('Error:', error.message);
    console.error(error.stack);
  }
}

main();

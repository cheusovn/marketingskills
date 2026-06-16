const fs = require('fs');
const pdfParse = require('pdf-parse');

const filePath = '/root/.claude/uploads/3e540a7b-5ab9-5cc8-9677-e3d80f77fa1e';
const files = fs.readdirSync(filePath).filter(f => f.endsWith('.pdf'));
console.log('PDF files:', files);

async function test() {
  try {
    const testFile = fs.readFileSync(filePath + '/' + files[0]);
    const data = await pdfParse(testFile);
    console.log('✓ Text extracted, length:', data.text.length);
    console.log('\nFirst 600 characters:\n');
    console.log(data.text.substring(0, 600));
  } catch (e) {
    console.error('Error:', e.message);
  }
}

test();

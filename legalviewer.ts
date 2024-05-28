import * as fs from 'fs';
import * as https from 'https';
import * as dotenv from 'dotenv';
import * as docx from 'docx';
import { Document,VerticalAlign, Packer, Paragraph, Table, TableCell, TableRow, Header, Footer, WidthType, AlignmentType } from 'docx';


// Load konfigurasi dari file .env
dotenv.config();

// Mendapatkan variabel konfigurasi dari file .env
const OPENSEARCH_ADDRESS:string = process.env.OPENSEARCH_ADDRESS;
const OPENSEARCH_PORT:string = process.env.OPENSEARCH_PORT;
const OPENSEARCH_USERNAME:string = process.env.OPENSEARCH_USERNAME;
const OPENSEARCH_PASSWORD:string = process.env.OPENSEARCH_PASSWORD;

// Nama indeks OpenSearch
const INDEX_NAME:string = 'law_analyzer_msib';

// Mengambil ID dari argumen command line
const args : string[]= process.argv.slice(2);
const idIndex:number = args.indexOf('-id');
if (idIndex === -1 || idIndex === args.length - 1) {
  console.error("Error: Please provide an ID using the '-id' argument.");
  process.exit(1);
}
let idValue: string = args[idIndex + 1];

// Menghapus tanda petik dari nilai ID
idValue = idValue.replace(/^'|'$/g, '');

// Konfigurasi koneksi ke cluster OpenSearch
const options = {
  hostname: OPENSEARCH_ADDRESS,
  port: OPENSEARCH_PORT,
  path: `/${INDEX_NAME}/_search`,
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': 'Basic ' + Buffer.from(`${OPENSEARCH_USERNAME}:${OPENSEARCH_PASSWORD}`).toString('base64')
  },
  rejectUnauthorized: false // Setel ini ke false
};

// Query pencarian untuk mencari dokumen berdasarkan ID dengan mengecualikan beberapa field
const searchQuery: any = {
  _source: {
    excludes: [
      "Blocks.ContentText.MainVector", "Blocks.ContentText.AdditionalContext.Vector"
    ]
  },
  query: {
    ids: {
      values: [idValue]
    }
  }
};

// Body permintaan dengan query pencarian
const postData: string = JSON.stringify(searchQuery);

async function fetchOpensearchData(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
    // Buat permintaan HTTPS
    const req = https.request(options, (res) => {
      console.log(`Status Code: ${res.statusCode}`);
      let data = '';

      res.on('data', (chunk) => {
        data += chunk;
      });

      res.on('end', () => {
        try {
          const response = JSON.parse(data);
          // Simpan respons ke dalam file JSON
          fs.writeFileSync('D:/magang/json_ke_word/response.json', JSON.stringify(response, null, 2));
          console.log('Response saved to response.json');
          resolve();
        } catch (error) {
          console.error('Error parsing JSON response:', error);
          reject(error);
        }
      });
    });

    req.on('error', (error) => {
      console.error(error);
      reject(error);
    });

    // Kirim data
    req.write(postData);
    req.end();
  });
}

async function parseAndModifyData(): Promise<void> {
    return new Promise((resolve, reject) => {
      // Read the JSON file
      fs.readFile('D:/magang/json_ke_word/response.json', 'utf8', (err, data) => {
        if (err) {
          console.error('Error reading file:', err);
          reject(err);
          return;
        }
        try {
          // Parse JSON
          const jsonData = JSON.parse(data);
  
          // Array to store modified data
          let modifiedData: Array<any> = [];
  
          // Variable to store the previous chapter value
          let previousBab: string | null = null;
  
          // Process modification of data
          jsonData.hits.hits.forEach((item: any) => {
            item._source.Blocks.forEach((block: any) => {
              if (block.ContentText && Array.isArray(block.ContentText)) {
                let combinedContent = "";
                block.ContentText.forEach((content: any) => {
                  // Check if content is an object
                  if (typeof content === 'object' && content !== null) {
                    // Modify text content
                    let modifiedValue = content.Value.replace(/dimaksud pada ayat \((\d+)\)/g, 'dimaksud pada ayat $1');
                    modifiedValue = modifiedValue.replace(/dimaksud dalam Pasal (\d+) ayat \((\d+)\)/g, 'dimaksud dalam Pasal $1 ayat $2');
                    let modifiedRef = content.Ref ? content.Ref.replace(/Ayat \((\d+)\)/g, '($1)') : null;
                    // Remove patterns like -number-
                    modifiedValue = modifiedValue.replace(/-\d+-/g, '');
                    // Add new patterns
                    modifiedValue = modifiedValue.replace(/(\d+)\. /g, '($1) ');
                    modifiedValue = modifiedValue.replace(/\. \((\d+)\)/g, '.\n($1)');
                    combinedContent += `${modifiedRef ? `${modifiedRef} ` : ''}${modifiedValue}\n`;
                  }
                });
  
                // Set chapter value according to the required logic
                let modifiedBab = block.Bab;
                if (modifiedBab === previousBab) {
                  modifiedBab = null; // Set to null if the same as the previous chapter value
                } else {
                  previousBab = block.Bab; // Update the previous chapter value
                }
  
                modifiedData.push({
                  bab: modifiedBab,
                  judulbab: block.BabContext,
                  bagian: block.Bagian,
                  paragraf: block.Paragraf,
                  pasal: block.Pasal ? `pasal-${block.Pasal.split(' ')[1]}` : null, 
                  ref: null,
                  type: "CONTENT_PASAL",
                  content: combinedContent.trim(),
                  additional_context: [],
                  context: item._source.Judul
                });
              }
            });
          });
  
          // Write back to a JSON file
          fs.writeFile('D:/magang/json_ke_word/new_file/typeS/material/material.json', JSON.stringify(modifiedData, null, 2), 'utf8', (err) => {
            if (err) {
              console.error('Error writing file:', err);
              reject(err);
              return;
            }
            console.log('File material.json has been successfully saved.');
            resolve();
          });
  
        } catch (error) {
          console.error('Error parsing JSON:', error);
          reject(error);
        }
      });
    });
  }

async function generateWordDocument(): Promise<void> {
  function readJsonFile(filePath: string): any {
    const jsonData: string = fs.readFileSync(filePath, 'utf8');
    return JSON.parse(jsonData);
  }


  interface TableCellProps {
    children: Paragraph[];
    width: { size: number; type: WidthType };
    verticalAlign?: VerticalAlign;
    shading: { fill: string };
  }
  
  interface TableRowProps {
    children: TableCell[];
    tableHeader?: boolean;
  }
  
  function createTable(data: any): TableRow[] {
    let numberingInstance: number = 0;
    let alphabeticalNumberingInstance: number = -1;
  
    const tableRows: TableRow[] = [
      new TableRow({
        children: [
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new docx.TextRun({ text: 'NO.', bold: true, font: "Bookman Old Style", size: 22 })
                ],
                alignment: AlignmentType.CENTER
              })
            ],
            width: { size: 5, type: WidthType.PERCENTAGE },
            shading: { fill: 'd9e2f3' }
          }),
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new docx.TextRun({ text: 'SAAT INI', bold: true, font: "Bookman Old Style", size: 22 })
                ],
                alignment: AlignmentType.CENTER
              })
            ],
            width: { size: 35, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: 'd9e2f3' }
          }),
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new docx.TextRun({ text: 'PERUBAHAN', color: '#4472c4', bold: true, font: "Bookman Old Style", size: 22 })
                ],
                alignment: AlignmentType.CENTER
              })
            ],
            width: { size: 35, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: 'd9e2f3' }
          }),
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new docx.TextRun({ text: 'KETERANGAN', bold: true, font: "Bookman Old Style", size: 22 })
                ],
                alignment: AlignmentType.CENTER
              })
            ],
            width: { size: 25, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: 'd9e2f3' }
          })
        ],
        tableHeader: true
      }),


      //................
      ...data.map((item, index) => {
        const cleanedContent: string = item.content.replace(/;/g, '.');
    const numberingReference: string = `my-numbering-${index + 1}`;
    const otherNumberingReference: string = `my-other-numbering-${index + 1}`;
    let numberingInstance: number = 0;
    let alphabeticalNumberingInstance: number = -1;

    const paragraphs: string[] = item.content.split(/(?=[0-9a-zA-Z]\. |\([0-9]+\)\s)/gm).map(contentPart => contentPart.trim());
    const contentParagraphs: Paragraph[] = [];

    paragraphs.forEach((contentPart) => {
      let numbering: any = undefined;

      if (contentPart.match(/^\(\d+\)\s/)) {
        numberingInstance++;
        alphabeticalNumberingInstance = 0;
        numbering = {
          reference: numberingReference,
          level: 0,
          format: docx.LevelFormat.BULLET,
          text: `${numberingInstance}`,
        };
        contentPart = contentPart.replace(/^\(\d+\)\s/, '');
      } else if (contentPart.match(/^[a-z]\.\s/)) {
        if (alphabeticalNumberingInstance === -1) {
          alphabeticalNumberingInstance = 0;
        }
        const charCode: number = 'a'.charCodeAt(0) + alphabeticalNumberingInstance;
        alphabeticalNumberingInstance++;
        numbering = {
          reference: otherNumberingReference,
          level: 0,
          format: docx.LevelFormat.LOWER_LETTER,
          text: `${String.fromCharCode(charCode)}.`,
        };
        contentPart = contentPart.replace(/^[a-z]\.\s/, '');
      }

      const lastChar: string = contentPart.charAt(contentPart.length - 1);
      if (!lastChar.match(/[a-zA-Z0-9]/)) {
        contentPart = contentPart.slice(0, -1);
      }

      contentParagraphs.push(new Paragraph({
        children: [new docx.TextRun({ text: contentPart, font: "Bookman Old Style", size: 22 })],
        numbering,
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 100, before: 100 },
        indent: { left: 400, right: 100 },
      }));
    });

    const pasalParagraph: Paragraph = new Paragraph({
      children: [new docx.TextRun({ text: item.pasal ? `${item.pasal.replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase())}` : '', bold: true, font: "Bookman Old Style", size: 22 })],
      alignment: AlignmentType.CENTER,
    });

    const babParagraph: Paragraph | null = item.bab ? new Paragraph({
      children: [
        new docx.TextRun({ text: item.bab, bold: true, font: "Bookman Old Style", size: 22 }),
        new docx.TextRun({ text: item.judulbab, break: 1, bold: true, font: "Bookman Old Style", size: 22 }),
      ],
      alignment: AlignmentType.CENTER,
    }) : null;

    return [
      babParagraph ? new TableRow({
        children: [
          new TableCell({ children: [new Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER }),
          new TableCell({ children: [babParagraph], verticalAlign: docx.VerticalAlign.CENTER, shading: { fill: 'F8E8EE' } }),
          new TableCell({ children: [new Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER }),
          new TableCell({ children: [new Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER }),
        ],
      }) : null,
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ text: String(index + 1), alignment: AlignmentType.CENTER, font: "Bookman Old Style", size: 22 })] }),
          new TableCell({ children: [pasalParagraph, ...contentParagraphs], alignment: AlignmentType.JUSTIFIED }),
          new TableCell({ children: [new Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER }),
          new TableCell({ children: [new Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER }),
        ],
      }),
    ].filter(Boolean);
  })].flat();

  const table: Table = new Table({ rows: tableRows });
  return table;
}

function createWordDocument(data: any): Document {
    const jsonData = data[0];
    const context: string = jsonData.context;
    const words: string[] = context.split(' ');
    const firstThreeWords: string = words.slice(0, 5).join(' ');
  
    const paragraph: Paragraph = new Paragraph({
      children: [new docx.TextRun({ text: firstThreeWords, bold: true, font: "Bookman Old Style", size: 22 })],
      alignment: AlignmentType.CENTER,
    });
  
    if (words.length > 5) {
      const remainingWords: string = words.slice(5).join(' ');
      paragraph.addChildElement(new docx.TextRun({ text: '\n' + remainingWords, bold: true, size: 22, font: "Bookman Old Style", break: 1 }));
    }
  
    const content: Table = createTable(data);
    const children: (Paragraph | Table)[] = [paragraph, content];
  
    const numberingConfig: any[] = data.map((item: any, index: number) => ({
      reference: `my-numbering-${index + 1}`,
      levels: [{ level: 0, format: docx.LevelFormat.DECIMAL, text: "(%1)" }],
    })).concat(data.map((item: any, index: number) => ({
      reference: `my-other-numbering-${index + 1}`,
      levels: [{ level: 0, format: docx.LevelFormat.LOWER_LETTER, text: "%1." }],
    })));
  
    const document: Document = new Document({
      numbering: { config: numberingConfig },
      sections: [{
        properties: {
          page: {
            size: { orientation: docx.PageOrientation.LANDSCAPE },
            margin: { top: 720, right: 720, bottom: 720, left: 720 },
          },
        },
        headers: { default: new Header({ children: [new Paragraph("Header placement")], properties: { footer: { marginTop: 50, marginBottom: 50 } } }) },
        footers: { default: new Footer({ children: [new Paragraph("Generated with https://github.com/gagahputrabangsa/Legal_Viewer_Msib.git")], properties: { footer: { marginTop: 50, marginBottom: 50 } } }) },
        children: children,
      }],
    });
  
    return document;
  }

  const jsonFilePath: string = 'D:/magang/json_ke_word/new_file/typeS/material/material.json';
  const jsonData = readJsonFile(jsonFilePath);
  const document = createWordDocument(jsonData);

  Packer.toBuffer(document).then((buffer) => {
    const args = process.argv.slice(2);
    const outputArgIndex = args.indexOf('-out');
    let outputPath = 'output.docx';

    if (outputArgIndex !== -1 && args[outputArgIndex + 1]) {
      outputPath = args[outputArgIndex + 1];
    } else {
      console.error('Error: Output path not specified. Use -out <outputPath> to specify the output file.');
      process.exit(1);
    }

    fs.writeFileSync(outputPath, buffer);
    console.log(`Dokumen Word berhasil dan disimpan di ${outputPath}`);
    console.log
(`
'YEAY'    
`);
});
}

// Main function
async function main() {
  try {
    await fetchOpensearchData();
    await parseAndModifyData();
    await generateWordDocument();
  } catch (error) {
    console.error('Error during processing:', error);
  }
}

// Run main function
main();

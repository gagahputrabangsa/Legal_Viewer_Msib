const fs = require('fs');
const docx = require('docx');
const { Document, Packer, Paragraph, Table, TableCell, TableRow, Header, Footer, AlignmentType } = docx;

// Function to read JSON file
function readJsonFile(filePath) {
    const jsonData = fs.readFileSync(filePath, 'utf8');
    return JSON.parse(jsonData);
}

// Function to create table based on JSON data
function createTable(data) {
    let numberingInstance = 0; // Initialize numberingInstance outside of the map function

    // Map through JSON data to create rows
    const tableRows = [
        new TableRow({
            children: [ 
                // Cell for 'NO.'
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new docx.TextRun({ text: 'NO.', bold: true, font: "Orbi", fontSize: 12 }),
                            ],
                            alignment: AlignmentType.CENTER,
                        }),
                    ],
                    shading: { fill: 'd9e2f3' },
                }),
                // Cell for 'SAAT INI'
                
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new docx.TextRun({ text: 'SAAT INI',
                                 bold: true,
                                  font: "Orbi", 
                                   fontSize: 12 }),
                            ],
                            alignment: AlignmentType.CENTER,
                        }),
                    ],
                    verticalAlign: docx.VerticalAlign.CENTER,
                    shading: { fill: 'd9e2f3' },
                }),
                // Cell for 'PERUBAHAN'
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new docx.TextRun({ text: 'PERUBAHAN', color: '#4472c4', bold: true, font: "Orbi", fontSize: 12 }),
                            ],
                            alignment: AlignmentType.CENTER,
                            indent: { left: 200, right: 200 },
                        }),
                    ],
                    verticalAlign: docx.VerticalAlign.CENTER,
                    shading: { fill: 'd9e2f3' },
                }),
                // Cell for 'KETERANGAN'
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new docx.TextRun({ text: 'KETERANGAN', bold: true, font: "Orbi", fontSize: 12 }),
                            ],
                            alignment: AlignmentType.CENTER,
                            indent: { left: 200, right: 200 },
                        }),
                    ],
                    verticalAlign: docx.VerticalAlign.CENTER,
                    shading: { fill: 'd9e2f3' },
                }),
            ],
            //==========IMPORTANT
            tableHeader: true,
        }),
    ...data.map((item, index) => {
        const cleanedContent = item.content.replace(/;/g, '.');
        // Generate dynamic reference names for numbering
        const numberingReference = `my-numbering-${index + 1}`;
        const otherNumberingReference = `my-other-numbering-${index + 1}`;
        // Reset numberingInstance at the beginning of each iteration
        numberingInstance = 0;
        
        const paragraphs = item.content.split(/(?=[0-9a-zA-Z]\. |\([0-9]+\)\s)/gm)
        .map(contentPart => contentPart.trim());
        const contentParagraphs = [];
        
        paragraphs.forEach((contentPart, paragraphIndex) => {
            let numbering = undefined;
            
            if (contentPart.match(/^\(\d+\)\s/)) {
                numberingInstance++;
                numbering = {
                        reference: numberingReference,
                        level: 0,
                        format: docx.LevelFormat.BULLET,
                        text: `${numberingInstance}`,
                    }; 

                    contentPart = contentPart.replace(/^\(\d+\)\s/, ''); 
                } else if (contentPart.match(/^[a-z]\.\s/)) {
                    const charCode = 97 + (numberingInstance - 1);
                    numbering = {
                        reference: otherNumberingReference,
                        level: 0,
                        format: docx.LevelFormat.BULLET,
                        text: `${String.fromCharCode(charCode)}.`,
                    };
                    contentPart = contentPart.replace(/^[a-z]\.\s/, '');
                }

                // Remove non-letter characters from the end of the sentence
                const lastChar = contentPart.charAt(contentPart.length - 1);
                if (!lastChar.match(/[a-zA-Z0-9]/)) {
                    contentPart = contentPart.slice(0, -1); // Remove the last character
                }

                contentParagraphs.push(new Paragraph({
                    children: [
                        new docx.TextRun({
                            text: contentPart,
                            font: "Orbi", // Add font property
                            fontSize: 12, // Add fontSize property
                        }),
                    ],
                    numbering,
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { after: 100, before: 100 },
                    indent: { left: 400, right: 100 },
                }));
            });

            const pasalAndBabParagraph = new Paragraph({
                children: [
                    new docx.TextRun({
                        text: item.bab ? `Bab ${item.bab}` : '',
                        bold: true,
                        // break: 1,
                        font: "Orbi", // Add font property
                        fontSize: 12, // Add fontSize property
                    }),
                    new docx.TextRun({
                        text: item.bab ? '\nKetentuan Umum\n' : '',
                        bold: true,
                        break: 1,
                        font: "Orbi", // Add font property
                        fontSize: 12, // Add fontSize property
                    }),
                    new docx.TextRun({
                        text: item.pasal ? `${item.pasal.replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase())}` : '',
                        bold: true,
                        break: 1,
                        font: "Orbi", // Add font property
                        fontSize: 12, // Add fontSize property
                    }),
                ],
                alignment: AlignmentType.CENTER,
            });
            

            return [
                // Row for 'SAAT INI' and 'PERUBAHAN'
                new TableRow({
                    children: [
                        // Cell for 'NO.'
                        new TableCell({
                            children: [new Paragraph({ text: String(index + 1), alignment: AlignmentType.CENTER, font: "Orbi", fontSize: 12 })], // Add font and fontSize properties
                            alignment: AlignmentType.CENTER,
                        }),
                        
                        // Cell for 'SAAT INI' and 'PERUBAHAN'
                        new TableCell({
                            children: [
                                pasalAndBabParagraph,
                                ...contentParagraphs,
                            ],
                            alignment: AlignmentType.JUSTIFIED,
                        }),
                        // Empty cells for 'KETERANGAN'
                        new TableCell({ children: [new Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER, font: "Orbi", fontSize: 12 }), // Add font and fontSize properties
                        new TableCell({ children: [new Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER, font: "Orbi", fontSize: 12 }), // Add font and fontSize properties
                    ],
                }),
            ];
        }).flat(), // Flatten the array of rows
    ];

    // Create table with rows 
    const table = new Table({
        rows: tableRows,
    });
    return table;
}


// Function to create Word document
function createWordDocument(data) {
    const jsonData = data[0]; // Assuming you are taking the first object from the JSON array

    // Get the context from JSON data
    const context = jsonData.context;

    // Split the context into words
    const words = context.split(' ');

    // Take the first three words
    const firstThreeWords = words.slice(0, 5).join(' ');

    // Combine first three words into a paragraph
    const paragraph = new Paragraph({
        children: [new docx.TextRun({ text: firstThreeWords, bold: true, fontSize: 14 })],
        alignment: AlignmentType.CENTER,
    });

    // If there are more than three words, add a break and include the remaining words
    if (words.length > 5) {
        const remainingWords = words.slice(5).join(' ');
        paragraph.addChildElement(new docx.TextRun({ text: '\n' + remainingWords, bold: true, fontSize: 14, break: 1}));
    }

    // Create table using createTable function with JSON data
    const content = createTable(data);

    // Combine paragraph and content into children array
    const children = [paragraph, content];

    // Create numbering configuration
    const numberingConfig = data.map((item, index) => ({
        reference: `my-numbering-${index + 1}`,
        levels: [
            {
                level: 0,
                format: docx.LevelFormat.DECIMAL,
                text: "(%1)",
            },
        ],
    })).concat(data.map((item, index) => ({
        reference: `my-other-numbering-${index + 1}`,
        levels: [
            {
                level: 0,
                format: docx.LevelFormat.LOWER_LETTER,
                text: "%1.",
            },
        ],
    })));

    // Create Word document
    const document = new Document({
        numbering: {
            config: numberingConfig,
        },
        sections: [
            {
                properties: {
                    page: {
                        size: {
                            orientation: docx.PageOrientation.LANDSCAPE,
                        },
                    },
                },
                headers: {
                    default: new Header({
                        children: [new Paragraph("Header placement")], // Add your header here
                    }),
                },
                footers: {
                    default: new Footer({
                        children: [new Paragraph("Footer placement")],
                    }),
                },
                children: children,
            },
        ],
    });
    

    return document;
}




// Path to JSON file
const jsonFilePath = 'D:/magang/json_ke_word/new_file/permenkeu-no-26-tahun-2023-updated.json';

// Read JSON file
const jsonData = readJsonFile(jsonFilePath);

// Create Word document based on JSON data
const document = createWordDocument(jsonData);

// Save document to file
Packer.toBuffer(document).then((buffer) => {
    fs.writeFileSync('D:/magang/json_ke_word/new_file/ouaaatput.docx', buffer);
    console.log('Dokumen Word berhasil dibuat: ouaaatput.docx');
});

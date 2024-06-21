const fs = require("fs");
const https = require("https");
const dotenv = require("dotenv");
const docx = require("docx");
const {
  Document,
  VerticalAlign,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  Header,
  Footer,
  WidthType,
  AlignmentType,
} = docx;

// Load konfigurasi dari file .env
dotenv.config();

// Mendapatkan variabel konfigurasi dari file .env
const OPENSEARCH_ADDRESS = process.env.OPENSEARCH_ADDRESS;
const OPENSEARCH_PORT = process.env.OPENSEARCH_PORT;
const OPENSEARCH_AUTH = process.env.OPENSEARCH_AUTH;

// Nama indeks OpenSearch dan slug value
const INDEX_NAME = "produk_hukum";
const slugValue = process.argv[2] || "14-pmk-02-2020";

// Konfigurasi koneksi ke cluster OpenSearch
const options = {
  hostname: OPENSEARCH_ADDRESS,
  port: OPENSEARCH_PORT,
  path: `/${INDEX_NAME}/_search`,
  method: "POST",
  headers: {
    "Content-Type": "application/json",
    Authorization: "Basic " + Buffer.from(OPENSEARCH_AUTH).toString("base64"),
  },
  rejectUnauthorized: false, // Setel ini ke false
};

// Query pencarian untuk mencari dokumen berdasarkan Slug
const searchQuery = {
  query: {
    term: {
      Slug: {
        value: slugValue,
      },
    },
  },
};

// Body permintaan dengan query pencarian
const postData = JSON.stringify(searchQuery);

async function fetchOpensearchData() {
  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = "";

      res.on("data", (chunk) => {
        data += chunk;
      });

      res.on("end", () => {
        try {
          const response = JSON.parse(data);
          const modifiedResponse = response.hits.hits
            .map((hit) => {
              return hit._source.Blocks.map((block) => ({
                bab: block.bab === "none" ? null : block.bab,
                judulbab: "Ketentuan Umum",
                bagian: block.bagian,
                paragraf: block.paragraf,
                pasal: block.pasal,
                ref: null,
                type: block.type,
                content: block.content,
                additional_context: [],
                context: hit._source.Judul,
              }));
            })
            .flat();

          fs.writeFileSync(
            "public/uper_response.json",
            JSON.stringify(modifiedResponse, null, 2)
          );
          resolve();
        } catch (error) {
          reject(error);
        }
      });
    });

    req.on("error", (error) => {
      reject(error);
    });

    req.write(postData);
    req.end();
  });
}

async function parseAndModifyData() {
  return new Promise((resolve, reject) => {
    // Read the JSON file
    fs.readFile("public/uper_response.json", "utf8", (err, data) => {
      if (err) {
        console.error("Error reading file:", err);
        reject(err);
        return;
      }
      try {
        // Parse JSON
        const jsonData = JSON.parse(data);

        // Array to store modified data
        let modifiedData = [];

        // Variable to store the previous chapter value
        let previousBab = null;

        // Process modification of data
        jsonData.forEach((item) => {
          if (item.type === "CONTENT_PASAL") {
            let combinedContent = item.content.replace(/\.\s+/g, ".\n");

            // Set chapter value according to the required logic
            let modifiedBab = item.bab;
            if (modifiedBab === previousBab) {
              modifiedBab = null; // Set to null if the same as the previous chapter value
            } else {
              previousBab = item.bab; // Update the previous chapter value
            }

            modifiedData.push({
              bab: modifiedBab,
              judulbab: item.judulbab,
              bagian: item.bagian,
              paragraf: item.paragraf,
              pasal: item.pasal,
              ref: null,
              type: "CONTENT_PASAL",
              content: combinedContent.trim(),
              additional_context: [],
              context: item.context,
            });
          }
        });

        // Write back to a JSON file
        fs.writeFile(
          "material.json",
          JSON.stringify(modifiedData, null, 2),
          "utf8",
          (err) => {
            if (err) {
              console.error("Error writing file:", err);
              reject(err);
              return;
            }
            console.log("File material.json has been successfully saved.");
            resolve();
          }
        );
      } catch (error) {
        console.error("Error parsing JSON:", error);
        reject(error);
      }
    });
  });
}

async function generateWordDocument() {
  function readJsonFile(filePath) {
    const jsonData = fs.readFileSync(filePath, "utf8");
    return JSON.parse(jsonData);
  }

  function createTable(data) {
    let numberingInstance = 0;
    let alphabeticalNumberingInstance = -1;

    const tableRows = [
      new TableRow({
        children: [
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new docx.TextRun({
                    text: "NO.",
                    bold: true,
                    font: "Bookman Old Style",
                    size: 22,
                  }),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
            width: { size: 5, type: WidthType.PERCENTAGE },
            shading: { fill: "d9e2f3" },
          }),
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new docx.TextRun({
                    text: "SAAT INI",
                    bold: true,
                    font: "Bookman Old Style",
                    size: 22,
                  }),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
            width: { size: 35, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: "d9e2f3" },
          }),
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new docx.TextRun({
                    text: "PERUBAHAN",
                    color: "#4472c4",
                    bold: true,
                    font: "Bookman Old Style",
                    size: 22,
                  }),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
            width: { size: 35, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: "d9e2f3" },
          }),
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new docx.TextRun({
                    text: "KETERANGAN",
                    bold: true,
                    font: "Bookman Old Style",
                    size: 22,
                  }),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
            width: { size: 25, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            shading: { fill: "d9e2f3" },
          }),
        ],
        tableHeader: true,
      }),

      // Add dynamic rows here...
      ...data.map((item, index) => {
        const cleanedContent = item.content.replace(/;/g, ".");
        const numberingReference = `my-numbering-${index + 1}`;
        const otherNumberingReference = `my-other-numbering-${index + 1}`;

        const paragraphs = item.content
          .split(/(?=[0-9a-zA-Z]\. |\([0-9]+\)\s)/gm)
          .map((contentPart) => contentPart.trim());
        const contentParagraphs = [];

        paragraphs.forEach((contentPart) => {
          let numbering = undefined;

          if (contentPart.match(/^\(\d+\)\s/)) {
            numberingInstance++;
            alphabeticalNumberingInstance = 0;
            numbering = {
              reference: numberingReference,
              level: 0,
              format: docx.LevelFormat.BULLET,
              text: `${numberingInstance}`,
            };
            contentPart = contentPart.replace(/^\(\d+\)\s/, "");
          } else if (contentPart.match(/^[a-z]\.\s/)) {
            if (alphabeticalNumberingInstance === -1) {
              alphabeticalNumberingInstance = 0;
            }
            const charCode = "a".charCodeAt(0) + alphabeticalNumberingInstance;
            alphabeticalNumberingInstance++;
            numbering = {
              reference: otherNumberingReference,
              level: 0,
              format: docx.LevelFormat.LOWER_LETTER,
              text: `${String.fromCharCode(charCode)}.`,
            };
            contentPart = contentPart.replace(/^[a-z]\.\s/, "");
          }

          const lastChar = contentPart.charAt(contentPart.length - 1);
          if (!lastChar.match(/[a-zA-Z0-9]/)) {
            contentPart = contentPart.slice(0, -1);
          }

          contentParagraphs.push(
            new Paragraph({
              children: [
                new docx.TextRun({
                  text: contentPart,
                  font: "Bookman Old Style",
                  size: 22,
                }),
              ],
              numbering,
              alignment: AlignmentType.JUSTIFIED,
              spacing: { after: 100, before: 100 },
              indent: { left: 400, right: 100 },
            })
          );
        });

        const pasalParagraph = new Paragraph({
          children: [
            new docx.TextRun({
              text: item.pasal
                ? `${item.pasal
                    .replace(/-/g, " ")
                    .replace(/\b\w/g, (c) => c.toUpperCase())}`
                : "",
              bold: true,
              font: "Bookman Old Style",
              size: 22,
            }),
          ],
          alignment: AlignmentType.CENTER,
        });

        const babParagraph = item.bab
          ? new Paragraph({
              children: [
                new docx.TextRun({
                  text: item.bab,
                  bold: true,
                  font: "Bookman Old Style",
                  size: 22,
                }),
                new docx.TextRun({
                  text: item.judulbab,
                  break: 1,
                  bold: true,
                  font: "Bookman Old Style",
                  size: 22,
                }),
              ],
              alignment: AlignmentType.CENTER,
            })
          : null;

        return [
          babParagraph
            ? new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph("")],
                    verticalAlign: docx.VerticalAlign.CENTER,
                  }),
                  new TableCell({
                    children: [babParagraph],
                    verticalAlign: docx.VerticalAlign.CENTER,
                    shading: { fill: "F8E8EE" },
                  }),
                  new TableCell({
                    children: [new Paragraph("")],
                    verticalAlign: docx.VerticalAlign.CENTER,
                  }),
                  new TableCell({
                    children: [new Paragraph("")],
                    verticalAlign: docx.VerticalAlign.CENTER,
                  }),
                ],
              })
            : null,
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    text: String(index + 1),
                    alignment: AlignmentType.CENTER,
                    font: "Bookman Old Style",
                    size: 22,
                  }),
                ],
              }),
              new TableCell({
                children: [pasalParagraph, ...contentParagraphs],
                alignment: AlignmentType.JUSTIFIED,
              }),
              new TableCell({
                children: [new Paragraph("")],
                verticalAlign: docx.VerticalAlign.CENTER,
              }),
              new TableCell({
                children: [new Paragraph("")],
                verticalAlign: docx.VerticalAlign.CENTER,
              }),
            ],
          }),
        ].filter(Boolean);
      }),
    ].flat();

    const table = new Table({ rows: tableRows });
    return table;
  }

  function createWordDocument(data) {
    const jsonData = data[0];
    const context = jsonData.context;
    const words = context.split(" ");
    const firstThreeWords = words.slice(0, 5).join(" ");

    const paragraph = new Paragraph({
      children: [
        new docx.TextRun({
          text: firstThreeWords,
          bold: true,
          font: "Bookman Old Style",
          size: 22,
        }),
      ],
      alignment: AlignmentType.CENTER,
    });

    if (words.length > 5) {
      const remainingWords = words.slice(5).join(" ");
      paragraph.addChildElement(
        new docx.TextRun({
          text: "\n" + remainingWords,
          bold: true,
          size: 22,
          font: "Bookman Old Style",
          break: 1,
        })
      );
    }

    const content = createTable(data);
    const children = [paragraph, content];

    const numberingConfig = data
      .map((item, index) => ({
        reference: `my-numbering-${index + 1}`,
        levels: [{ level: 0, format: docx.LevelFormat.DECIMAL, text: "(%1)" }],
      }))
      .concat(
        data.map((item, index) => ({
          reference: `my-other-numbering-${index + 1}`,
          levels: [
            { level: 0, format: docx.LevelFormat.LOWER_LETTER, text: "%1." },
          ],
        }))
      );

    const document = new Document({
      numbering: { config: numberingConfig },
      sections: [
        {
          properties: {
            page: {
              size: { orientation: docx.PageOrientation.LANDSCAPE },
              margin: { top: 720, right: 720, bottom: 720, left: 720 },
            },
          },
          headers: {
            default: new Header({
              children: [new Paragraph("Header placement")],
              properties: {
                footer: { marginTop: 50, marginBottom: 50 },
              },
            }),
          },
          footers: {
            default: new Footer({
              children: [
                new Paragraph(
                  "Generated with https://github.com/gagahputrabangsa/Legal_Viewer_Msib.git"
                ),
              ],
              properties: {
                footer: { marginTop: 50, marginBottom: 50 },
              },
            }),
          },
          children: children,
        },
      ],
    });

    return document;
  }

  const jsonFilePath = "material.json";
  const jsonData = readJsonFile(jsonFilePath);
  const document = createWordDocument(jsonData);

  Packer.toBuffer(document).then((buffer) => {
    const outputPath = "public/matriks_peran.docx";

    fs.writeFileSync(outputPath, buffer);
    console.log(`Dokumen Word berhasil dan disimpan di ${outputPath}`);
    console.log("YEAY");
  });
}

// Main function
async function main() {
  try {
    await fetchOpensearchData();
    await parseAndModifyData();
    await generateWordDocument();
  } catch (error) {
    console.error("Error during processing:", error);
  }
}

// Run main function
main();

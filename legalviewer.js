"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", { value: true });
var fs = require("fs");
var https = require("https");
var dotenv = require("dotenv");
var docx = require("docx");
var docx_1 = require("docx");
// Load konfigurasi dari file .env
dotenv.config();
// Mendapatkan variabel konfigurasi dari file .env
var OPENSEARCH_ADDRESS = process.env.OPENSEARCH_ADDRESS;
var OPENSEARCH_PORT = process.env.OPENSEARCH_PORT;
var OPENSEARCH_USERNAME = process.env.OPENSEARCH_USERNAME;
var OPENSEARCH_PASSWORD = process.env.OPENSEARCH_PASSWORD;
// Nama indeks OpenSearch
var INDEX_NAME = 'law_analyzer_msib';
// Mengambil ID dari argumen command line
var args = process.argv.slice(2);
var idIndex = args.indexOf('-id');
if (idIndex === -1 || idIndex === args.length - 1) {
    console.error("Error: Please provide an ID using the '-id' argument.");
    process.exit(1);
}
var idValue = args[idIndex + 1];
// Menghapus tanda petik dari nilai ID
idValue = idValue.replace(/^'|'$/g, '');
// Konfigurasi koneksi ke cluster OpenSearch
var options = {
    hostname: OPENSEARCH_ADDRESS,
    port: OPENSEARCH_PORT,
    path: "/".concat(INDEX_NAME, "/_search"),
    method: 'POST',
    headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Basic ' + Buffer.from("".concat(OPENSEARCH_USERNAME, ":").concat(OPENSEARCH_PASSWORD)).toString('base64')
    },
    rejectUnauthorized: false // Setel ini ke false
};
// Query pencarian untuk mencari dokumen berdasarkan ID dengan mengecualikan beberapa field
var searchQuery = {
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
var postData = JSON.stringify(searchQuery);
function fetchOpensearchData() {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve, reject) {
                    // Buat permintaan HTTPS
                    var req = https.request(options, function (res) {
                        console.log("Status Code: ".concat(res.statusCode));
                        var data = '';
                        res.on('data', function (chunk) {
                            data += chunk;
                        });
                        res.on('end', function () {
                            try {
                                var response = JSON.parse(data);
                                // Simpan respons ke dalam file JSON
                                fs.writeFileSync('D:/magang/json_ke_word/response.json', JSON.stringify(response, null, 2));
                                console.log('Response saved to response.json');
                                resolve();
                            }
                            catch (error) {
                                console.error('Error parsing JSON response:', error);
                                reject(error);
                            }
                        });
                    });
                    req.on('error', function (error) {
                        console.error(error);
                        reject(error);
                    });
                    // Kirim data
                    req.write(postData);
                    req.end();
                })];
        });
    });
}
function parseAndModifyData() {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve, reject) {
                    // Read the JSON file
                    fs.readFile('D:/magang/json_ke_word/response.json', 'utf8', function (err, data) {
                        if (err) {
                            console.error('Error reading file:', err);
                            reject(err);
                            return;
                        }
                        try {
                            // Parse JSON
                            var jsonData = JSON.parse(data);
                            // Array to store modified data
                            var modifiedData_1 = [];
                            // Variable to store the previous chapter value
                            var previousBab_1 = null;
                            // Process modification of data
                            jsonData.hits.hits.forEach(function (item) {
                                item._source.Blocks.forEach(function (block) {
                                    if (block.ContentText && Array.isArray(block.ContentText)) {
                                        var combinedContent_1 = "";
                                        block.ContentText.forEach(function (content) {
                                            // Check if content is an object
                                            if (typeof content === 'object' && content !== null) {
                                                // Modify text content
                                                var modifiedValue = content.Value.replace(/dimaksud pada ayat \((\d+)\)/g, 'dimaksud pada ayat $1');
                                                modifiedValue = modifiedValue.replace(/dimaksud dalam Pasal (\d+) ayat \((\d+)\)/g, 'dimaksud dalam Pasal $1 ayat $2');
                                                var modifiedRef = content.Ref ? content.Ref.replace(/Ayat \((\d+)\)/g, '($1)') : null;
                                                // Remove patterns like -number-
                                                modifiedValue = modifiedValue.replace(/-\d+-/g, '');
                                                // Add new patterns
                                                modifiedValue = modifiedValue.replace(/(\d+)\. /g, '($1) ');
                                                modifiedValue = modifiedValue.replace(/\. \((\d+)\)/g, '.\n($1)');
                                                combinedContent_1 += "".concat(modifiedRef ? "".concat(modifiedRef, " ") : '').concat(modifiedValue, "\n");
                                            }
                                        });
                                        // Set chapter value according to the required logic
                                        var modifiedBab = block.Bab;
                                        if (modifiedBab === previousBab_1) {
                                            modifiedBab = null; // Set to null if the same as the previous chapter value
                                        }
                                        else {
                                            previousBab_1 = block.Bab; // Update the previous chapter value
                                        }
                                        modifiedData_1.push({
                                            bab: modifiedBab,
                                            judulbab: block.BabContext,
                                            bagian: block.Bagian,
                                            paragraf: block.Paragraf,
                                            pasal: block.Pasal ? "pasal-".concat(block.Pasal.split(' ')[1]) : null,
                                            ref: null,
                                            type: "CONTENT_PASAL",
                                            content: combinedContent_1.trim(),
                                            additional_context: [],
                                            context: item._source.Judul
                                        });
                                    }
                                });
                            });
                            // Write back to a JSON file
                            fs.writeFile('D:/magang/json_ke_word/new_file/typeS/material/material.json', JSON.stringify(modifiedData_1, null, 2), 'utf8', function (err) {
                                if (err) {
                                    console.error('Error writing file:', err);
                                    reject(err);
                                    return;
                                }
                                console.log('File material.json has been successfully saved.');
                                resolve();
                            });
                        }
                        catch (error) {
                            console.error('Error parsing JSON:', error);
                            reject(error);
                        }
                    });
                })];
        });
    });
}
function generateWordDocument() {
    return __awaiter(this, void 0, void 0, function () {
        function readJsonFile(filePath) {
            var jsonData = fs.readFileSync(filePath, 'utf8');
            return JSON.parse(jsonData);
        }
        function createTable(data) {
            var numberingInstance = 0;
            var alphabeticalNumberingInstance = -1;
            var tableRows = __spreadArray([
                new docx_1.TableRow({
                    children: [
                        new docx_1.TableCell({
                            children: [
                                new docx_1.Paragraph({
                                    children: [
                                        new docx.TextRun({ text: 'NO.', bold: true, font: "Bookman Old Style", size: 22 })
                                    ],
                                    alignment: docx_1.AlignmentType.CENTER
                                })
                            ],
                            width: { size: 5, type: docx_1.WidthType.PERCENTAGE },
                            shading: { fill: 'd9e2f3' }
                        }),
                        new docx_1.TableCell({
                            children: [
                                new docx_1.Paragraph({
                                    children: [
                                        new docx.TextRun({ text: 'SAAT INI', bold: true, font: "Bookman Old Style", size: 22 })
                                    ],
                                    alignment: docx_1.AlignmentType.CENTER
                                })
                            ],
                            width: { size: 35, type: docx_1.WidthType.PERCENTAGE },
                            verticalAlign: docx_1.VerticalAlign.CENTER,
                            shading: { fill: 'd9e2f3' }
                        }),
                        new docx_1.TableCell({
                            children: [
                                new docx_1.Paragraph({
                                    children: [
                                        new docx.TextRun({ text: 'PERUBAHAN', color: '#4472c4', bold: true, font: "Bookman Old Style", size: 22 })
                                    ],
                                    alignment: docx_1.AlignmentType.CENTER
                                })
                            ],
                            width: { size: 35, type: docx_1.WidthType.PERCENTAGE },
                            verticalAlign: docx_1.VerticalAlign.CENTER,
                            shading: { fill: 'd9e2f3' }
                        }),
                        new docx_1.TableCell({
                            children: [
                                new docx_1.Paragraph({
                                    children: [
                                        new docx.TextRun({ text: 'KETERANGAN', bold: true, font: "Bookman Old Style", size: 22 })
                                    ],
                                    alignment: docx_1.AlignmentType.CENTER
                                })
                            ],
                            width: { size: 25, type: docx_1.WidthType.PERCENTAGE },
                            verticalAlign: docx_1.VerticalAlign.CENTER,
                            shading: { fill: 'd9e2f3' }
                        })
                    ],
                    tableHeader: true
                })
            ], data.map(function (item, index) {
                var cleanedContent = item.content.replace(/;/g, '.');
                var numberingReference = "my-numbering-".concat(index + 1);
                var otherNumberingReference = "my-other-numbering-".concat(index + 1);
                var numberingInstance = 0;
                var alphabeticalNumberingInstance = -1;
                var paragraphs = item.content.split(/(?=[0-9a-zA-Z]\. |\([0-9]+\)\s)/gm).map(function (contentPart) { return contentPart.trim(); });
                var contentParagraphs = [];
                paragraphs.forEach(function (contentPart) {
                    var numbering = undefined;
                    if (contentPart.match(/^\(\d+\)\s/)) {
                        numberingInstance++;
                        alphabeticalNumberingInstance = 0;
                        numbering = {
                            reference: numberingReference,
                            level: 0,
                            format: docx.LevelFormat.BULLET,
                            text: "".concat(numberingInstance),
                        };
                        contentPart = contentPart.replace(/^\(\d+\)\s/, '');
                    }
                    else if (contentPart.match(/^[a-z]\.\s/)) {
                        if (alphabeticalNumberingInstance === -1) {
                            alphabeticalNumberingInstance = 0;
                        }
                        var charCode = 'a'.charCodeAt(0) + alphabeticalNumberingInstance;
                        alphabeticalNumberingInstance++;
                        numbering = {
                            reference: otherNumberingReference,
                            level: 0,
                            format: docx.LevelFormat.LOWER_LETTER,
                            text: "".concat(String.fromCharCode(charCode), "."),
                        };
                        contentPart = contentPart.replace(/^[a-z]\.\s/, '');
                    }
                    var lastChar = contentPart.charAt(contentPart.length - 1);
                    if (!lastChar.match(/[a-zA-Z0-9]/)) {
                        contentPart = contentPart.slice(0, -1);
                    }
                    contentParagraphs.push(new docx_1.Paragraph({
                        children: [new docx.TextRun({ text: contentPart, font: "Bookman Old Style", size: 22 })],
                        numbering: numbering,
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { after: 100, before: 100 },
                        indent: { left: 400, right: 100 },
                    }));
                });
                var pasalParagraph = new docx_1.Paragraph({
                    children: [new docx.TextRun({ text: item.pasal ? "".concat(item.pasal.replace(/-/g, ' ').replace(/\b\w/g, function (c) { return c.toUpperCase(); })) : '', bold: true, font: "Bookman Old Style", size: 22 })],
                    alignment: docx_1.AlignmentType.CENTER,
                });
                var babParagraph = item.bab ? new docx_1.Paragraph({
                    children: [
                        new docx.TextRun({ text: item.bab, bold: true, font: "Bookman Old Style", size: 22 }),
                        new docx.TextRun({ text: item.judulbab, break: 1, bold: true, font: "Bookman Old Style", size: 22 }),
                    ],
                    alignment: docx_1.AlignmentType.CENTER,
                }) : null;
                return [
                    babParagraph ? new docx_1.TableRow({
                        children: [
                            new docx_1.TableCell({ children: [new docx_1.Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER }),
                            new docx_1.TableCell({ children: [babParagraph], verticalAlign: docx.VerticalAlign.CENTER, shading: { fill: 'F8E8EE' } }),
                            new docx_1.TableCell({ children: [new docx_1.Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER }),
                            new docx_1.TableCell({ children: [new docx_1.Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER }),
                        ],
                    }) : null,
                    new docx_1.TableRow({
                        children: [
                            new docx_1.TableCell({ children: [new docx_1.Paragraph({ text: String(index + 1), alignment: docx_1.AlignmentType.CENTER, font: "Bookman Old Style", size: 22 })] }),
                            new docx_1.TableCell({ children: __spreadArray([pasalParagraph], contentParagraphs, true), alignment: docx_1.AlignmentType.JUSTIFIED }),
                            new docx_1.TableCell({ children: [new docx_1.Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER }),
                            new docx_1.TableCell({ children: [new docx_1.Paragraph('')], verticalAlign: docx.VerticalAlign.CENTER }),
                        ],
                    }),
                ].filter(Boolean);
            }), true).flat();
            var table = new docx_1.Table({ rows: tableRows });
            return table;
        }
        function createWordDocument(data) {
            var jsonData = data[0];
            var context = jsonData.context;
            var words = context.split(' ');
            var firstThreeWords = words.slice(0, 5).join(' ');
            var paragraph = new docx_1.Paragraph({
                children: [new docx.TextRun({ text: firstThreeWords, bold: true, font: "Bookman Old Style", size: 22 })],
                alignment: docx_1.AlignmentType.CENTER,
            });
            if (words.length > 5) {
                var remainingWords = words.slice(5).join(' ');
                paragraph.addChildElement(new docx.TextRun({ text: '\n' + remainingWords, bold: true, size: 22, font: "Bookman Old Style", break: 1 }));
            }
            var content = createTable(data);
            var children = [paragraph, content];
            var numberingConfig = data.map(function (item, index) { return ({
                reference: "my-numbering-".concat(index + 1),
                levels: [{ level: 0, format: docx.LevelFormat.DECIMAL, text: "(%1)" }],
            }); }).concat(data.map(function (item, index) { return ({
                reference: "my-other-numbering-".concat(index + 1),
                levels: [{ level: 0, format: docx.LevelFormat.LOWER_LETTER, text: "%1." }],
            }); }));
            var document = new docx_1.Document({
                numbering: { config: numberingConfig },
                sections: [{
                        properties: {
                            page: {
                                size: { orientation: docx.PageOrientation.LANDSCAPE },
                                margin: { top: 720, right: 720, bottom: 720, left: 720 },
                            },
                        },
                        headers: { default: new docx_1.Header({ children: [new docx_1.Paragraph("Header placement")], properties: { footer: { marginTop: 50, marginBottom: 50 } } }) },
                        footers: { default: new docx_1.Footer({ children: [new docx_1.Paragraph("Generated with https://github.com/gagahputrabangsa/Legal_Viewer_Msib.git")], properties: { footer: { marginTop: 50, marginBottom: 50 } } }) },
                        children: children,
                    }],
            });
            return document;
        }
        var jsonFilePath, jsonData, document;
        return __generator(this, function (_a) {
            jsonFilePath = 'D:/magang/json_ke_word/new_file/typeS/material/material.json';
            jsonData = readJsonFile(jsonFilePath);
            document = createWordDocument(jsonData);
            docx_1.Packer.toBuffer(document).then(function (buffer) {
                var args = process.argv.slice(2);
                var outputArgIndex = args.indexOf('-out');
                var outputPath = 'output.docx';
                if (outputArgIndex !== -1 && args[outputArgIndex + 1]) {
                    outputPath = args[outputArgIndex + 1];
                }
                else {
                    console.error('Error: Output path not specified. Use -out <outputPath> to specify the output file.');
                    process.exit(1);
                }
                fs.writeFileSync(outputPath, buffer);
                console.log("Dokumen Word berhasil dan disimpan di ".concat(outputPath));
                console.log("\n'YEAY'    \n");
            });
            return [2 /*return*/];
        });
    });
}
// Main function
function main() {
    return __awaiter(this, void 0, void 0, function () {
        var error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 4, , 5]);
                    return [4 /*yield*/, fetchOpensearchData()];
                case 1:
                    _a.sent();
                    return [4 /*yield*/, parseAndModifyData()];
                case 2:
                    _a.sent();
                    return [4 /*yield*/, generateWordDocument()];
                case 3:
                    _a.sent();
                    return [3 /*break*/, 5];
                case 4:
                    error_1 = _a.sent();
                    console.error('Error during processing:', error_1);
                    return [3 /*break*/, 5];
                case 5: return [2 /*return*/];
            }
        });
    });
}
// Run main function
main();

const { createWorker } = require('tesseract.js');
const pdfParse = require('pdf-parse');
const xlsx = require('xlsx');
const { Document } = require('docx');
const mammoth = require('mammoth');

class ConvertFileToJsonEnhanced {
    description = {
        displayName: 'Convert File to JSON (Enhanced)',
        name: 'convertFileToJsonEnhanced',
        icon: 'fa:file-code',
        group: ['transform'],
        version: 1,
        description: 'Convert files (PDF, Word, Excel, Images, etc.) to JSON with enhanced extraction',
        defaults: {
            name: 'Convert File to JSON (Enhanced)',
        },
        inputs: ['main'],
        outputs: ['main'],
        properties: [
            {
                displayName: 'Binary Property',
                name: 'binaryPropertyName',
                type: 'string',
                default: 'data',
                required: true,
                description: 'Name of the binary property containing the file',
            },
            {
                displayName: 'Include File Name',
                name: 'includeFileName',
                type: 'boolean',
                default: true,
                description: 'Include the original file name in the output',
            },
            {
                displayName: 'Include Sheet Name',
                name: 'includeSheetName',
                type: 'boolean',
                default: true,
                displayOptions: {
                    show: {
                        fileType: ['xlsx', 'xls'],
                    },
                },
                description: 'Include sheet names for Excel files',
            },
            {
                displayName: 'Include Original Row Numbers',
                name: 'includeRowNumbers',
                type: 'boolean',
                default: false,
                description: 'Include original row numbers in the output',
            },
            {
                displayName: 'Output Sheets as Separate Items',
                name: 'separateSheets',
                type: 'boolean',
                default: false,
                displayOptions: {
                    show: {
                        fileType: ['xlsx', 'xls'],
                    },
                },
                description: 'Create separate output items for each sheet',
            },
        ],
    };

    async execute() {
        const items = this.getInputData();
        const returnData = [];

        for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
            try {
                const binaryPropertyName = this.getNodeParameter('binaryPropertyName', itemIndex);
                const includeFileName = this.getNodeParameter('includeFileName', itemIndex);
                const includeRowNumbers = this.getNodeParameter('includeRowNumbers', itemIndex);

                const binaryData = items[itemIndex].binary[binaryPropertyName];
                if (!binaryData) {
                    throw new Error(`No binary data found in property "${binaryPropertyName}"`);
                }

                const buffer = Buffer.from(binaryData.data, 'base64');
                const mimeType = binaryData.mimeType;
                const fileName = binaryData.fileName || 'unknown';

                let result;

                // Detect file type and process
                if (mimeType === 'application/pdf' || fileName.endsWith('.pdf')) {
                    result = await this.processPDF(buffer);
                } else if (mimeType.includes('word') || fileName.endsWith('.docx')) {
                    result = await this.processWord(buffer);
                } else if (mimeType.includes('spreadsheet') || fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
                    const includeSheetName = this.getNodeParameter('includeSheetName', itemIndex);
                    const separateSheets = this.getNodeParameter('separateSheets', itemIndex);
                    result = await this.processExcel(buffer, includeSheetName, includeRowNumbers, separateSheets);
                } else if (mimeType === 'text/csv' || fileName.endsWith('.csv')) {
                    result = await this.processCSV(buffer, includeRowNumbers);
                } else if (mimeType.startsWith('image/')) {
                    result = await this.processImage(buffer);
                } else if (mimeType === 'application/json' || fileName.endsWith('.json')) {
                    result = JSON.parse(buffer.toString());
                } else if (mimeType.startsWith('text/')) {
                    result = { text: buffer.toString() };
                } else {
                    throw new Error(`Unsupported file type: ${mimeType}`);
                }

                // Add metadata if requested
                if (includeFileName && result) {
                    if (Array.isArray(result)) {
                        result.forEach(item => {
                            if (typeof item === 'object') item.fileName = fileName;
                        });
                    } else if (typeof result === 'object') {
                        result.fileName = fileName;
                    }
                }

                // Handle array results (e.g., multiple sheets)
                if (Array.isArray(result)) {
                    result.forEach(item => {
                        returnData.push({ json: item });
                    });
                } else {
                    returnData.push({ json: result });
                }
            } catch (error) {
                if (this.continueOnFail()) {
                    returnData.push({
                        json: {
                            error: error.message,
                        },
                    });
                    continue;
                }
                throw error;
            }
        }

        return [returnData];
    }

    async processPDF(buffer) {
        const data = await pdfParse(buffer);
        return {
            text: data.text,
            pages: data.numpages,
            info: data.info,
        };
    }

    async processWord(buffer) {
        const result = await mammoth.extractRawText({ buffer });
        return {
            text: result.value,
        };
    }

    async processExcel(buffer, includeSheetName, includeRowNumbers, separateSheets) {
        const workbook = xlsx.read(buffer);
        const results = [];

        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            let data = xlsx.utils.sheet_to_json(worksheet);

            if (includeRowNumbers) {
                data = data.map((row, idx) => ({
                    _rowNumber: idx + 2, // Excel rows start at 1, headers at 1
                    ...row,
                }));
            }

            if (separateSheets) {
                const sheetData = includeSheetName ? { sheetName, data } : { data };
                results.push(sheetData);
            } else {
                if (includeSheetName) {
                    results.push({ sheetName, data });
                } else {
                    results.push(...data);
                }
            }
        });

        return separateSheets ? results : { sheets: results };
    }

    async processCSV(buffer, includeRowNumbers) {
        const text = buffer.toString();
        const lines = text.split('\n');
        const headers = lines[0].split(',').map(h => h.trim());
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            if (lines[i].trim()) {
                const values = lines[i].split(',');
                const row = {};
                if (includeRowNumbers) row._rowNumber = i + 1;
                headers.forEach((header, idx) => {
                    row[header] = values[idx] ? values[idx].trim() : '';
                });
                data.push(row);
            }
        }

        return { data };
    }

    async processImage(buffer) {
        const worker = await createWorker();
        await worker.loadLanguage('eng');
        await worker.initialize('eng');
        const { data: { text } } = await worker.recognize(buffer);
        await worker.terminate();

        return {
            text,
            extractedBy: 'OCR',
        };
    }
}

module.exports.nodeClass = ConvertFileToJsonEnhanced;

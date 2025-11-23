# Convert File to JSON (Enhanced)

A powerful n8n community node that converts various file formats to JSON with enhanced extraction capabilities.

## Features

This node supports automatic detection and conversion of:

- **PDF** files (extracts text, pages, metadata)
- **Word documents** (.docx)
- **Excel spreadsheets** (.xlsx, .xls) with multi-sheet support
- **CSV** files
- **Images** (PNG, JPG) with OCR using Tesseract.js
- **JSON** files
- **Text** files

## Installation

### From n8n UI

1. Go to **Settings** → **Community Nodes**
2. Click **Install**
3. Enter: `https://github.com/naifue/n8n-nodes-convert-file-json-enhanced`
4. Accept the security warning
5. Wait for n8n to restart

### From npm

```bash
npm install n8n-nodes-convert-file-json-enhanced
```

## Configuration Options

- **Binary Property**: Name of the binary property containing the file (default: `data`)
- **Include File Name**: Add the original file name to output (default: `true`)
- **Include Sheet Name**: Include sheet names for Excel files (default: `true`)
- **Include Original Row Numbers**: Add row numbers to Excel/CSV data (default: `false`)
- **Output Sheets as Separate Items**: Create separate items for each Excel sheet (default: `false`)

## Use Cases

- Extract text from PDF documents for RAG systems
- Convert Excel data for Pinecone vector storage
- OCR images from Google Drive
- Process Word documents in workflows
- Parse CSV files with metadata

## Example Workflow

```
Google Drive Trigger → Convert File to JSON (Enhanced) → Pinecone Vector Store
```

## Supported File Types

| Format | Extension | Features |
|--------|-----------|----------|
| PDF | .pdf | Text extraction, page count, metadata |
| Word | .docx | Full text extraction |
| Excel | .xlsx, .xls | Multi-sheet, row numbers, structured data |
| CSV | .csv | Automatic header detection, row numbers |
| Images | .png, .jpg, .jpeg | OCR text extraction |
| JSON | .json | Direct parsing |
| Text | .txt | Plain text extraction |

## Output Format

The node outputs JSON objects compatible with downstream n8n nodes like Pinecone:

```json
{
  "text": "Extracted content...",
  "fileName": "document.pdf",
  "pages": 5
}
```

## Limitations

- OCR accuracy depends on image quality
- Large files may take time to process
- Excel files with complex formulas may not render calculated values

## License

MIT

## Author

naifue

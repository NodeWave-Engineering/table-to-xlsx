# @nodewave/table-to-xlsx

A lightweight Node.js package to convert HTML tables to Excel files with styling, handles merged cells, and customizable titles.

## This Package Depends on : 
- cheerio (https://www.npmjs.com/package/cheerio) (for html parsing)
- xlsx-js-style (https://www.npmjs.com/package/xlsx-js-style) (for xlsx styling)

## Features

- 🚀 **Fast & Lightweight**: Uses Cheerio instead of Puppeteer for fast HTML parsing
- 📊 **Advanced Table Support**: Handles complex tables with `colspan` and `rowspan`
- 🎨 **Professional Styling**: Automatic styling with borders, colors, and formatting
- 📝 **Custom Titles**: Add multiple title rows with custom text and styling
- 🔧 **Flexible Configuration**: Customize appearance and behavior
- 📦 **Multiple Import Styles**: Support for class-based, functional, and namespace imports

## Installation

```bash
npm install @nodewave/table-to-xlsx
```

## Quick Start

### **Style 1: Default Import (Class-based)**
```typescript
import TableToXlsx from '@nodewave/table-to-xlsx';

const html = `
<table>
    <tr>
        <th>Name</th>
        <th>Age</th>
    </tr>
    <tr>
        <td>John</td>
        <td>30</td>
    </tr>
</table>
`;

const titleConfig = {
    numOfRows: 1,
    titles: ['User Report']
};

// Convert to file
await Html2Xlsx.convert(html, titleConfig, 'output.xlsx');

// Or convert to buffer
const buffer = await Html2Xlsx.convert(html, titleConfig);
```

### **Style 2: Namespace Import**
```typescript
import * as Html2Xlsx from '@nodewave/table-to-xlsx';

// Same usage as above
await Html2Xlsx.convert(html, titleConfig, 'output.xlsx');
```

### **Style 3: Functional Import**
```typescript
import { convert, convertToFile, convertToBuffer } from '@nodewave/table-to-xlsx';

// Direct function calls
await convert(html, titleConfig, 'output.xlsx');
await convertToFile(html, 'output.xlsx', titleConfig);
const buffer = await convertToBuffer(html, titleConfig);
```

## API Reference

### **Main Methods**

#### `convert(html: string, titleConfig: TitleConfig, outputPath?: string): Promise<string | Buffer>`

Main conversion method. If `outputPath` is provided, saves to file and returns the path. If not provided, returns a Buffer.

**Parameters:**
- `html`: HTML string containing a table
- `titleConfig`: Configuration for title rows
- `outputPath`: Optional path where the Excel file will be saved

**Returns:** Promise that resolves to output file path (string) or buffer (Buffer)

#### `convertToFile(html: string, outputPath: string, titleConfig: TitleConfig): Promise<string>`

Converts HTML to Excel and saves to file.

#### `convertToBuffer(html: string, titleConfig: TitleConfig): Promise<Buffer>`

Converts HTML to Excel and returns as buffer.

## Interfaces

### `TitleConfig`
```typescript
interface TitleConfig {
    numOfRows: number;    // Number of title rows
    titles: string[];     // Array of title texts
}
```

### `TableCell`
```typescript
interface TableCell {
    content: string;      // Cell content
    colspan: number;      // Column span
    rowspan: number;      // Row span
    isHeader: boolean;    // Whether it's a header cell
}
```

### `TableData`
```typescript
interface TableData {
    rows: TableRow[];     // Array of table rows
    maxCols: number;      // Maximum number of columns
}
```

## Styling Features

The generated Excel files include:

- **Automatic Borders**: Thin borders around all cells
- **Center Alignment**: Text centered both horizontally and vertically
- **Title Styling**: Blue background with white text for title rows
- **Header Styling**: Gray background for table headers
- **Font Sizing**: Larger fonts for titles, medium for headers, normal for data
- **Cell Merging**: Proper handling of colspan and rowspan attributes

## Examples

### Basic Table
```typescript
import { convert } from '@nodewave/table-to-xlsx';

const simpleHtml = `
<table>
    <tr><th>Name</th><th>Score</th></tr>
    <tr><td>Alice</td><td>95</td></tr>
    <tr><td>Bob</td><td>87</td></tr>
</table>
`;

await convert(simpleHtml, {
    numOfRows: 1,
    titles: ['Student Scores']
}, 'scores.xlsx');
```

### Table with Custom Titles
```typescript
import Html2Xlsx from '@nodewave/table-to-xlsx';

await Html2Xlsx.convert(html, {
    numOfRows: 2,
    titles: ['Department Report', 'Employee Performance']
}, 'report.xlsx');
```

### Get as Buffer (for web apps)
```typescript
import { convertToBuffer } from '@nodewave/table-to-xlsx';

const buffer = await convertToBuffer(html, {
    numOfRows: 1,
    titles: ['Report']
});

// Use buffer in web response or save to cloud storage
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

MIT License - see [LICENSE](LICENSE) file for details.


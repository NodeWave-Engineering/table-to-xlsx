# @nodewave/table-to-xlsx

A lightweight Node.js package to convert HTML tables to Excel files with styling, handles merged cells, and enhanced styling from HTML.

## This Package Depends on : 
- cheerio (https://www.npmjs.com/package/cheerio) (for html parsing)
- xlsx-js-style (https://www.npmjs.com/package/xlsx-js-style) (for xlsx styling)

## Features

- 🚀 **Fast & Lightweight**: Uses Cheerio instead of Puppeteer for fast HTML parsing
- 📊 **Advanced Table Support**: Handles complex tables with `colspan` and `rowspan`
- 🎨 **Professional Styling**: Automatic styling with borders, colors, and formatting
- 🎨 **Enhanced Styling**: Parse and apply CSS styles from HTML (colors, fonts, alignment, borders)
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

// Convert to file
await TableToXlsx.convert(html, 'output.xlsx');

// Or convert to buffer
const buffer = await TableToXlsx.convert(html);
```

### **Style 2: Namespace Import**
```typescript
import * as TableToXlsx from '@nodewave/table-to-xlsx';

// Same usage as above
await TableToXlsx.convert(html, 'output.xlsx');
```

### **Style 3: Functional Import**
```typescript
import { convert, convertToFile, convertToBuffer } from '@nodewave/table-to-xlsx';

// Direct function calls
await convert(html, 'output.xlsx');
await convertToFile(html, 'output.xlsx');
const buffer = await convertToBuffer(html);
```

## API Reference

### **Main Methods**

#### `convert(html: string, outputPath?: string): Promise<string | Buffer>`

Main conversion method. If `outputPath` is provided, saves to file and returns the path. If not provided, returns a Buffer.

**Parameters:**
- `html`: HTML string containing a table
- `outputPath`: Optional path where the Excel file will be saved

**Returns:** Promise that resolves to output file path (string) or buffer (Buffer)

#### `convertToFile(html: string, outputPath: string): Promise<string>`

Converts HTML to Excel and saves to file.

#### `convertToBuffer(html: string): Promise<Buffer>`

Converts HTML to Excel and returns as buffer.



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

### **Enhanced Custom Styling**

You can now apply custom styles directly in your HTML:

```html
<table>
    <tr>
        <th style="background-color: #FF6B6B; color: white; font-size: 18px; font-weight: bold; text-align: center;">
            Custom Header
        </th>
        <td style="background-color: #F7F7F7; text-align: left; font-size: 14px;">
            Left-aligned content
        </td>
    </tr>
    <tr>
        <td class="text-right font-bold" style="background-color: #E8F5E8;">
            Bold right-aligned text
        </td>
    </tr>
</table>
```

#### **Supported CSS Properties:**

- **`background-color`** - Custom cell backgrounds (hex, named colors)
- **`text-align`** - Text alignment (`left`, `center`, `right`)
- **`font-size`** - Custom font sizes (in pixels)
- **`font-weight`** - Font weight (`normal`, `bold`)
- **`color`** - Text color (hex, named colors)
- **`border-color`** - Custom border colors
- **`border-style`** - Border styles (`thin`, `medium`, `thick`)

#### **CSS Class Support:**

- **`.text-left`** - Left-aligned text
- **`.text-center`** - Center-aligned text  
- **`.text-right`** - Right-aligned text
- **`.font-bold`** - Bold text
- **`.font-normal`** - Normal weight text

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

await convert(simpleHtml, 'scores.xlsx');
```

### Get as Buffer (for web apps)
```typescript
import { convertToBuffer } from '@nodewave/table-to-xlsx';

const buffer = await convertToBuffer(html);

// Use buffer in web response or save to cloud storage
```

### Enhanced Styling Example
```typescript
import Html2Xlsx from '@nodewave/table-to-xlsx';

const styledHtml = `
<table>
    <tr>
        <th style="background-color: #2E86AB; color: white; font-size: 18px; text-align: center;">
            Product Report
        </th>
    </tr>
    <tr>
        <td class="text-left font-bold" style="background-color: #F8F9FA;">
            Product Name
        </td>
        <td class="text-center" style="background-color: #F8F9FA;">
            Sales
        </td>
    </tr>
    <tr>
        <td style="text-align: left;">Laptop</td>
        <td style="text-align: right; font-weight: bold; color: #28A745;">$1,200</td>
    </tr>
</table>
`;

await Html2Xlsx.convert(styledHtml, {
    numOfRows: 1,
    titles: ['Company Sales Report']
}, 'styled-report.xlsx');
```


### Streaming API
To handle big amount of data, says like more than 10K Rows, we recommend you to use the `createStreamProcessor` 

You can take a look at the example below : 
```ts
async function testChunkStreaming() {
    console.log('\n🔧 Testing chunk-based streaming...')

    const processor = TableToXlsx.createStreamProcessor('./chunk-streaming.xlsx', {
        chunkSize: 2000,
        onChunk: (chunk, rows) => console.log(`  📦 Chunk ${chunk}: ${rows} total rows`)
    })

    // Send header
    processor.writeHeader(`
    <thead>
        <tr>
            <th style="background-color: #007bff; color: white;">ID</th>
            <th style="background-color: #007bff; color: white;">Name</th>
            <th style="background-color: #007bff; color: white;">Value</th>
        </tr>
        <tr>
    </thead>`)

    // Send data in chunks (simulating receiving chunks from external source)
    const chunkSize = 10000
    const totalRows = 1000000

    for (let chunkStart = 1; chunkStart <= totalRows; chunkStart += chunkSize) {
        const chunkEnd = Math.min(chunkStart + chunkSize - 1, totalRows)

        // Build chunk HTML
        let chunkHtml = ''
        for (let i = chunkStart; i <= chunkEnd; i++) {
            const bgColor = i % 2 === 0 ? '#ffffff' : '#FF9E54'
            chunkHtml += `<tr><td style="background-color: ${bgColor};">${i}</td><td style="background-color: ${bgColor};">Item ${i}</td><td style="background-color: ${bgColor};">$${(Math.random() * 100).toFixed(2)}</td></tr>`
        }

        // Send chunk to processor
        processor.writeChunk(chunkHtml)
        console.log(`  📤 Sent chunk ${Math.ceil(chunkStart / chunkSize)} (rows ${chunkStart}-${chunkEnd})`)
    }

    await processor.finalize()
    console.log('✅ Chunk streaming completed!')
}
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Changelog

### v0.4.0
- Added `createStreamProcessor` , which support streaming capability via `chunking` method , - Added `createStreamProcessor`, which supports streaming capability via `chunking` method, see [Streaming API](#streaming-api)


### v0.3.0

- Address issue on bordering for col spanned column and rows 
- Address issue on font sizing 
- Implement auto fit for column to match content width and height

### v0.2.0
- 🗑️ **Removed TitleConfig**: Simplified API by removing separate title configuration
- 📝 **Direct HTML Titles**: Titles can now be set directly in HTML table headers
- 🎯 **Cleaner API**: Simplified method signatures without titleConfig parameter
- Enhanced styling system with custom CSS support
- Background colors, text alignment, font sizes, font weights
- Border customization (color and style)
- CSS class support (text-left, text-center, text-right, font-bold, font-normal)
- Named color support (red, blue, green, etc.)

### v0.1.2

- Backward compatible with existing functionality

### v0.1.1
- Initial release
- HTML table parsing with Cheerio
- Excel generation with styling
- Support for merged cells
- Multiple import styles (class-based, functional, namespace)


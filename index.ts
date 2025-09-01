import * as cheerio from 'cheerio'
import * as XLSX from 'xlsx-js-style'

export interface TableCell {
    content: string
    colspan: number
    rowspan: number
    isHeader: boolean
    // Enhanced styling properties
    styles?: {
        backgroundColor?: string
        textAlign?: 'left' | 'center' | 'right'
        fontSize?: number
        fontWeight?: 'normal' | 'bold'
        color?: string
        borderColor?: string
        borderStyle?: 'thin' | 'medium' | 'thick' | 'none'
    }
}

export interface TableRow {
    cells: TableCell[]
}

export interface TableData {
    rows: TableRow[]
    maxCols: number
}



export default class TableToXlsx {
    /**
     * Convert HTML table to Excel file
     * @param html HTML string containing a table
     * @param outputPath Optional output path (if not provided, returns buffer)
     * @returns Promise that resolves to output path or buffer
     */
    static async convert(html: string, outputPath?: string): Promise<string | Buffer> {
        try {
            // Parse the HTML table
            const tableData = await this.parseHtmlTable(html)

            // Convert to Excel-compatible data structure
            const { data: excelData, merges } = this.createExcelData(tableData)

            if (outputPath) {
                // Create and save the Excel file
                this.createExcelFile(excelData, merges, outputPath, tableData)
                return outputPath
            } else {
                // Return buffer instead of saving to file
                return this.createExcelBuffer(excelData, merges, tableData)
            }
        } catch (error) {
            throw new Error(`Failed to convert HTML to Excel: ${error}`)
        }
    }

    /**
     * Convert HTML table to Excel and save to file
     * @param html HTML string containing a table
     * @param outputPath Path where the Excel file will be saved
     * @param titleConfig Configuration for title rows
     * @returns Promise that resolves to the output file path
     */
    static async convertToFile(html: string, outputPath: string): Promise<string> {
        return this.convert(html, outputPath) as Promise<string>
    }

    /**
     * Convert HTML table to Excel and return as buffer
     * @param html HTML string containing a table
     * @param titleConfig Configuration for title rows
     * @returns Promise that resolves to buffer
     */
    static async convertToBuffer(html: string): Promise<Buffer> {
        return this.convert(html) as Promise<Buffer>
    }

    private static async parseHtmlTable(html: string): Promise<TableData> {
        const $ = cheerio.load(html)
        const table = $('table')

        if (table.length === 0) {
            throw new Error('No table found in HTML')
        }

        const rows = table.find('tr')
        const parsedRows: TableRow[] = []
        let maxCols = 0

        rows.each((rowIndex, rowElement) => {
            const cells = $(rowElement).find('th, td')
            const parsedCells: TableCell[] = []

            cells.each((cellIndex, cellElement) => {
                const $cell = $(cellElement)
                const colspan = parseInt($cell.attr('colspan') || '1')
                const rowspan = parseInt($cell.attr('rowspan') || '1')
                const isHeader = $cell.prop('tagName')?.toLowerCase() === 'th' || false
                const styles = this.parseCellStyles($cell)

                parsedCells.push({
                    content: $cell.text().trim() || '',
                    colspan: colspan,
                    rowspan: rowspan,
                    isHeader: isHeader,
                    styles: styles
                })
            })

            parsedRows.push({ cells: parsedCells })
            maxCols = Math.max(maxCols, parsedCells.length)
        })

        return { rows: parsedRows, maxCols }
    }

    private static parseCellStyles($cell: cheerio.Cheerio<any>): TableCell['styles'] {
        const styles: TableCell['styles'] = {}

        const styleAttr = $cell.attr('style')
        if (styleAttr) {
            const stylePairs = styleAttr.split(';').filter(s => s.trim())

            stylePairs.forEach(pair => {
                const [property, value] = pair.split(':').map(s => s.trim())

                switch (property) {
                    case 'background-color':
                        styles.backgroundColor = this.normalizeColor(value)
                        break
                    case 'text-align':
                        if (['left', 'center', 'right'].includes(value)) {
                            styles.textAlign = value as 'left' | 'center' | 'right'
                        }
                        break
                    case 'font-size':
                        const fontSize = parseInt(value)
                        if (!isNaN(fontSize)) {
                            styles.fontSize = fontSize
                        }
                        break
                    case 'font-weight':
                        if (value === 'bold' || value === '700') {
                            styles.fontWeight = 'bold'
                        } else if (value === 'normal' || value === '400') {
                            styles.fontWeight = 'normal'
                        }
                        break
                    case 'color':
                        styles.color = this.normalizeColor(value)
                        break
                    case 'border':
                        if (value === 'none') {
                            styles.borderStyle = 'none'
                        } else {
                            const borderParts = value.split(' ')
                            borderParts.forEach(part => {
                                if (['thin', 'medium', 'thick', 'none'].includes(part)) {
                                    styles.borderStyle = part as 'thin' | 'medium' | 'thick' | 'none'
                                } else if (part.startsWith('#') || part.match(/^[a-zA-Z]+$/)) {
                                    styles.borderColor = this.normalizeColor(value)
                                } else if (part.endsWith('px')) {
                                    const pixelValue = parseInt(part)
                                    if (!isNaN(pixelValue)) {
                                        if (pixelValue <= 1) styles.borderStyle = 'thin'
                                        else if (pixelValue <= 3) styles.borderStyle = 'medium'
                                        else styles.borderStyle = 'thick'
                                    }
                                }
                            })
                        }
                        break
                    case 'border-color':
                        styles.borderColor = this.normalizeColor(value)
                        break
                    case 'border-style':
                        if (['thin', 'medium', 'thick', 'none'].includes(value)) {
                            styles.borderStyle = value as 'thin' | 'medium' | 'thick' | 'none'
                        }
                        break
                }
            })
        }

        const classAttr = $cell.attr('class')
        if (classAttr) {
            const classes = classAttr.split(' ').filter(c => c.trim())

            classes.forEach(className => {
                switch (className.toLowerCase()) {
                    case 'text-left':
                        styles.textAlign = 'left'
                        break
                    case 'text-center':
                        styles.textAlign = 'center'
                        break
                    case 'text-right':
                        styles.textAlign = 'right'
                        break
                    case 'font-bold':
                    case 'bold':
                        styles.fontWeight = 'bold'
                        break
                    case 'font-normal':
                        styles.fontWeight = 'normal'
                        break
                    case 'border-none':
                        styles.borderStyle = 'none'
                        break
                }
            })
        }

        return Object.keys(styles).length > 0 ? styles : undefined
    }



    private static normalizeColor(color: string): string {
        if (color.startsWith('#')) {
            return color.substring(1)
        }

        if (color.startsWith('rgb(') || color.startsWith('rgba(')) {
            return '000000'
        }

        const colorMap: { [key: string]: string } = {
            'red': 'FF0000',
            'green': '00FF00',
            'blue': '0000FF',
            'yellow': 'FFFF00',
            'black': '000000',
            'white': 'FFFFFF',
            'gray': '808080',
            'grey': '808080',
            'orange': 'FFA500',
            'purple': '800080',
            'pink': 'FFC0CB',
            'brown': 'A52A2A',
            'cyan': '00FFFF',
            'magenta': 'FF00FF'
        }

        const normalizedColor = colorMap[color.toLowerCase()]
        return normalizedColor ? normalizedColor : '000000'
    }

    private static getCellData(row: number, col: number, tableData?: TableData, excelData?: any[][]): { content: string, styles?: any } | undefined {
        if (excelData && excelData[row] && excelData[row][col]) {
            const cellData = excelData[row][col]
            if (typeof cellData === 'object' && cellData !== null && 'styles' in cellData) {
                return {
                    content: cellData.content,
                    styles: cellData.styles
                }
            }
        }

        if (!tableData) return undefined

        const tableRowIndex = row
        if (tableRowIndex < 0 || tableRowIndex >= tableData.rows.length) return undefined

        const tableRow = tableData.rows[tableRowIndex]
        if (col >= tableRow.cells.length) return undefined

        return {
            content: tableRow.cells[col].content,
            styles: tableRow.cells[col].styles
        }
    }

    private static convertToSimpleData(excelData: any[][]): any[][] {
        return excelData.map(row =>
            row.map(cell => {
                if (typeof cell === 'object' && cell !== null && 'content' in cell) {
                    return cell.content
                }
                return cell
            })
        )
    }

    private static createExcelData(tableData: TableData): { data: any[][], merges: any[] } {
        const excelData: any[][] = []
        const merges: any[] = []

        for (let i = 0; i < tableData.rows.length; i++) {
            excelData[i] = new Array(tableData.maxCols).fill('')
        }

        console.log(`Processing ${tableData.rows.length} rows with maxCols: ${tableData.maxCols}`)

        tableData.rows.forEach((row, rowIndex) => {
            let currentCol = 0

            row.cells.forEach((cell) => {
                while (currentCol < tableData.maxCols && excelData[rowIndex][currentCol] !== '') {
                    currentCol++
                }

                if (currentCol >= tableData.maxCols) {
                    return
                }

                excelData[rowIndex][currentCol] = {
                    content: cell.content,
                    styles: cell.styles
                }

                if (cell.colspan > 1 || cell.rowspan > 1) {
                    merges.push({
                        s: { r: rowIndex, c: currentCol },
                        e: { r: rowIndex + cell.rowspan - 1, c: currentCol + cell.colspan - 1 }
                    })
                }

                for (let r = 0; r < cell.rowspan; r++) {
                    for (let c = 0; c < cell.colspan; c++) {
                        if (r === 0 && c === 0) continue
                        if (rowIndex + r < excelData.length && currentCol + c < excelData[0].length) {
                            excelData[rowIndex + r][currentCol + c] = {
                                content: '',
                                styles: cell.styles
                            }
                        }
                    }
                }

                currentCol += cell.colspan
            })
        })

        return { data: excelData, merges }
    }

    private static createExcelFile(excelData: any[][], merges: any[], outputPath: string, tableData?: TableData) {
        const workbook = XLSX.utils.book_new()
        const simpleData = this.convertToSimpleData(excelData)
        const worksheet = XLSX.utils.aoa_to_sheet(simpleData)

        if (merges.length > 0) {
            worksheet['!merges'] = merges
        }

        this.applyStyling(worksheet, tableData, excelData)

        const colWidths = [];
        const maxCols = tableData?.maxCols || simpleData[0]?.length || 1;

        for (let colIndex = 0; colIndex < maxCols; colIndex++) {
            let maxLength = 0;

            for (let rowIndex = 0; rowIndex < simpleData.length; rowIndex++) {
                const row = simpleData[rowIndex];
                if (row && row[colIndex]) {
                    const cellValue = String(row[colIndex]);
                    maxLength = Math.max(maxLength, cellValue.length);
                }
            }

            colWidths.push({ wch: maxLength + 2 });
        }

        worksheet['!cols'] = colWidths;

        const rowHeights = simpleData.map((row, rowIndex) => {
            let maxHeight = 15;

            row.forEach((cellValue, colIndex) => {
                if (cellValue) {
                    const cellData = this.getCellData(rowIndex, colIndex, tableData, excelData);
                    const fontSize = cellData?.styles?.fontSize || 11;
                    const fontHeight = fontSize * 1.2;
                    const contentLength = String(cellValue).length;
                    const estimatedLines = Math.ceil(contentLength / 30);
                    const contentHeight = estimatedLines * fontHeight;
                    maxHeight = Math.max(maxHeight, fontHeight, contentHeight);
                }
            });

            return { hpt: Math.max(maxHeight, 15) };
        });

        worksheet['!rows'] = rowHeights;
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')
        XLSX.writeFile(workbook, outputPath)

        console.log(`Excel file created successfully: ${outputPath}`)
        console.log(`Applied ${merges.length} cell merges`)
        console.log('Applied styling: centered alignment, borders, title formatting, and header formatting')
    }

    private static createExcelBuffer(excelData: any[][], merges: any[], tableData?: TableData): Buffer {
        // Create a new workbook
        const workbook = XLSX.utils.book_new()

        // Convert data back to simple values for Excel
        const simpleData = this.convertToSimpleData(excelData)

        // Create a worksheet from the data
        const worksheet = XLSX.utils.aoa_to_sheet(simpleData)

        // Apply the merges
        if (merges.length > 0) {
            worksheet['!merges'] = merges
        }

        // Add styling to all cells (center alignment, borders, etc.)
        this.applyStyling(worksheet, tableData, excelData)

        // Calculate column widths by looking at all cells in each column
        const colWidths = [];
        const maxCols = tableData?.maxCols || simpleData[0]?.length || 1;

        for (let colIndex = 0; colIndex < maxCols; colIndex++) {
            let maxLength = 0;

            // Check each row for this column
            for (let rowIndex = 0; rowIndex < simpleData.length; rowIndex++) {
                const row = simpleData[rowIndex];
                if (row && row[colIndex]) {
                    const cellValue = String(row[colIndex]);
                    maxLength = Math.max(maxLength, cellValue.length);
                }
            }

            colWidths.push({ wch: maxLength + 2 });
        }

        // Apply the calculated widths to the worksheet
        worksheet['!cols'] = colWidths;

        // Calculate row heights based on content and font sizes
        const rowHeights = simpleData.map((row, rowIndex) => {
            let maxHeight = 15; // Default minimum height

            row.forEach((cellValue, colIndex) => {
                if (cellValue) {
                    const cellData = this.getCellData(rowIndex, colIndex, tableData, excelData);
                    const fontSize = cellData?.styles?.fontSize || 11;

                    // Calculate height based on font size and content length
                    // Excel row height is roughly 1.2x the font size
                    const fontHeight = fontSize * 1.2;

                    // If content is long, we might need more height for wrapping
                    const contentLength = String(cellValue).length;
                    const estimatedLines = Math.ceil(contentLength / 30); // Rough estimate: 30 chars per line
                    const contentHeight = estimatedLines * fontHeight;

                    maxHeight = Math.max(maxHeight, fontHeight, contentHeight);
                }

            });

            return { hpt: Math.max(maxHeight, 15) }; // Minimum height of 15 points
        });

        // Apply the calculated heights to the worksheet
        worksheet['!rows'] = rowHeights;

        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')

        // Write to buffer
        return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
    }

    private static applyStyling(worksheet: XLSX.WorkSheet, tableData?: TableData, excelData?: any[][]) {
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')

        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C })
                const cellData = this.getCellData(R, C, tableData, excelData)
                const customStyles = cellData?.styles

                if (!worksheet[cellAddress]) {
                    worksheet[cellAddress] = { v: '' }
                }

                const cellStyle: any = {
                    alignment: {
                        horizontal: customStyles?.textAlign || 'center',
                        vertical: 'center',
                        wrapText: true
                    },
                    font: {
                        bold: customStyles?.fontWeight === 'bold' || false,
                        sz: customStyles?.fontSize || 11,
                        color: customStyles?.color ? { rgb: customStyles.color } : undefined
                    },
                    fill: {
                        fgColor: {
                            rgb: customStyles?.backgroundColor || 'FFFFFF'
                        }
                    }
                }

                if (customStyles?.borderStyle !== 'none') {
                    cellStyle.border = {
                        top: {
                            style: customStyles?.borderStyle || 'thin',
                            color: customStyles?.borderColor ? { rgb: customStyles.borderColor } : undefined
                        },
                        bottom: {
                            style: customStyles?.borderStyle || 'thin',
                            color: customStyles?.borderColor ? { rgb: customStyles.borderColor } : undefined
                        },
                        left: {
                            style: customStyles?.borderStyle || 'thin',
                            color: customStyles?.borderColor ? { rgb: customStyles.borderColor } : undefined
                        },
                        right: {
                            style: customStyles?.borderStyle || 'thin',
                            color: customStyles?.borderColor ? { rgb: customStyles.borderColor } : undefined
                        }
                    }
                }

                worksheet[cellAddress].s = cellStyle
            }
        }
    }
}

// Named exports for functional approach - properly bound to the class
export const convert = TableToXlsx.convert.bind(TableToXlsx)
export const convertToFile = TableToXlsx.convertToFile.bind(TableToXlsx)
export const convertToBuffer = TableToXlsx.convertToBuffer.bind(TableToXlsx)

// Also export the class as a named export for namespace imports
export { TableToXlsx }

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

export interface StreamOptions {
    chunkSize?: number
    onChunk?: (chunkNumber: number, processedRows: number) => void
    onComplete?: (totalRows: number, outputPath?: string) => void
    onError?: (error: Error) => void
}

export interface TableStreamProcessor {
    writeHeader: (headerHtml: string) => void
    writeRow: (rowHtml: string) => void
    writeChunk: (htmlChunk: string) => void
    finalize: () => Promise<string | Buffer>
    getProcessedRows: () => number
}



export default class TableToXlsx {
    private static readonly LARGE_TABLE_THRESHOLD = 10000 // rows
    private static readonly MAX_ROWS_WARNING = 100000

    /**
     * Convert HTML table to Excel file
     * @param html HTML string containing a table
     * @param outputPath Optional output path (if not provided, returns buffer)
     * @returns Promise that resolves to output path or buffer
     */
    static async convert(html: string, outputPath?: string): Promise<string | Buffer> {
        try {
            const tableData = await this.parseHtmlTable(html)

            // Warn about large tables
            if (tableData.rows.length > this.MAX_ROWS_WARNING) {
                console.warn(`⚠️ Large table detected: ${tableData.rows.length} rows. This may take a while and use significant memory.`)
            }

            // Use optimized processing for large tables
            if (tableData.rows.length > this.LARGE_TABLE_THRESHOLD) {
                return this.convertLargeTable(tableData, outputPath)
            }

            const { data: excelData, merges } = this.createExcelData(tableData)

            if (outputPath) {
                this.createExcelFile(excelData, merges, outputPath, tableData)
                return outputPath
            } else {
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

    /**
 * Create a true streaming processor that receives HTML data incrementally
 * This allows processing massive tables without building large HTML strings in memory
 * @param outputPath Optional output path (if not provided, returns buffer)
 * @param options Streaming options
 * @returns TableStreamProcessor for incremental data writing
 */
    static createStreamProcessor(outputPath?: string, options: StreamOptions = {}): TableStreamProcessor {
        let headerProcessed = false
        let rowCount = 0
        const allRows: TableRow[] = []
        let maxCols = 0
        let chunkNumber = 0
        const chunkSize = options.chunkSize || 1000

        const writeHeader = (headerHtml: string) => {
            if (headerProcessed) {
                throw new Error('Header already processed')
            }

            // Parse header HTML
            const $ = cheerio.load(`<table>${headerHtml}</table>`)
            $('tr').each((_, row) => {
                const cells: TableCell[] = []
                $(row).find('th, td').each((_, cell) => {
                    const $cell = $(cell)
                    cells.push({
                        content: $cell.text().trim(),
                        colspan: parseInt($cell.attr('colspan') || '1'),
                        rowspan: parseInt($cell.attr('rowspan') || '1'),
                        isHeader: true,
                        styles: this.parseCellStyles($cell)
                    })
                })
                if (cells.length > 0) {
                    allRows.push({ cells })
                    maxCols = Math.max(maxCols, cells.length)
                }
            })

            headerProcessed = true
            console.log(`[TableToXlsx] 📋 Header processed: ${allRows.length} header rows`)
        }

        const writeRow = (rowHtml: string) => {
            if (!headerProcessed) {
                throw new Error('Header must be processed before writing rows')
            }

            // Parse single row HTML
            const $ = cheerio.load(`<table><tbody>${rowHtml}</tbody></table>`)
            $('tr').each((_, row) => {
                const cells: TableCell[] = []
                $(row).find('td, th').each((_, cell) => {
                    const $cell = $(cell)
                    cells.push({
                        content: $cell.text().trim(),
                        colspan: parseInt($cell.attr('colspan') || '1'),
                        rowspan: parseInt($cell.attr('rowspan') || '1'),
                        isHeader: false,
                        styles: this.parseCellStyles($cell)
                    })
                })
                if (cells.length > 0) {
                    allRows.push({ cells })
                    maxCols = Math.max(maxCols, cells.length)
                    rowCount++
                }
            })

            // Process in chunks
            if (rowCount % chunkSize === 0) {
                chunkNumber++
                options.onChunk?.(chunkNumber, rowCount)
                console.log(`[TableToXlsx] 📦 Processed chunk ${chunkNumber} (${rowCount} rows)`)
            }
        }

        const writeChunk = (htmlChunk: string) => {
            if (!headerProcessed) {
                throw new Error('Header must be processed before writing chunks')
            }

            // Parse chunk of rows
            const $ = cheerio.load(`<table><tbody>${htmlChunk}</tbody></table>`)
            $('tr').each((_, row) => {
                const cells: TableCell[] = []
                $(row).find('td, th').each((_, cell) => {
                    const $cell = $(cell)
                    cells.push({
                        content: $cell.text().trim(),
                        colspan: parseInt($cell.attr('colspan') || '1'),
                        rowspan: parseInt($cell.attr('rowspan') || '1'),
                        isHeader: false,
                        styles: this.parseCellStyles($cell)
                    })
                })
                if (cells.length > 0) {
                    allRows.push({ cells })
                    maxCols = Math.max(maxCols, cells.length)
                    rowCount++
                }
            })

            chunkNumber++
            options.onChunk?.(chunkNumber, rowCount)

            if (rowCount % 10000 === 0) {
                console.log(`[TableToXlsx] 📦 Processed ${rowCount} rows in ${chunkNumber} chunks`)
            }
        }

        const finalize = async (): Promise<string | Buffer> => {
            if (!headerProcessed) {
                throw new Error('No header data processed')
            }

            console.log(`[TableToXlsx] ✅ Finalizing stream: ${rowCount} data rows, ${allRows.length} total rows, ${maxCols} columns`)

            const tableData: TableData = { rows: allRows, maxCols }

            try {
                // Use optimized processing for large tables
                if (tableData.rows.length > this.LARGE_TABLE_THRESHOLD) {
                    const result = await this.convertLargeTable(tableData, outputPath)
                    options.onComplete?.(tableData.rows.length, typeof result === 'string' ? result : undefined)
                    return result
                } else {
                    const result = await this.convertTableDataToExcel(tableData, outputPath)
                    options.onComplete?.(tableData.rows.length, typeof result === 'string' ? result : undefined)
                    return result
                }
            } catch (error) {
                options.onError?.(error as Error)
                throw error
            }
        }

        const getProcessedRows = (): number => {
            return rowCount
        }

        return {
            writeHeader,
            writeRow,
            writeChunk,
            finalize,
            getProcessedRows
        }
    }

    /**
     * Convert HTML table to Excel with streaming processing for large tables
     * Processes HTML in chunks to avoid memory issues with massive HTML strings
     * @param html HTML string containing a table
     * @param outputPath Optional output path (if not provided, returns buffer)
     * @param options Streaming options
     * @returns Promise that resolves to output path or buffer
     */
    static async convertStream(html: string, outputPath?: string, options: StreamOptions = {}): Promise<string | Buffer> {
        try {
            console.log('[TableToXlsx] 🚀 Using streaming HTML processing...')

            // Parse the HTML to get table structure
            const $ = cheerio.load(html)
            const table = $('table').first()

            if (table.length === 0) {
                throw new Error('No table found in HTML')
            }

            // Extract header rows
            const headerRows: TableRow[] = []
            table.find('thead tr, tr:first-child').each((_, row) => {
                const cells: TableCell[] = []
                $(row).find('th, td').each((_, cell) => {
                    const $cell = $(cell)
                    cells.push({
                        content: $cell.text().trim(),
                        colspan: parseInt($cell.attr('colspan') || '1'),
                        rowspan: parseInt($cell.attr('rowspan') || '1'),
                        isHeader: $cell.prop('tagName') === 'TH',
                        styles: this.parseCellStyles($cell)
                    })
                })
                if (cells.length > 0) {
                    headerRows.push({ cells })
                }
            })

            // Process body rows in chunks
            const chunkSize = options.chunkSize || 1000
            const allRows: TableRow[] = [...headerRows]
            let maxCols = 0
            let chunkNumber = 0

            // Get all body rows
            const bodyRows = table.find('tbody tr, tr:not(:first-child)').toArray()

            for (let i = 0; i < bodyRows.length; i += chunkSize) {
                const chunk = bodyRows.slice(i, i + chunkSize)

                chunk.forEach((row) => {
                    const cells: TableCell[] = []
                    $(row).find('td, th').each((_, cell) => {
                        const $cell = $(cell)
                        cells.push({
                            content: $cell.text().trim(),
                            colspan: parseInt($cell.attr('colspan') || '1'),
                            rowspan: parseInt($cell.attr('rowspan') || '1'),
                            isHeader: $cell.prop('tagName') === 'TH',
                            styles: this.parseCellStyles($cell)
                        })
                    })
                    if (cells.length > 0) {
                        allRows.push({ cells })
                        maxCols = Math.max(maxCols, cells.length)
                    }
                })

                chunkNumber++
                options.onChunk?.(chunkNumber, allRows.length)

                // Progress indicator
                if (allRows.length % 10000 === 0) {
                    console.log(`[TableToXlsx] 📝 Processed ${allRows.length}/${bodyRows.length + headerRows.length} rows`)
                }
            }

            const tableData: TableData = { rows: allRows, maxCols }

            console.log(`[TableToXlsx] ✅ HTML parsing completed: ${tableData.rows.length} rows, ${maxCols} columns`)

            // Use optimized processing for large tables
            if (tableData.rows.length > this.LARGE_TABLE_THRESHOLD) {
                const result = await this.convertLargeTable(tableData, outputPath)
                options.onComplete?.(tableData.rows.length, typeof result === 'string' ? result : undefined)
                return result
            } else {
                const result = await this.convertTableDataToExcel(tableData, outputPath)
                options.onComplete?.(tableData.rows.length, typeof result === 'string' ? result : undefined)
                return result
            }
        } catch (error) {
            options.onError?.(error as Error)
            throw new Error(`Failed to convert HTML to Excel with streaming: ${error}`)
        }
    }

    /**
 * Convert TableData directly to Excel with full styling (for streaming)
 */
    private static async convertTableDataToExcel(tableData: TableData, outputPath?: string): Promise<string | Buffer> {
        const workbook = XLSX.utils.book_new()

        // Create Excel data with merges
        const { data: excelData, merges } = this.createExcelData(tableData)

        // Convert to simple data for Excel (this was missing!)
        const simpleData = this.convertToSimpleData(excelData)

        // Convert to worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(simpleData)

        // Apply merges
        if (merges.length > 0) {
            worksheet['!merges'] = merges
        }

        // Apply full styling to all rows
        this.applyStyling(worksheet, tableData, excelData)

        // Calculate column widths using the existing method
        const colWidths = this.calculateOptimizedColumnWidths(simpleData, tableData.maxCols)
        worksheet['!cols'] = colWidths

        // Calculate row heights (simplified for streaming)
        const rowHeights = this.calculateRowHeightsForStreaming(excelData, tableData)
        worksheet['!rows'] = rowHeights

        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')

        if (outputPath) {
            XLSX.writeFile(workbook, outputPath)
            return outputPath
        } else {
            return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
        }
    }

    /**
     * Optimized conversion for large tables (>10k rows)
     * Uses minimal styling and optimized processing
     */
    private static async convertLargeTable(tableData: TableData, outputPath?: string): Promise<string | Buffer> {
        console.log(`[TableToXlsx] 🚀 Using optimized processing for large table (${tableData.rows.length} rows)`)

        const workbook = XLSX.utils.book_new()
        const simpleData: any[][] = []
        const merges: any[] = []

        // Process data in chunks to avoid memory issues
        const chunkSize = 1000
        let currentRow = 0

        for (let chunkStart = 0; chunkStart < tableData.rows.length; chunkStart += chunkSize) {
            const chunkEnd = Math.min(chunkStart + chunkSize, tableData.rows.length)
            const chunk = tableData.rows.slice(chunkStart, chunkEnd)

            // Process chunk
            chunk.forEach((row, rowIndex) => {
                const actualRowIndex = chunkStart + rowIndex
                const rowData: any[] = new Array(tableData.maxCols).fill('')
                let currentCol = 0

                row.cells.forEach((cell) => {
                    while (currentCol < tableData.maxCols && rowData[currentCol] !== '') {
                        currentCol++
                    }

                    if (currentCol >= tableData.maxCols) return

                    rowData[currentCol] = cell.content

                    // Only track merges for the first few rows to avoid memory issues
                    if (actualRowIndex < 100 && (cell.colspan > 1 || cell.rowspan > 1)) {
                        merges.push({
                            s: { r: actualRowIndex, c: currentCol },
                            e: { r: actualRowIndex + cell.rowspan - 1, c: currentCol + cell.colspan - 1 }
                        })
                    }

                    currentCol += cell.colspan
                })

                simpleData.push(rowData)
            })

            // Progress indicator
            if (chunkStart % (chunkSize * 10) === 0) {
                console.log(`📊 Processed ${Math.min(chunkEnd, tableData.rows.length)}/${tableData.rows.length} rows`)
            }
        }

        const worksheet = XLSX.utils.aoa_to_sheet(simpleData)

        // Apply minimal styling only to header rows
        this.applyStyling(worksheet, tableData)

        // Optimized column widths - sample-based
        const colWidths = this.calculateOptimizedColumnWidths(simpleData, tableData.maxCols)
        worksheet['!cols'] = colWidths

        // Skip row height calculation for large tables
        console.log('⏭️ Skipping row height calculation for performance')

        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')

        if (outputPath) {
            XLSX.writeFile(workbook, outputPath)
            console.log(`✅ Large table conversion completed: ${outputPath}`)
            return outputPath
        } else {
            return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
        }
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

        const colWidths = this.calculateOptimizedColumnWidths(simpleData, tableData?.maxCols || simpleData[0]?.length || 1);

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

        const colWidths = this.calculateOptimizedColumnWidths(simpleData, tableData?.maxCols || simpleData[0]?.length || 1);

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

    /**
     * Apply minimal styling only to header rows for large tables
     */
    private static applyMinimalStyling(worksheet: XLSX.WorkSheet, tableData: TableData) {
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')

        // Only style the first few rows (headers)
        const maxStyledRows = Math.min(10, range.e.r)

        for (let R = 0; R <= maxStyledRows; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C })

                if (!worksheet[cellAddress]) {
                    worksheet[cellAddress] = { v: '' }
                }

                // Basic styling for headers
                const cellStyle: any = {
                    alignment: { horizontal: 'center', vertical: 'center' },
                    font: { bold: R === 0, size: 11 },
                    fill: { fgColor: { rgb: R === 0 ? 'F0F0F0' : 'FFFFFF' } },
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    }
                }

                worksheet[cellAddress].s = cellStyle
            }
        }
    }

    /**
     * Calculate column widths using sampling for large tables
     */
    private static calculateOptimizedColumnWidths(simpleData: any[][], maxCols: number): any[] {
        const colWidths = []
        const sampleSize = Math.min(1000, simpleData.length) // Sample first 1000 rows

        for (let colIndex = 0; colIndex < maxCols; colIndex++) {
            let maxLength = 0

            // Sample rows for width calculation
            for (let rowIndex = 0; rowIndex < sampleSize; rowIndex++) {
                const row = simpleData[rowIndex]
                if (row && row[colIndex]) {
                    const cellValue = String(row[colIndex])
                    maxLength = Math.max(maxLength, cellValue.length)
                }
            }

            // Set reasonable limits
            const width = Math.min(Math.max(maxLength + 2, 10), 50)
            colWidths.push({ wch: width })
        }

        return colWidths
    }

    /**
     * Calculate row heights for streaming (simplified version)
     */
    private static calculateRowHeightsForStreaming(excelData: any[][], tableData: TableData): any[] {
        const rowHeights = []

        for (let rowIndex = 0; rowIndex < excelData.length; rowIndex++) {
            const row = excelData[rowIndex]
            let maxFontSize = 12 // Default font size

            // Find the maximum font size in this row
            for (let colIndex = 0; colIndex < row.length; colIndex++) {
                const cellData = this.getCellData(rowIndex, colIndex, tableData, excelData)
                if (cellData?.styles?.fontSize) {
                    maxFontSize = Math.max(maxFontSize, cellData.styles.fontSize)
                }
            }

            // Calculate height based on font size (simplified)
            const height = Math.max(15, maxFontSize * 1.2)
            rowHeights.push({ hpt: height })
        }

        return rowHeights
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
export const convertStream = TableToXlsx.convertStream.bind(TableToXlsx)
export const createStreamProcessor = TableToXlsx.createStreamProcessor.bind(TableToXlsx)

// Also export the class as a named export for namespace imports
export { TableToXlsx }

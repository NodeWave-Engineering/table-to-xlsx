import * as cheerio from 'cheerio'
import * as XLSX from 'xlsx-js-style'

export interface TableCell {
    content: string
    colspan: number
    rowspan: number
    isHeader: boolean
}

export interface TableRow {
    cells: TableCell[]
}

export interface TableData {
    rows: TableRow[]
    maxCols: number
}

export interface TitleConfig {
    numOfRows: number
    titles: string[]
}

export default class Html2Xlsx {
    /**
     * Convert HTML table to Excel file
     * @param html HTML string containing a table
     * @param titleConfig Configuration for title rows
     * @param outputPath Optional output path (if not provided, returns buffer)
     * @returns Promise that resolves to output path or buffer
     */
    static async convert(html: string, titleConfig: TitleConfig, outputPath?: string): Promise<string | Buffer> {
        try {
            // Parse the HTML table
            const tableData = await this.parseHtmlTable(html)

            // Convert to Excel-compatible data structure
            const { data: excelData, merges } = this.createExcelData(tableData, titleConfig)

            if (outputPath) {
                // Create and save the Excel file
                this.createExcelFile(excelData, merges, outputPath, titleConfig)
                return outputPath
            } else {
                // Return buffer instead of saving to file
                return this.createExcelBuffer(excelData, merges, titleConfig)
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
    static async convertToFile(html: string, outputPath: string, titleConfig: TitleConfig): Promise<string> {
        return this.convert(html, titleConfig, outputPath) as Promise<string>
    }

    /**
     * Convert HTML table to Excel and return as buffer
     * @param html HTML string containing a table
     * @param titleConfig Configuration for title rows
     * @returns Promise that resolves to buffer
     */
    static async convertToBuffer(html: string, titleConfig: TitleConfig): Promise<Buffer> {
        return this.convert(html, titleConfig) as Promise<Buffer>
    }

    private static async parseHtmlTable(html: string): Promise<TableData> {
        // Load HTML with Cheerio
        const $ = cheerio.load(html)

        // Find the table
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

                parsedCells.push({
                    content: $cell.text().trim() || '',
                    colspan: colspan,
                    rowspan: rowspan,
                    isHeader: isHeader
                })
            })

            parsedRows.push({ cells: parsedCells })
            maxCols = Math.max(maxCols, parsedCells.length)
        })

        return { rows: parsedRows, maxCols }
    }

    private static createExcelData(tableData: TableData, titleConfig: TitleConfig): { data: any[][], merges: any[] } {
        // Create a 2D array representing the Excel sheet
        const excelData: any[][] = []
        const merges: any[] = []

        // Add title rows
        for (let i = 0; i < titleConfig.numOfRows; i++) {
            const titleRow = new Array(tableData.maxCols).fill('')
            if (titleConfig.titles[i]) {
                titleRow[0] = titleConfig.titles[i]
                // Merge the entire row for the title
                merges.push({
                    s: { r: i, c: 0 },
                    e: { r: i, c: tableData.maxCols - 1 }
                })
            }
            excelData.push(titleRow)
        }

        // Initialize table rows with empty cells
        for (let i = 0; i < tableData.rows.length; i++) {
            const rowIndex = i + titleConfig.numOfRows
            excelData[rowIndex] = new Array(tableData.maxCols).fill('')
        }

        // Fill in the table data respecting colspan and rowspan
        tableData.rows.forEach((row, rowIndex) => {
            const actualRowIndex = rowIndex + titleConfig.numOfRows
            let currentCol = 0

            row.cells.forEach((cell) => {
                // Find the next available cell position
                while (excelData[actualRowIndex][currentCol] !== '') {
                    currentCol++
                }

                // Place the cell content
                excelData[actualRowIndex][currentCol] = cell.content

                // Add merge information for colspan/rowspan
                if (cell.colspan > 1 || cell.rowspan > 1) {
                    merges.push({
                        s: { r: actualRowIndex, c: currentCol }, // start cell
                        e: { r: actualRowIndex + cell.rowspan - 1, c: currentCol + cell.colspan - 1 } // end cell
                    })
                }

                // Mark cells that are covered by colspan/rowspan
                for (let r = 0; r < cell.rowspan; r++) {
                    for (let c = 0; c < cell.colspan; c++) {
                        if (r === 0 && c === 0) continue // Skip the main cell
                        if (actualRowIndex + r < excelData.length && currentCol + c < excelData[0].length) {
                            excelData[actualRowIndex + r][currentCol + c] = '' // Empty for merged cells
                        }
                    }
                }

                currentCol += cell.colspan
            })
        })

        return { data: excelData, merges }
    }

    private static createExcelFile(excelData: any[][], merges: any[], outputPath: string, titleConfig: TitleConfig) {
        // Create a new workbook
        const workbook = XLSX.utils.book_new()

        // Create a worksheet from the data
        const worksheet = XLSX.utils.aoa_to_sheet(excelData)

        // Apply the merges
        if (merges.length > 0) {
            worksheet['!merges'] = merges
        }

        // Add styling to all cells (center alignment, borders, etc.)
        this.applyStyling(worksheet, titleConfig)

        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')

        // Write the Excel file
        XLSX.writeFile(workbook, outputPath)

        console.log(`Excel file created successfully: ${outputPath}`)
        console.log(`Applied ${merges.length} cell merges`)
        console.log('Applied styling: centered alignment, borders, title formatting, and header formatting')
    }

    private static createExcelBuffer(excelData: any[][], merges: any[], titleConfig: TitleConfig): Buffer {
        // Create a new workbook
        const workbook = XLSX.utils.book_new()

        // Create a worksheet from the data
        const worksheet = XLSX.utils.aoa_to_sheet(excelData)

        // Apply the merges
        if (merges.length > 0) {
            worksheet['!merges'] = merges
        }

        // Add styling to all cells (center alignment, borders, etc.)
        this.applyStyling(worksheet, titleConfig)

        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')

        // Write to buffer
        return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
    }

    private static applyStyling(worksheet: XLSX.WorkSheet, titleConfig: TitleConfig) {
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')

        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C })

                if (!worksheet[cellAddress]) continue

                // Determine if this is a title row, header row, or data row
                const isTitleRow = R < titleConfig.numOfRows
                const isHeaderRow = R >= titleConfig.numOfRows && R < titleConfig.numOfRows + 2
                const isDataRow = R >= titleConfig.numOfRows + 2

                // Create cell style object
                worksheet[cellAddress].s = {
                    alignment: {
                        horizontal: 'center',
                        vertical: 'center',
                        wrapText: true
                    },
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    },
                    font: {
                        bold: isTitleRow || isHeaderRow,
                        size: isTitleRow ? 16 : (isHeaderRow ? 12 : 11)
                    },
                    fill: {
                        fgColor: {
                            rgb: isTitleRow ? '4472C4' : (isHeaderRow ? 'E6E6E6' : 'FFFFFF') // Blue for titles, gray for headers
                        }
                    }
                }

                // Special styling for title rows (white text on blue background)
                if (isTitleRow) {
                    worksheet[cellAddress].s.font.color = { rgb: 'FFFFFF' }
                }
            }
        }
    }
}

// Named exports for functional approach - properly bound to the class
export const convert = Html2Xlsx.convert.bind(Html2Xlsx)
export const convertToFile = Html2Xlsx.convertToFile.bind(Html2Xlsx)
export const convertToBuffer = Html2Xlsx.convertToBuffer.bind(Html2Xlsx)

// Also export the class as a named export for namespace imports
export { Html2Xlsx }

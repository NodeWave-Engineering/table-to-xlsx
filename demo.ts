import * as Html2Xlsx from "./index"


async function demo() {

    // Data that will be passed below
    const productData = [
        {
            name: 'Product 1',
            target: [10000, 20000, 30000, 40000],
        },
        {
            name: 'Product 2',
            target: [5000, 10000, 15000, 20000],
        }
    ]


    // The HTML Value
    const htmlValue = `
    <table>
        <thead>
            <tr>
                <th rowspan="2">No</th>
                <th colspan="4">Target</th>
                <th rowspan="2">Product Name</th>
            </tr>
            <tr>
                <th></th>
                <th>Quarter 1</th>
                <th>Quarter 2</th>
                <th>Quarter 3</th>
                <th>Quarter 4</th>
                <th></th>
            </tr>
        </thead>
        <tbody>
            ${productData.map((product, index) => `
                <tr>
                    <td>${index + 1}</td>
                    ${product.target.map((target, index) => `
                        <td>${target}</td>
                    `).join('')}
                    <td>${product.name}</td>
                </tr>
            `).join('')}
        </tbody>
    </table>
    `


    try {
        const outputPath = './tableTest.xlsx'

        // Custom title configuration
        const customTitleConfig: Html2Xlsx.TitleConfig = {
            numOfRows: 3,
            titles: ['DEMO System', 'Sales Target Report', 'Q1-Q4 2024']
        }

        // Using the new class-based API
        await Html2Xlsx.convert(htmlValue, customTitleConfig, outputPath)
        console.log('Conversion completed successfully!')
        console.log(`Excel file saved as: ${outputPath}`)
        console.log('The Excel file now contains title rows, proper table structure with merged cells, and styling!')
    } catch (error) {
        console.error('Error during conversion:', error)
    }
}


demo()
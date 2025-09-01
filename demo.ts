import * as TableToXlsx from "./index"


async function demo() {
    // The HTML Value
    const htmlValue = `
    <table>
        <thead>
            <tr>
                <th style="background-color: #2E86AB; color: white; font-size: 20px; text-align: center; border:none" colspan="4">
                    Enhanced Styling Demo
                </th>
            </tr>
            <tr>
                <td style="border:none;" colspan="4"></td>
            <tr>
                <th style="background-color: #F8F9FA; text-align: left; font-weight: bold;">Product</th>
                <th style="background-color: #F8F9FA; text-align: center; font-weight: bold;">Category</th>
                <th style="background-color: #F8F9FA; text-align: center; font-weight: bold;">Price</th>
                <th style="background-color: #F8F9FA; text-align: center; font-weight: bold;">Status</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td style="text-align: center; font-weight: bold; color: #2E86AB;">Laptop Pro - Engineers Coders Laptop Pro</td>
                <td class="text-center">Electronics</td>
                <td style="text-align: right; font-weight: bold; color: #28A745;">$1,299</td>
                <td class="text-center font-bold" style="background-color: #D4EDDA; color: #155724;">In Stock</td>
            </tr>
            <tr>
                <td style="text-align: center; font-weight: bold; color: #2E86AB;">Wireless Mouse</td>
                <td class="text-center">Accessories</td>
                <td style="text-align: right; font-weight: bold; color: #28A745;">$49</td>
                <td class="text-center font-bold" style="background-color: #F8D7DA; color: #721C24;">Low Stock</td>
            </tr>
            <tr>
                <td style="text-align: center; font-weight: bold; color: #2E86AB;">USB Cable</td>
                <td class="text-center">Accessories</td>
                <td style="text-align: right; font-weight: bold; color: #28A745;">$12</td>
                <td class="text-center font-bold" style="background-color: #D4EDDA; color: #155724;">In Stock</td>
            </tr>
        </tbody>
    </table>
    `


    try {
        const outputPath = './tableTest.xlsx'


        // Using the new class-based API
        await TableToXlsx.convert(htmlValue, outputPath)
        console.log('Conversion completed successfully!')
        console.log(`Excel file saved as: ${outputPath}`)
        console.log('The Excel file now contains title rows, proper table structure with merged cells, and styling!')
    } catch (error) {
        console.error('Error during conversion:', error)
    }
}


demo()
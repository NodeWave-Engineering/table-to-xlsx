import * as TableToXlsx from "./index"

async function testChunkStreaming() {
    console.log('\n🔧 Testing chunk-based streaming...')

    const processor = TableToXlsx.createStreamProcessor('./chunk-streaming2.xlsx', {
        chunkSize: 2000,
        onChunk: (chunk, rows) => console.log(`  📦 Chunk ${chunk}: ${rows} total rows`)
    })

    // Send header
    processor.writeHeader(/* html */`
       <table>
            <thead>
                <tr>
                    <th style="background-color: #007bff; color: white;" colspan="3">Users</th>
                </tr>
                <tr>
                    <th style="background-color: #007bff; color: white;" colspan="3">Products</th>
                </tr>
                <tr>
                    <th style="background-color: #007bff; color: white;">ID</th>
                    <th style="background-color: #007bff; color: white;">Name</th>
                    <th style="background-color: #007bff; color: white;">Value</th>
                </tr>
            </thead>
        </table>
    `)

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

// Demonstrate the difference
async function demonstrateDifference() {
    console.log('\n🔄 Comparing Traditional vs True Streaming...\n')

    console.log('❌ Traditional approach (builds massive string):')
    console.log('   let html = "<table>"')
    console.log('   for (1M rows) { html += "<tr>...</tr>" }  // 4GB memory!')
    console.log('   await convert(html)  // Memory explosion')

    console.log('\n✅ True streaming approach (constant memory):')
    console.log('   const processor = createStreamProcessor()')
    console.log('   processor.writeHeader("<thead>...")')
    console.log('   for (1M rows) { processor.writeRow("<tr>...</tr>") }  // ~50MB memory')
    console.log('   await processor.finalize()  // Memory efficient')

    console.log('\n🎯 Key Benefits:')
    console.log('   • Constant memory usage (~50MB regardless of row count)')
    console.log('   • Can process millions of rows without GC issues')
    console.log('   • Real streaming (data processed as it arrives)')
    console.log('   • Perfect for APIs, file readers, database cursors')
}

// Run tests
async function runAllTests() {
    await demonstrateDifference()
    await testChunkStreaming()
}

runAllTests().catch(console.error)

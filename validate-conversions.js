// Validation script for file conversions
// Run this in browser console or Node.js environment

async function validateConversions() {
    console.log('🔍 開始驗證檔案轉換功能...');
    
    const testResults = {
        docx: { success: false, error: null, fileSize: 0 },
        pdf: { success: false, error: null, fileSize: 0 },
        xlsx: { success: false, error: null, fileSize: 0 }
    };

    // Test data
    const testDocument = {
        title: '測試文件',
        content: `這是一個測試文件。

包含多個段落的內容測試。

中文字符測試：你好世界！
英文字符測試：Hello World!
數字測試：1234567890
特殊符號：!@#$%^&*()`
    };

    const testSpreadsheet = [
        ['產品', '價格', '數量'],
        ['iPhone', '30000', '10'],
        ['iPad', '20000', '5'],
        ['MacBook', '50000', '3']
    ];

    // Test DOCX conversion
    try {
        console.log('📄 測試 DOCX 轉換...');
        const docxBlob = await DocumentConverter.convertToDocx(
            testDocument.content, 
            testDocument.title
        );
        
        if (docxBlob.size > 1000) {
            // Verify DOCX signature (ZIP header: PK)
            const header = new Uint8Array(await docxBlob.slice(0, 4).arrayBuffer());
            const signature = Array.from(header).map(b => b.toString(16).padStart(2, '0')).join('');
            
            if (signature.startsWith('504b')) {
                testResults.docx.success = true;
                testResults.docx.fileSize = docxBlob.size;
                console.log('✅ DOCX 轉換成功 -', (docxBlob.size / 1024).toFixed(2), 'KB');
            } else {
                throw new Error('DOCX 檔案簽名不正確');
            }
        } else {
            throw new Error('DOCX 檔案過小');
        }
    } catch (error) {
        testResults.docx.error = error.message;
        console.error('❌ DOCX 轉換失敗:', error.message);
    }

    // Test PDF conversion
    try {
        console.log('📋 測試 PDF 轉換...');
        const pdfBlob = await DocumentConverter.convertToPdf(
            testDocument.content, 
            testDocument.title
        );
        
        if (pdfBlob.size > 1000) {
            // Verify PDF signature
            const header = new Uint8Array(await pdfBlob.slice(0, 5).arrayBuffer());
            const signature = new TextDecoder().decode(header);
            
            if (signature.startsWith('%PDF')) {
                testResults.pdf.success = true;
                testResults.pdf.fileSize = pdfBlob.size;
                console.log('✅ PDF 轉換成功 -', (pdfBlob.size / 1024).toFixed(2), 'KB');
            } else {
                throw new Error('PDF 檔案簽名不正確');
            }
        } else {
            throw new Error('PDF 檔案過小');
        }
    } catch (error) {
        testResults.pdf.error = error.message;
        console.error('❌ PDF 轉換失敗:', error.message);
    }

    // Test XLSX conversion
    try {
        console.log('📊 測試 XLSX 轉換...');
        const xlsxBlob = await SpreadsheetConverter.convertToExcel(
            testSpreadsheet,
            '測試表格'
        );
        
        if (xlsxBlob.size > 1000) {
            // Verify XLSX signature (ZIP header: PK)
            const header = new Uint8Array(await xlsxBlob.slice(0, 4).arrayBuffer());
            const signature = Array.from(header).map(b => b.toString(16).padStart(2, '0')).join('');
            
            if (signature.startsWith('504b')) {
                testResults.xlsx.success = true;
                testResults.xlsx.fileSize = xlsxBlob.size;
                console.log('✅ XLSX 轉換成功 -', (xlsxBlob.size / 1024).toFixed(2), 'KB');
            } else {
                throw new Error('XLSX 檔案簽名不正確');
            }
        } else {
            throw new Error('XLSX 檔案過小');
        }
    } catch (error) {
        testResults.xlsx.error = error.message;
        console.error('❌ XLSX 轉換失敗:', error.message);
    }

    // Summary
    console.log('\n📊 轉換測試結果總結:');
    console.table(testResults);
    
    const successCount = Object.values(testResults).filter(r => r.success).length;
    const totalTests = Object.keys(testResults).length;
    
    if (successCount === totalTests) {
        console.log(`🎉 所有 ${totalTests} 項轉換測試都成功！檔案轉換器已修復。`);
        return true;
    } else {
        console.log(`⚠️ ${totalTests} 項測試中有 ${successCount} 項成功，${totalTests - successCount} 項失敗。`);
        return false;
    }
}

// Export for use
if (typeof module !== 'undefined' && module.exports) {
    module.exports = validateConversions;
} else if (typeof window !== 'undefined') {
    window.validateConversions = validateConversions;
}

// Auto-run if in browser
if (typeof window !== 'undefined') {
    // Wait for libraries to load, then validate
    setTimeout(async () => {
        if (window.DocumentConverter && window.SpreadsheetConverter) {
            await validateConversions();
        }
    }, 5000);
}
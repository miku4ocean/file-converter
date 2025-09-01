// 快速功能驗證腳本 - 在瀏覽器控制台中執行

console.log('🧪 開始檔案轉換器快速測試...');

async function quickTest() {
    const results = {
        systemCheck: false,
        libraryCheck: false,
        documentConversion: false,
        spreadsheetConversion: false,
        presentationConversion: false,
        imageConversion: false
    };

    console.log('📋 1. 檢查系統組件...');
    
    // 檢查核心類別
    const coreClasses = ['LibraryLoader', 'DocumentConverter', 'SpreadsheetConverter', 'PresentationConverter', 'ImageConverter'];
    let systemComponentsOk = 0;
    
    coreClasses.forEach(className => {
        if (window[className]) {
            console.log(`✅ ${className} - 可用`);
            systemComponentsOk++;
        } else {
            console.log(`❌ ${className} - 不可用`);
        }
    });
    
    results.systemCheck = systemComponentsOk === coreClasses.length;
    console.log(`📊 系統組件檢查: ${systemComponentsOk}/${coreClasses.length} 可用`);

    console.log('\n📚 2. 檢查函式庫狀態...');
    
    if (window.libLoader) {
        const libraryStatus = window.libLoader.getLibraryStatus();
        let loadedLibs = 0;
        
        Object.entries(libraryStatus).forEach(([name, status]) => {
            const icon = status.loaded ? '✅' : '❌';
            console.log(`${icon} ${name}: ${status.description} - ${status.loaded ? '已載入' : '未載入'}`);
            if (status.loaded) loadedLibs++;
        });
        
        results.libraryCheck = loadedLibs > 0;
        console.log(`📊 函式庫狀態: ${loadedLibs}/${Object.keys(libraryStatus).length} 已載入`);
    } else {
        console.log('❌ LibraryLoader 不可用');
    }

    console.log('\n📄 3. 測試文書轉換...');
    
    try {
        const testContent = {
            title: '測試文件',
            content: '這是一個測試文件內容。\n包含多行文字。',
            originalHtml: '<h1>測試</h1><p>內容</p>'
        };
        
        // 測試 TXT 轉換
        const txtBlob = await DocumentConverter.convertToFormat(testContent, 'txt');
        console.log(`✅ TXT 轉換成功 (${txtBlob.size} bytes)`);
        
        // 測試 HTML 轉換
        const htmlBlob = await DocumentConverter.convertToFormat(testContent, 'html');
        console.log(`✅ HTML 轉換成功 (${htmlBlob.size} bytes)`);
        
        // 測試 DOCX 轉換
        const docxBlob = await DocumentConverter.convertToFormat(testContent, 'docx');
        console.log(`✅ DOCX 轉換成功 (${docxBlob.size} bytes)`);
        
        results.documentConversion = true;
        console.log('📊 文書轉換測試: 通過');
        
    } catch (error) {
        console.log(`❌ 文書轉換測試失敗: ${error.message}`);
    }

    console.log('\n📊 4. 測試表單轉換...');
    
    try {
        // 創建測試 CSV 資料
        const csvData = {
            data: [
                ['姓名', '年齡', '城市'],
                ['張三', '25', '台北'],
                ['李四', '30', '台中']
            ],
            sheets: [{ name: 'Sheet1', data: [] }],
            fileName: 'test',
            rowCount: 3,
            columnCount: 3
        };
        
        // 測試 CSV 轉換
        const csvBlob = await SpreadsheetConverter.convertToFormat(csvData, 'csv');
        console.log(`✅ CSV 轉換成功 (${csvBlob.size} bytes)`);
        
        // 測試 XLSX 轉換
        const xlsxBlob = await SpreadsheetConverter.convertToFormat(csvData, 'xlsx');
        console.log(`✅ XLSX 轉換成功 (${xlsxBlob.size} bytes)`);
        
        results.spreadsheetConversion = true;
        console.log('📊 表單轉換測試: 通過');
        
    } catch (error) {
        console.log(`❌ 表單轉換測試失敗: ${error.message}`);
    }

    console.log('\n🎯 5. 測試簡報轉換...');
    
    try {
        const presentationData = {
            title: '測試簡報',
            slides: [
                {
                    slideNumber: 1,
                    title: '第一張投影片',
                    content: ['內容項目 1', '內容項目 2'],
                    notes: '備註內容'
                }
            ],
            fileName: 'test'
        };
        
        // 測試 HTML 轉換
        const htmlBlob = await PresentationConverter.convertToFormat(presentationData, 'html');
        console.log(`✅ 簡報 HTML 轉換成功 (${htmlBlob.size} bytes)`);
        
        // 測試 TXT 轉換
        const txtBlob = await PresentationConverter.convertToFormat(presentationData, 'txt');
        console.log(`✅ 簡報 TXT 轉換成功 (${txtBlob.size} bytes)`);
        
        results.presentationConversion = true;
        console.log('📊 簡報轉換測試: 通過');
        
    } catch (error) {
        console.log(`❌ 簡報轉換測試失敗: ${error.message}`);
    }

    console.log('\n🖼️ 6. 測試圖片轉換...');
    
    try {
        // 創建一個簡單的測試畫布
        const canvas = document.createElement('canvas');
        canvas.width = 100;
        canvas.height = 100;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#007bff';
        ctx.fillRect(0, 0, 100, 100);
        ctx.fillStyle = 'white';
        ctx.font = '16px Arial';
        ctx.fillText('TEST', 30, 55);
        
        // 轉換為 Blob
        canvas.toBlob(async (blob) => {
            try {
                const file = new File([blob], 'test.png', { type: 'image/png' });
                const jpgBlob = await ImageConverter.convertImage(file, 'jpg');
                console.log(`✅ 圖片轉換成功 (${jpgBlob.size} bytes)`);
                results.imageConversion = true;
                console.log('📊 圖片轉換測試: 通過');
                
                // 完成所有測試
                printSummary();
                
            } catch (error) {
                console.log(`❌ 圖片轉換測試失敗: ${error.message}`);
                printSummary();
            }
        }, 'image/png');
        
    } catch (error) {
        console.log(`❌ 圖片轉換測試失敗: ${error.message}`);
        printSummary();
    }
    
    function printSummary() {
        console.log('\n🎉 測試總結:');
        console.log('='.repeat(40));
        
        const testItems = [
            { name: '系統組件檢查', result: results.systemCheck },
            { name: '函式庫狀態檢查', result: results.libraryCheck },
            { name: '文書轉換功能', result: results.documentConversion },
            { name: '表單轉換功能', result: results.spreadsheetConversion },
            { name: '簡報轉換功能', result: results.presentationConversion },
            { name: '圖片轉換功能', result: results.imageConversion }
        ];
        
        let passedTests = 0;
        testItems.forEach(item => {
            const icon = item.result ? '✅' : '❌';
            console.log(`${icon} ${item.name}: ${item.result ? '通過' : '失敗'}`);
            if (item.result) passedTests++;
        });
        
        const successRate = (passedTests / testItems.length) * 100;
        console.log('='.repeat(40));
        console.log(`📊 總體結果: ${passedTests}/${testItems.length} 通過 (${successRate.toFixed(1)}%)`);
        
        if (successRate >= 80) {
            console.log('🎉 測試結果良好！大部分功能正常運作。');
        } else if (successRate >= 60) {
            console.log('⚠️ 測試結果尚可，部分功能可能需要改進。');
        } else {
            console.log('❌ 測試結果不佳，建議檢查系統配置。');
        }
        
        console.log('\n💡 如需詳細測試，請開啟 automated-test.html');
    }
}

// 執行測試
quickTest().catch(error => {
    console.error('❌ 測試執行錯誤:', error);
});
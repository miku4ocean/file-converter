// 檢查 GitHub Pages 設定狀態的腳本
const https = require('https');

console.log('🔍 檢查 GitHub Pages 設定狀態...\n');

// 檢查 GitHub Pages 網站是否可訪問
function checkGitHubPages() {
    const url = 'https://miku4ocean.github.io/file-converter/';
    
    return new Promise((resolve) => {
        const req = https.get(url, (res) => {
            console.log(`📡 GitHub Pages 狀態: ${res.statusCode}`);
            
            if (res.statusCode === 200) {
                console.log('✅ GitHub Pages 已設定並可以訪問');
                console.log(`🌐 網站網址: ${url}`);
                resolve(true);
            } else if (res.statusCode === 404) {
                console.log('❌ GitHub Pages 尚未設定或部署中');
                resolve(false);
            } else {
                console.log(`⚠️  意外的狀態碼: ${res.statusCode}`);
                resolve(false);
            }
        });
        
        req.on('error', (error) => {
            console.log('❌ 無法連接到 GitHub Pages');
            console.log('原因: GitHub Pages 可能尚未設定');
            resolve(false);
        });
        
        req.setTimeout(10000, () => {
            console.log('⏰ 連接超時');
            resolve(false);
        });
    });
}

// 提供設定指引
function provideSetupInstructions() {
    console.log('\n🛠️  設定 GitHub Pages 的步驟:');
    console.log('='.repeat(50));
    console.log('1. 前往: https://github.com/miku4ocean/file-converter/settings/pages');
    console.log('2. 在 "Source" 選擇 "Deploy from a branch"');
    console.log('3. 選擇 "main" 分支');
    console.log('4. 點擊 "Save"');
    console.log('5. 等待 1-5 分鐘完成部署');
    console.log('='.repeat(50));
    console.log('\n🎯 設定完成後的網址:');
    console.log('   主要轉換器: https://miku4ocean.github.io/file-converter/');
    console.log('   PDF 測試: https://miku4ocean.github.io/file-converter/direct-pdf-test.html');
}

// 檢查本地檔案結構
function checkLocalFiles() {
    const fs = require('fs');
    const path = require('path');
    
    console.log('\n📁 檢查本地檔案結構:');
    
    const importantFiles = [
        'index.html',
        'assets/js/app.js',
        'assets/js/converters/document.js',
        'direct-pdf-test.html',
        'comprehensive-pdf-test.html'
    ];
    
    let allFilesExist = true;
    
    importantFiles.forEach(file => {
        if (fs.existsSync(file)) {
            const stats = fs.statSync(file);
            console.log(`✅ ${file} (${Math.round(stats.size/1024)}KB)`);
        } else {
            console.log(`❌ ${file} - 檔案不存在`);
            allFilesExist = false;
        }
    });
    
    if (allFilesExist) {
        console.log('\n🎉 所有重要檔案都存在，準備好部署到 GitHub Pages！');
    } else {
        console.log('\n⚠️  部分檔案缺失，可能會影響 GitHub Pages 功能');
    }
    
    return allFilesExist;
}

// 主要執行函數
async function main() {
    console.log('GitHub Pages 設定檢查工具');
    console.log('倉庫: miku4ocean/file-converter\n');
    
    // 檢查本地檔案
    const filesOk = checkLocalFiles();
    
    // 檢查 GitHub Pages 狀態
    const pagesActive = await checkGitHubPages();
    
    if (!pagesActive) {
        provideSetupInstructions();
        
        console.log('\n⏰ 設定完成後，請再次執行此腳本進行驗證:');
        console.log('   node check-github-pages.js');
    } else {
        console.log('\n🌐 GitHub Pages 設定完成！');
        console.log('🧪 建議測試項目:');
        console.log('   1. 基本檔案轉換功能');
        console.log('   2. PDF 轉檔功能');
        console.log('   3. 特殊字符處理');
        console.log('   4. 中文文字支援');
    }
    
    console.log('\n📊 狀態摘要:');
    console.log(`   本地檔案: ${filesOk ? '✅ 完整' : '❌ 不完整'}`);
    console.log(`   GitHub Pages: ${pagesActive ? '✅ 已啟用' : '❌ 未設定'}`);
}

// 執行檢查
main().catch(error => {
    console.error('執行錯誤:', error.message);
});
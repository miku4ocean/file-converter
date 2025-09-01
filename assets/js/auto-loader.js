// Auto-loader for essential libraries
(async function() {
    console.log('🚀 檔案轉換器 - 自動載入必要函式庫...');
    
    // Show loading indicator if available
    function showLoadingStatus(message) {
        console.log(`📦 ${message}`);
        
        // Try to update UI if progress elements exist
        const progressDetails = document.getElementById('progressDetails');
        if (progressDetails) {
            progressDetails.textContent = message;
        }
    }
    
    try {
        // Wait for libLoader to be available
        while (!window.libLoader) {
            await new Promise(resolve => setTimeout(resolve, 100));
        }
        
        // Essential libraries for core functionality
        const essentialLibs = [
            'jszip',    // For DOCX/XLSX creation (highest priority)
            'jspdf',    // For PDF generation
            'sheetjs',  // For Excel file processing
        ];
        
        // Optional libraries (load if possible, but don't fail if unavailable)
        const optionalLibs = [
            'mammoth',     // For DOCX parsing
            'pptxgenjs',   // For PowerPoint generation
            'html2canvas', // For HTML to image conversion
            'pdfjslib'     // For PDF parsing
        ];
        
        // Load essential libraries first
        showLoadingStatus('載入核心函式庫...');
        
        for (const libName of essentialLibs) {
            try {
                showLoadingStatus(`載入 ${libName}...`);
                await window.libLoader.loadLibrary(libName);
                console.log(`✅ ${libName} 載入成功`);
            } catch (error) {
                console.warn(`⚠️ 核心函式庫 ${libName} 載入失敗:`, error.message);
                // Continue with other libraries even if one fails
            }
        }
        
        // Load optional libraries in background
        showLoadingStatus('載入增強功能函式庫...');
        
        optionalLibs.forEach(async (libName) => {
            try {
                await window.libLoader.loadLibrary(libName);
                console.log(`✅ ${libName} 載入成功 (增強功能)`);
            } catch (error) {
                console.log(`ℹ️ 可選函式庫 ${libName} 載入失敗 (不影響基本功能):`, error.message);
            }
        });
        
        // Initialize file converter when ready
        await new Promise(resolve => {
            if (document.readyState === 'loading') {
                document.addEventListener('DOMContentLoaded', resolve);
            } else {
                resolve();
            }
        });
        
        // Initialize the main application
        if (window.FileConverter) {
            console.log('🎯 初始化檔案轉換器...');
            window.fileConverter = new FileConverter();
            console.log('✅ 檔案轉換器已就緒');
            
            // Hide loading if shown
            const progressSection = document.getElementById('progressSection');
            if (progressSection && progressSection.style.display === 'block') {
                progressSection.style.display = 'none';
            }
        }
        
        // Show library status in console
        console.log('📊 函式庫載入狀態:');
        const status = window.libLoader.getLibraryStatus();
        Object.entries(status).forEach(([name, info]) => {
            const icon = info.loaded ? '✅' : '❌';
            console.log(`${icon} ${name}: ${info.description} - ${info.loaded ? '已載入' : '未載入'}`);
        });
        
        console.log('🎉 檔案轉換器初始化完成！');
        
    } catch (error) {
        console.error('❌ 自動載入器錯誤:', error);
        
        // Still try to initialize basic functionality
        if (window.FileConverter) {
            try {
                window.fileConverter = new FileConverter();
                console.log('⚠️ 基本功能已啟動 (部分函式庫可能不可用)');
            } catch (initError) {
                console.error('❌ 無法初始化檔案轉換器:', initError);
            }
        }
    }
})();

// Global error handler for unhandled promise rejections
window.addEventListener('unhandledrejection', function(event) {
    console.warn('⚠️ 未處理的 Promise 錯誤:', event.reason);
    // Don't prevent default to allow other error handlers to run
});

// Global error handler for JavaScript errors
window.addEventListener('error', function(event) {
    console.warn('⚠️ JavaScript 錯誤:', event.error);
    // Don't prevent default to allow other error handlers to run
});

console.log('📋 自動載入器已啟動');
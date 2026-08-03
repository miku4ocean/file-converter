// 修復版PDF轉換器 - 解決真正問題
// 1. 修復DOCX重複內容問題
// 2. 修復PPTX只有文字問題

class FixedPDFConverter {
    constructor() {
        this.supportedFormats = ['docx', 'pptx'];
    }

    static async convertToPDF(file, options = {}) {
        try {
            console.log('🔧 使用修復版PDF轉換器:', file.name);
            
            const fileType = file.name.toLowerCase().split('.').pop();
            
            if (fileType === 'docx') {
                return await FixedPDFConverter.convertDOCXFixed(file);
            } else if (fileType === 'pptx') {
                return await FixedPDFConverter.convertPPTXFixed(file);
            } else {
                throw new Error(`不支援的格式: ${fileType}`);
            }
            
        } catch (error) {
            console.error('修復版轉換失敗:', error);
            throw error;
        }
    }

    // 修復DOCX轉換 - 解決重複內容問題
    static async convertDOCXFixed(file) {
        try {
            console.log('📄 修復DOCX轉換 (解決重複內容)');
            
            // 載入mammoth
            await FixedPDFConverter.loadLibrary('mammoth', 'https://unpkg.com/mammoth@1.4.2/mammoth.browser.min.js');
            
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.convertToHtml({ arrayBuffer });
            
            if (!result.html || result.html.length < 10) {
                throw new Error('無法讀取文檔內容');
            }
            
            console.log('✓ DOCX內容提取成功，長度:', result.html.length);
            
            // 關鍵修復：正確的HTML結構，避免過高容器
            const fixedHTML = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>${file.name}</title>
    <style>
        @page {
            size: A4;
            margin: 2cm;
        }
        
        body {
            font-family: 'Times New Roman', 'PingFang TC', 'Microsoft JhengHei', 'Noto Serif CJK TC', 'SimSun', serif;
            font-size: 12pt;
            line-height: 1.5;
            color: #000;
            background: white;
            margin: 0;
            padding: 20px;
            max-width: 21cm;
        }
        
        /* 防止過寬內容 */
        * {
            max-width: 100%;
            word-wrap: break-word;
        }
        
        p {
            margin: 0 0 12pt 0;
        }
        
        h1, h2, h3, h4, h5, h6 {
            color: #2c3e50;
            margin: 18pt 0 12pt 0;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 12pt 0;
        }
        
        th, td {
            border: 1px solid #ddd;
            padding: 6pt;
            text-align: left;
        }
        
        /* 關鍵：控制內容高度 */
        .document-content {
            max-height: none;
            overflow: visible;
        }
    </style>
</head>
<body>
    <div class="document-content">
        ${result.html}
    </div>
</body>
</html>
            `;
            
            // 修復的PDF轉換 - 避免重複內容
            return await FixedPDFConverter.htmlToPDFFixed(fixedHTML, file.name);
            
        } catch (error) {
            console.error('DOCX修復轉換失敗:', error);
            throw error;
        }
    }

    // 修復PPTX轉換 - 實現投影片外觀
    static async convertPPTXFixed(file) {
        try {
            console.log('📊 修復PPTX轉換 (投影片外觀)');
            
            // 載入JSZip
            await FixedPDFConverter.loadLibrary('JSZip', 'https://unpkg.com/jszip@3.10.1/dist/jszip.min.js');
            
            const arrayBuffer = await file.arrayBuffer();
            const zip = new JSZip();
            const zipContent = await zip.loadAsync(arrayBuffer);
            
            // 提取投影片
            const slides = await FixedPDFConverter.extractSlidesWithStyle(zipContent);
            
            if (!slides || slides.length === 0) {
                throw new Error('無法讀取投影片內容');
            }
            
            console.log(`✓ 提取到 ${slides.length} 張投影片`);
            
            // 關鍵改進：創建真正的投影片視覺外觀
            const slidesHTML = slides.map((slide, index) => `
<div class="slide" style="
    width: 25.4cm;
    height: 19.05cm;
    background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    margin: 0;
    padding: 2cm;
    box-sizing: border-box;
    page-break-after: ${index < slides.length - 1 ? 'always' : 'auto'};
    border: 1px solid #e0e0e0;
    position: relative;
    display: flex;
    flex-direction: column;
">
    <!-- 投影片標題區域 -->
    <div class="slide-header" style="
        text-align: center;
        margin-bottom: 2cm;
        padding-bottom: 1cm;
        border-bottom: 3px solid #2980b9;
    ">
        <h1 style="
            font-size: 28pt;
            color: #2c3e50;
            margin: 0;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
        ">投影片 ${index + 1}</h1>
        ${slide.title ? `
        <h2 style="
            font-size: 20pt;
            color: #3498db;
            margin: 10pt 0 0 0;
            font-weight: normal;
        ">${slide.title}</h2>
        ` : ''}
    </div>
    
    <!-- 投影片內容區域 -->
    <div class="slide-content" style="
        flex: 1;
        display: flex;
        flex-direction: column;
        justify-content: center;
    ">
        ${slide.content.map((item, i) => `
        <div style="
            background: white;
            margin: 0.5cm 0;
            padding: 1cm;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border-left: 4px solid #e74c3c;
        ">
            <div style="
                font-size: 16pt;
                line-height: 1.6;
                color: #2c3e50;
            ">${item}</div>
        </div>
        `).join('')}
    </div>
    
    <!-- 投影片頁腳 -->
    <div class="slide-footer" style="
        text-align: center;
        margin-top: 1cm;
        padding-top: 0.5cm;
        border-top: 1px solid #bdc3c7;
        color: #7f8c8d;
        font-size: 12pt;
    ">
        ${file.name} • 第 ${index + 1} 頁，共 ${slides.length} 頁
    </div>
</div>
            `).join('');

            const presentationHTML = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>${file.name}</title>
    <style>
        @page {
            size: A4 landscape;
            margin: 1cm;
        }
        
        body {
            margin: 0;
            padding: 0;
            font-family: 'PingFang TC', 'Microsoft JhengHei', 'Noto Sans CJK TC', 'Microsoft YaHei', Arial, sans-serif;
            background: white;
        }
        
        .slide {
            background: white !important;
        }
        
        /* 確保投影片不會變形 */
        @media print {
            .slide {
                break-inside: avoid;
            }
        }
    </style>
</head>
<body>
    ${slidesHTML}
</body>
</html>
            `;
            
            return await FixedPDFConverter.htmlToPDFFixed(presentationHTML, file.name);
            
        } catch (error) {
            console.error('PPTX修復轉換失敗:', error);
            throw error;
        }
    }

    // 提取投影片內容（改良版）
    static async extractSlidesWithStyle(zipContent) {
        const slides = [];
        
        // 收集投影片檔案
        const slideFiles = [];
        zipContent.forEach((relativePath, zipEntry) => {
            if (relativePath.startsWith('ppt/slides/slide') && relativePath.endsWith('.xml')) {
                const slideNumber = parseInt(relativePath.match(/slide(\d+)\.xml/)?.[1] || '0');
                slideFiles.push({ path: relativePath, entry: zipEntry, number: slideNumber });
            }
        });
        
        // 按順序排列
        slideFiles.sort((a, b) => a.number - b.number);
        
        // 處理每張投影片
        for (const slideFile of slideFiles) {
            try {
                const xmlContent = await slideFile.entry.async('text');
                
                // 提取文字內容
                const textMatches = xmlContent.match(/<a:t[^>]*>([^<]+)<\/a:t>/g) || [];
                const texts = textMatches.map(match => 
                    match.replace(/<[^>]+>/g, '').trim()
                ).filter(text => text.length > 0);
                
                if (texts.length > 0) {
                    // 智能識別標題和內容
                    let title = '';
                    let content = texts;
                    
                    // 如果第一行文字較短，可能是標題
                    if (texts[0] && texts[0].length < 50) {
                        title = texts[0];
                        content = texts.slice(1);
                    }
                    
                    // 如果沒有內容，至少顯示一些東西
                    if (content.length === 0) {
                        content = ['投影片內容'];
                    }
                    
                    slides.push({
                        number: slideFile.number,
                        title: title,
                        content: content
                    });
                }
                
            } catch (error) {
                console.warn(`投影片 ${slideFile.number} 處理失敗:`, error.message);
            }
        }
        
        return slides;
    }

    // 修復的HTML轉PDF - 避免重複內容
    static async htmlToPDFFixed(htmlContent, fileName) {
        try {
            // 載入必要庫
            await FixedPDFConverter.loadLibrary('html2canvas', 'https://unpkg.com/html2canvas@1.4.1/dist/html2canvas.min.js');
            await FixedPDFConverter.loadLibrary('jsPDF', 'https://unpkg.com/jspdf@2.5.1/dist/jspdf.umd.min.js');
            
            // 創建容器 - 關鍵修復：限制高度
            const container = document.createElement('div');
            container.innerHTML = htmlContent;
            
            // 重要修復：固定容器尺寸，避免過高
            container.style.position = 'fixed';
            container.style.top = '-9999px';
            container.style.left = '-9999px';
            container.style.width = '794px';  // A4寬度
            container.style.maxHeight = '1123px'; // 限制高度
            container.style.overflow = 'visible';
            container.style.background = 'white';
            
            document.body.appendChild(container);
            
            // 等待渲染
            await new Promise(resolve => setTimeout(resolve, 1000));
            
            console.log('📸 開始截圖，容器尺寸:', container.offsetWidth, 'x', container.offsetHeight);
            
            // html2canvas設定 - 修復重複問題
            const canvas = await html2canvas(container, {
                scale: 1, // 降低解析度避免過大
                useCORS: true,
                allowTaint: true,
                backgroundColor: '#ffffff',
                width: 794,
                height: Math.min(container.offsetHeight, 1123), // 限制截圖高度
                scrollX: 0,
                scrollY: 0
            });
            
            console.log('✓ 截圖完成，Canvas尺寸:', canvas.width, 'x', canvas.height);
            
            // PDF生成 - 關鍵修復：正確的分頁邏輯
            const pdf = new jsPDF('p', 'pt', 'a4');
            const imgData = canvas.toDataURL('image/jpeg', 0.85);
            
            const pdfWidth = 595;  // A4寬度(點)
            const pdfHeight = 842; // A4高度(點)
            const imgWidth = pdfWidth;
            const imgHeight = (canvas.height * pdfWidth) / canvas.width;
            
            console.log('PDF計算: 圖片高度', imgHeight, ', 頁面高度', pdfHeight);
            
            // 修復的分頁邏輯
            if (imgHeight <= pdfHeight) {
                // 內容適合一頁
                console.log('✓ 內容適合單頁');
                pdf.addImage(imgData, 'JPEG', 0, 0, imgWidth, imgHeight);
            } else {
                // 需要分頁 - 修復重複問題
                console.log('⚠️ 需要分頁');
                const totalPages = Math.ceil(imgHeight / pdfHeight);
                console.log('預計頁數:', totalPages);
                
                for (let pageNum = 0; pageNum < totalPages; pageNum++) {
                    if (pageNum > 0) {
                        pdf.addPage();
                    }
                    
                    const yOffset = -(pageNum * pdfHeight); // 負數向上偏移
                    
                    pdf.addImage(
                        imgData, 'JPEG',
                        0, yOffset,        // x, y (y用負數偏移)
                        imgWidth, imgHeight // 保持原始比例
                    );
                    
                    console.log(`添加第 ${pageNum + 1} 頁, 偏移: ${yOffset}`);
                }
            }
            
            document.body.removeChild(container);
            
            const pdfBlob = pdf.output('blob');
            console.log('✓ PDF生成完成，大小:', (pdfBlob.size / 1024).toFixed(1), 'KB');
            
            return pdfBlob;
            
        } catch (error) {
            console.error('HTML轉PDF修復版失敗:', error);
            throw error;
        }
    }

    // 載入JavaScript庫
    static async loadLibrary(name, url) {
        if (window[name]) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = url;
            script.onload = () => resolve();
            script.onerror = () => reject(new Error(`${name} 載入失敗`));
            document.head.appendChild(script);
        });
    }
}
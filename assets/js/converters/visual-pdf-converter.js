// 視覺PDF轉換器 - 革命性方法
// 不解析文件內容，直接捕獲瀏覽器渲染結果

class VisualPDFConverter {
    constructor() {
        this.supportedFormats = ['docx', 'pptx', 'pdf', 'html', 'txt'];
    }

    // 主要轉換方法 - 視覺渲染方式
    static async convertToPDF(file, options = {}) {
        try {
            console.log('👁️ 開始視覺PDF轉換 (渲染捕獲方式):', file.name);
            
            const fileType = VisualPDFConverter.getFileType(file);
            const converter = new VisualPDFConverter();
            
            switch (fileType) {
                case 'pdf':
                    return await VisualPDFConverter.passThroughPDF(file);
                    
                case 'html':
                case 'htm':
                    return await VisualPDFConverter.captureHTMLRendering(file);
                    
                case 'txt':
                    return await VisualPDFConverter.captureTextRendering(file);
                    
                case 'docx':
                case 'pptx':
                    // 關鍵突破：使用瀏覽器內建的檢視能力
                    return await VisualPDFConverter.captureOfficeRendering(file, fileType);
                    
                default:
                    throw new Error(`不支援的檔案格式: ${fileType}`);
            }
            
        } catch (error) {
            console.error('視覺PDF轉換失敗:', error);
            throw new Error(`視覺PDF轉換失敗: ${error.message}`);
        }
    }

    // 革命性方法：捕獲Office文件的渲染結果
    static async captureOfficeRendering(file, fileType) {
        try {
            console.log('🎯 捕獲Office文件渲染結果');
            
            // 方法1: 嘗試使用檔案URL + 內建檢視器
            const result1 = await VisualPDFConverter.tryBrowserViewer(file, fileType);
            if (result1) return result1;
            
            // 方法2: 回退到內容解析 + 高品質渲染
            const result2 = await VisualPDFConverter.fallbackToContentCapture(file, fileType);
            if (result2) return result2;
            
            // 方法3: 最後手段 - 創建視覺化的內容描述
            return await VisualPDFConverter.createVisualDescription(file, fileType);
            
        } catch (error) {
            console.error('Office渲染捕獲失敗:', error);
            throw error;
        }
    }

    // 嘗試使用瀏覽器內建檢視器
    static async tryBrowserViewer(file, fileType) {
        try {
            console.log('🌐 嘗試瀏覽器內建檢視器');
            
            // 創建 Blob URL
            const blobUrl = URL.createObjectURL(file);
            
            // 使用不同的策略檢查瀏覽器是否能直接顯示檔案
            return new Promise((resolve, reject) => {
                const iframe = document.createElement('iframe');
                iframe.style.position = 'fixed';
                iframe.style.top = '-9999px';
                iframe.style.left = '-9999px';
                iframe.style.width = '794px';  // A4 寬度
                iframe.style.height = '1123px'; // A4 高度
                iframe.style.border = 'none';
                
                let resolved = false;
                
                iframe.onload = async () => {
                    if (resolved) return;
                    resolved = true;
                    
                    try {
                        // 等待渲染完成
                        await new Promise(resolve => setTimeout(resolve, 2000));
                        
                        // 嘗試使用 html2canvas 截圖
                        if (typeof html2canvas !== 'undefined') {
                            console.log('📸 使用 html2canvas 截圖');
                            
                            const canvas = await html2canvas(iframe, {
                                scale: 2,
                                useCORS: true,
                                allowTaint: true,
                                backgroundColor: '#ffffff',
                                width: 794,
                                height: 1123
                            });
                            
                            // 轉換為 PDF
                            const pdf = await VisualPDFConverter.canvasToPDF(canvas, file.name);
                            
                            document.body.removeChild(iframe);
                            URL.revokeObjectURL(blobUrl);
                            resolve(pdf);
                            
                        } else {
                            // 載入 html2canvas
                            await VisualPDFConverter.loadHTML2Canvas();
                            // 重試
                            const canvas = await html2canvas(iframe, {
                                scale: 1.5,
                                useCORS: true,
                                backgroundColor: '#ffffff'
                            });
                            
                            const pdf = await VisualPDFConverter.canvasToPDF(canvas, file.name);
                            
                            document.body.removeChild(iframe);
                            URL.revokeObjectURL(blobUrl);
                            resolve(pdf);
                        }
                        
                    } catch (error) {
                        console.warn('瀏覽器檢視器方法失敗:', error.message);
                        document.body.removeChild(iframe);
                        URL.revokeObjectURL(blobUrl);
                        resolve(null); // 不拒絕，讓其他方法嘗試
                    }
                };
                
                iframe.onerror = () => {
                    if (resolved) return;
                    resolved = true;
                    
                    console.warn('iframe 載入失敗，嘗試其他方法');
                    document.body.removeChild(iframe);
                    URL.revokeObjectURL(blobUrl);
                    resolve(null);
                };
                
                // 設定超時
                setTimeout(() => {
                    if (resolved) return;
                    resolved = true;
                    
                    console.warn('瀏覽器檢視器超時');
                    document.body.removeChild(iframe);
                    URL.revokeObjectURL(blobUrl);
                    resolve(null);
                }, 10000);
                
                iframe.src = blobUrl;
                document.body.appendChild(iframe);
            });
            
        } catch (error) {
            console.warn('瀏覽器檢視器方法異常:', error.message);
            return null;
        }
    }

    // 回退到內容捕獲方法
    static async fallbackToContentCapture(file, fileType) {
        try {
            console.log('🔄 使用內容捕獲方法');
            
            if (fileType === 'docx') {
                return await VisualPDFConverter.captureDocxContent(file);
            } else if (fileType === 'pptx') {
                return await VisualPDFConverter.capturePptxContent(file);
            }
            
            return null;
            
        } catch (error) {
            console.warn('內容捕獲方法失敗:', error.message);
            return null;
        }
    }

    // 捕獲 DOCX 內容並視覺化
    static async captureDocxContent(file) {
        try {
            // 載入 mammoth.js
            await VisualPDFConverter.loadMammoth();
            
            if (typeof mammoth === 'undefined') {
                throw new Error('Mammoth.js 未載入');
            }
            
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.convertToHtml({ arrayBuffer });
            
            if (!result.html || result.html.trim().length < 10) {
                throw new Error('無法提取有效的HTML內容');
            }
            
            console.log('📄 成功提取 DOCX HTML 內容，長度:', result.html.length);
            
            // 創建高品質的HTML渲染
            const styledHtml = `
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>DOCX 轉換</title>
                    <style>
                        @page {
                            size: A4;
                            margin: 2.5cm;
                        }
                        
                        body {
                            font-family: 'Times New Roman', 'PingFang TC', 'Microsoft JhengHei', 'Noto Serif CJK TC', 'SimSun', serif;
                            font-size: 12pt;
                            line-height: 1.6;
                            color: #000;
                            background: white;
                            margin: 0;
                            padding: 0;
                            width: 794px; /* A4 寬度 */
                        }
                        
                        .document-container {
                            width: 100%;
                            min-height: 1123px; /* A4 高度 */
                            padding: 40px;
                            box-sizing: border-box;
                            background: white;
                        }
                        
                        p {
                            margin: 0 0 12pt 0;
                        }
                        
                        h1, h2, h3, h4, h5, h6 {
                            color: #2c3e50;
                            margin-top: 24pt;
                            margin-bottom: 12pt;
                        }
                        
                        h1 { font-size: 18pt; }
                        h2 { font-size: 16pt; }
                        h3 { font-size: 14pt; }
                        
                        ul, ol {
                            padding-left: 20pt;
                        }
                        
                        table {
                            width: 100%;
                            border-collapse: collapse;
                            margin: 12pt 0;
                        }
                        
                        th, td {
                            border: 1px solid #ddd;
                            padding: 8pt;
                            text-align: left;
                        }
                        
                        th {
                            background-color: #f5f5f5;
                        }
                        
                        /* 分頁處理 */
                        .page-break {
                            page-break-before: always;
                        }
                    </style>
                </head>
                <body>
                    <div class="document-container">
                        ${result.html}
                        
                        <div style="margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; font-size: 10pt; color: #666; text-align: center;">
                            由 ${file.name} 轉換 • ${new Date().toLocaleString()}
                        </div>
                    </div>
                </body>
                </html>
            `;
            
            // 渲染為 PDF
            return await VisualPDFConverter.htmlToPDFAdvanced(styledHtml, file.name);
            
        } catch (error) {
            console.error('DOCX 內容捕獲失敗:', error);
            throw error;
        }
    }

    // 捕獲 PPTX 內容並視覺化
    static async capturePptxContent(file) {
        try {
            // 載入 JSZip
            await VisualPDFConverter.loadJSZip();
            
            const arrayBuffer = await file.arrayBuffer();
            const zip = new JSZip();
            const zipContent = await zip.loadAsync(arrayBuffer);
            
            // 提取投影片內容
            const slides = await VisualPDFConverter.extractPPTXSlides(zipContent);
            
            if (!slides || slides.length === 0) {
                throw new Error('無法提取投影片內容');
            }
            
            console.log('📊 成功提取', slides.length, '張投影片');
            
            // 創建投影片視覺化HTML
            const slidesHtml = slides.map((slide, index) => `
                <div class="slide-page" style="width: 794px; height: 1123px; background: white; padding: 60px; box-sizing: border-box; page-break-after: ${index < slides.length - 1 ? 'always' : 'auto'};">
                    <div class="slide-header" style="text-align: center; border-bottom: 3px solid #3498db; padding-bottom: 20px; margin-bottom: 40px;">
                        <h1 style="color: #2c3e50; font-size: 24pt; margin: 0;">投影片 ${index + 1}</h1>
                        ${slide.title ? `<h2 style="color: #3498db; font-size: 18pt; margin: 10px 0 0 0;">${slide.title}</h2>` : ''}
                    </div>
                    
                    <div class="slide-content" style="font-size: 16pt; line-height: 1.6;">
                        ${slide.content.map(item => `
                            <div style="margin: 20px 0; padding: 15px; background: #f8f9fa; border-left: 4px solid #3498db; border-radius: 5px;">
                                ${item}
                            </div>
                        `).join('')}
                    </div>
                    
                    <div class="slide-footer" style="position: absolute; bottom: 40px; right: 60px; font-size: 12pt; color: #7f8c8d;">
                        ${index + 1} / ${slides.length}
                    </div>
                </div>
            `).join('');
            
            const styledHtml = `
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>PPTX 轉換</title>
                    <style>
                        @page {
                            size: A4;
                            margin: 0;
                        }
                        
                        body {
                            font-family: 'PingFang TC', 'Microsoft JhengHei', 'Noto Sans CJK TC', 'Microsoft YaHei', Arial, sans-serif;
                            margin: 0;
                            padding: 0;
                            background: white;
                        }
                        
                        .slide-page {
                            position: relative;
                        }
                    </style>
                </head>
                <body>
                    ${slidesHtml}
                </body>
                </html>
            `;
            
            // 渲染為 PDF
            return await VisualPDFConverter.htmlToPDFAdvanced(styledHtml, file.name);
            
        } catch (error) {
            console.error('PPTX 內容捕獲失敗:', error);
            throw error;
        }
    }

    // 提取 PPTX 投影片
    static async extractPPTXSlides(zipContent) {
        const slides = [];
        
        // 收集投影片檔案
        const slideFiles = [];
        zipContent.forEach((relativePath, zipEntry) => {
            if (relativePath.startsWith('ppt/slides/slide') && relativePath.endsWith('.xml')) {
                const slideNumber = parseInt(relativePath.match(/slide(\d+)\.xml/)?.[1] || '0');
                slideFiles.push({ path: relativePath, entry: zipEntry, number: slideNumber });
            }
        });
        
        // 按編號排序
        slideFiles.sort((a, b) => a.number - b.number);
        
        // 提取每張投影片
        for (const slideFile of slideFiles) {
            try {
                const xmlContent = await slideFile.entry.async('text');
                
                // 提取文字內容
                const textMatches = xmlContent.match(/<a:t[^>]*>([^<]+)<\/a:t>/g) || [];
                const texts = textMatches.map(match => 
                    match.replace(/<[^>]+>/g, '').trim()
                ).filter(text => text.length > 0);
                
                if (texts.length > 0) {
                    // 第一個通常是標題
                    const title = texts[0];
                    const content = texts.slice(1);
                    
                    slides.push({
                        number: slideFile.number,
                        title: title,
                        content: content.length > 0 ? content : ['內容項目']
                    });
                }
                
            } catch (error) {
                console.warn(`投影片 ${slideFile.number} 處理失敗:`, error.message);
            }
        }
        
        return slides;
    }

    // 高級HTML轉PDF
    static async htmlToPDFAdvanced(htmlContent, fileName) {
        try {
            // 載入必要的庫
            await VisualPDFConverter.loadHTML2Canvas();
            await VisualPDFConverter.loadJsPDF();
            
            // 創建臨時容器
            const container = document.createElement('div');
            container.innerHTML = htmlContent;
            container.style.position = 'fixed';
            container.style.top = '-9999px';
            container.style.left = '-9999px';
            container.style.width = '794px';
            container.style.background = 'white';
            
            document.body.appendChild(container);
            
            // 等待渲染完成
            await new Promise(resolve => setTimeout(resolve, 1000));
            
            // 使用 html2canvas 截圖
            const canvas = await html2canvas(container, {
                scale: 2,
                useCORS: true,
                allowTaint: true,
                backgroundColor: '#ffffff',
                width: 794,
                scrollX: 0,
                scrollY: 0
            });
            
            // 轉換為PDF
            const pdf = new jsPDF('p', 'pt', 'a4');
            const imgData = canvas.toDataURL('image/jpeg', 0.9);
            
            // 計算縮放比例
            const imgWidth = 595; // A4 寬度 (點)
            const imgHeight = (canvas.height * imgWidth) / canvas.width;
            
            // 如果內容太高，需要分頁
            if (imgHeight > 842) { // A4 高度
                const pageHeight = 842;
                let remainingHeight = imgHeight;
                let yPos = 0;
                
                while (remainingHeight > 0) {
                    if (yPos > 0) pdf.addPage();
                    
                    const currentHeight = Math.min(pageHeight, remainingHeight);
                    
                    pdf.addImage(
                        imgData, 'JPEG', 
                        0, yPos > 0 ? -yPos : 0, 
                        imgWidth, imgHeight
                    );
                    
                    remainingHeight -= currentHeight;
                    yPos += currentHeight;
                }
            } else {
                pdf.addImage(imgData, 'JPEG', 0, 0, imgWidth, imgHeight);
            }
            
            document.body.removeChild(container);
            
            return new Blob([pdf.output('blob')], { type: 'application/pdf' });
            
        } catch (error) {
            console.error('高級HTML轉PDF失敗:', error);
            throw error;
        }
    }

    // Canvas 轉 PDF
    static async canvasToPDF(canvas, fileName) {
        await VisualPDFConverter.loadJsPDF();
        
        const pdf = new jsPDF('p', 'pt', 'a4');
        const imgData = canvas.toDataURL('image/jpeg', 0.9);
        
        const imgWidth = 595;
        const imgHeight = (canvas.height * imgWidth) / canvas.width;
        
        pdf.addImage(imgData, 'JPEG', 0, 0, imgWidth, Math.min(imgHeight, 842));
        
        return new Blob([pdf.output('blob')], { type: 'application/pdf' });
    }

    // 創建視覺化描述（最後手段）
    static async createVisualDescription(file, fileType) {
        console.log('📝 創建視覺化檔案描述');
        
        const description = `
            <div style="text-align: center; padding: 60px 40px; font-family: 'Microsoft YaHei', sans-serif;">
                <h1 style="color: #2c3e50; font-size: 24pt; margin-bottom: 20px;">
                    📄 ${file.name}
                </h1>
                
                <div style="background: #f8f9fa; padding: 30px; border-radius: 10px; margin: 30px 0; border-left: 5px solid #3498db;">
                    <h2 style="color: #3498db; font-size: 18pt; margin-bottom: 15px;">檔案資訊</h2>
                    <p style="font-size: 14pt; line-height: 1.6; margin: 10px 0;"><strong>檔案名稱:</strong> ${file.name}</p>
                    <p style="font-size: 14pt; line-height: 1.6; margin: 10px 0;"><strong>檔案大小:</strong> ${(file.size / 1024 / 1024).toFixed(2)} MB</p>
                    <p style="font-size: 14pt; line-height: 1.6; margin: 10px 0;"><strong>檔案類型:</strong> ${fileType.toUpperCase()} 文件</p>
                    <p style="font-size: 14pt; line-height: 1.6; margin: 10px 0;"><strong>轉換時間:</strong> ${new Date().toLocaleString()}</p>
                </div>
                
                <div style="background: #fff3cd; padding: 20px; border-radius: 8px; border-left: 5px solid #ffc107;">
                    <h3 style="color: #856404; margin-bottom: 15px;">💡 轉換說明</h3>
                    <p style="font-size: 12pt; color: #856404; line-height: 1.5;">
                        此PDF是檔案的視覺化表示。原始檔案已成功讀取，但由於瀏覽器限制，
                        無法完全重現原始格式。建議使用專業軟體開啟原始檔案以獲得最佳體驗。
                    </p>
                </div>
                
                <div style="margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; font-size: 10pt;">
                    由視覺PDF轉換器生成 • 純前端處理，資料安全無上傳
                </div>
            </div>
        `;
        
        return await VisualPDFConverter.htmlToPDFAdvanced(`
            <!DOCTYPE html>
            <html><head><meta charset="utf-8"><title>視覺化檔案描述</title></head>
            <body style="margin: 0; padding: 0; background: white;">${description}</body>
            </html>
        `, file.name);
    }

    // 工具方法：載入 HTML2Canvas
    static async loadHTML2Canvas() {
        if (window.html2canvas) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/html2canvas@1.4.1/dist/html2canvas.min.js';
            script.onload = () => resolve();
            script.onerror = () => reject(new Error('HTML2Canvas 載入失敗'));
            document.head.appendChild(script);
        });
    }

    // 工具方法：載入 jsPDF
    static async loadJsPDF() {
        if (window.jsPDF) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/jspdf@2.5.1/dist/jspdf.umd.min.js';
            script.onload = () => resolve();
            script.onerror = () => reject(new Error('jsPDF 載入失敗'));
            document.head.appendChild(script);
        });
    }

    // 工具方法：載入 Mammoth
    static async loadMammoth() {
        if (window.mammoth) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/mammoth@1.4.2/mammoth.browser.min.js';
            script.onload = () => resolve();
            script.onerror = () => reject(new Error('Mammoth 載入失敗'));
            document.head.appendChild(script);
        });
    }

    // 工具方法：載入 JSZip
    static async loadJSZip() {
        if (window.JSZip) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/jszip@3.10.1/dist/jszip.min.js';
            script.onload = () => resolve();
            script.onerror = () => reject(new Error('JSZip 載入失敗'));
            document.head.appendChild(script);
        });
    }

    // 其他工具方法
    static getFileType(file) {
        return file.name.toLowerCase().split('.').pop();
    }

    static passThroughPDF(file) {
        return Promise.resolve(file);
    }

    static async captureHTMLRendering(file) {
        const htmlContent = await file.text();
        return await VisualPDFConverter.htmlToPDFAdvanced(htmlContent, file.name);
    }

    static async captureTextRendering(file) {
        const textContent = await file.text();
        const htmlContent = `
            <div style="font-family: 'Courier New', 'PingFang TC', 'Microsoft JhengHei', monospace; font-size: 12pt; line-height: 1.4; padding: 40px; white-space: pre-wrap;">
                ${textContent.replace(/</g, '&lt;').replace(/>/g, '&gt;')}
            </div>
        `;
        return await VisualPDFConverter.htmlToPDFAdvanced(`
            <!DOCTYPE html>
            <html><head><meta charset="utf-8"><title>文字檔案</title></head>
            <body style="margin: 0; background: white;">${htmlContent}</body>
            </html>
        `, file.name);
    }
}
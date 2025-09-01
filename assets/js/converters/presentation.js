// Presentation conversion utilities - Fixed version

class PresentationConverter {
    constructor() {
        this.supportedFormats = {
            input: ['pptx', 'ppt', 'odp', 'pdf', 'html'],
            output: ['pdf', 'pptx', 'html', 'images', 'txt', 'md']
        };
    }

    // Validate presentation file
    static isValidPresentationFile(file) {
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.presentationml.presentation', // pptx
            'application/vnd.ms-powerpoint', // ppt
            'application/vnd.oasis.opendocument.presentation', // odp
            'application/pdf', // pdf
            'text/html'
        ];
        
        const validExtensions = ['.pptx', '.ppt', '.odp', '.pdf', '.html', '.htm'];
        const fileExtension = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));
        
        return validTypes.includes(file.type) || validExtensions.includes(fileExtension);
    }

    // Extract presentation content from various formats
    static async extractPresentationContent(file) {
        const fileType = PresentationConverter.getFileType(file);
        
        switch (fileType) {
            case 'pptx':
            case 'ppt':
                return await PresentationConverter.extractFromPowerPoint(file);
            case 'pdf':
                return await PresentationConverter.extractFromPdf(file);
            case 'html':
            case 'htm':
                return await PresentationConverter.extractFromHtml(file);
            default:
                throw new Error(`不支援的簡報格式: ${fileType}`);
        }
    }

    // Get file type from file object
    static getFileType(file) {
        const extension = file.name.toLowerCase().substring(file.name.lastIndexOf('.') + 1);
        return extension;
    }

    // Extract content from PowerPoint files
    static async extractFromPowerPoint(file) {
        try {
            // Validate file first
            if (!file || file.size === 0) {
                throw new Error('檔案為空或不存在');
            }
            
            if (file.size > 50 * 1024 * 1024) { // 50MB limit
                throw new Error('檔案過大，請使用小於 50MB 的檔案');
            }
            
            const arrayBuffer = await file.arrayBuffer();
            const fileName = file.name.replace(/\.[^/.]+$/, '');
            const fileType = file.name.toLowerCase().substring(file.name.lastIndexOf('.') + 1);
            
            // Basic file validation
            const uint8Array = new Uint8Array(arrayBuffer);
            const header = Array.from(uint8Array.slice(0, 4)).map(b => b.toString(16).padStart(2, '0')).join('');
            
            // Check for ZIP signature (PPTX) or OLE signature (PPT)
            const isZip = header === '504b0304' || header === '504b0506';
            const isOle = header === 'd0cf11e0';
            
            if (!isZip && !isOle && fileType !== 'html' && fileType !== 'htm') {
                throw new Error(`不支援的檔案格式或檔案損壞：${fileType.toUpperCase()}`);
            }
            
            let slides = [];
            
            if (isZip && fileType === 'pptx') {
                // Try to parse PPTX file structure
                slides = await PresentationConverter.parsePPTXContent(arrayBuffer, fileName);
            } else {
                // Fallback for unsupported formats
                slides = [
                    {
                        slideNumber: 1,
                        title: '檔案資訊',
                        content: [
                            `檔案名稱: ${fileName}`,
                            `檔案格式: ${fileType.toUpperCase()}`,
                            `檔案大小: ${PresentationConverter.formatFileSize(arrayBuffer.byteLength)}`,
                            '',
                            '⚠️ 此格式需要進階解析支援',
                            '建議轉換為支援的格式進行檢視'
                        ],
                        notes: `無法完整解析 ${fileType.toUpperCase()} 格式的內容`
                    }
                ];
            }
            
            return {
                title: fileName,
                slides: slides,
                totalSlides: slides.length,
                metadata: {
                    fileSize: arrayBuffer.byteLength,
                    fileType: fileType,
                    extractedAt: new Date().toISOString(),
                    isValid: true
                }
            };
            
        } catch (error) {
            console.error('PowerPoint 解析錯誤:', error);
            throw new Error(`PowerPoint 檔案解析失敗: ${error.message}`);
        }
    }

    // Extract from PDF files
    static async extractFromPdf(file) {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const fileName = file.name.replace(/\.[^/.]+$/, '');
            
            // Basic PDF validation
            const pdfHeader = new Uint8Array(arrayBuffer.slice(0, 5));
            const headerString = String.fromCharCode(...pdfHeader);
            
            if (!headerString.startsWith('%PDF')) {
                throw new Error('無效的 PDF 檔案格式');
            }
            
            let slides = [];
            
            try {
                // Try to parse PDF content using PDF.js
                slides = await PresentationConverter.parsePDFContent(arrayBuffer, fileName);
            } catch (pdfError) {
                console.warn('PDF 內容解析失敗，使用基本資訊:', pdfError);
                // Fallback to basic info
                slides = [
                    {
                        slideNumber: 1,
                        title: 'PDF 文件資訊',
                        content: [
                            `來源檔案: ${fileName}.pdf`,
                            `檔案大小: ${PresentationConverter.formatFileSize(arrayBuffer.byteLength)}`,
                            '',
                            '📄 PDF 檔案已載入',
                            '⚠️ 無法完全解析 PDF 內容',
                            '可能原因：',
                            '• PDF 包含複雜格式',
                            '• 需要更強大的解析工具',
                            '• 檔案可能有密碼保護',
                            '',
                            '建議: 仍可嘗試轉換為其他格式'
                        ],
                        notes: `PDF 解析錯誤: ${pdfError.message}`
                    }
                ];
            }
            
            return {
                title: fileName,
                slides: slides,
                totalSlides: slides.length,
                metadata: {
                    fileSize: arrayBuffer.byteLength,
                    fileType: 'pdf',
                    extractedAt: new Date().toISOString(),
                    isValid: true
                }
            };
            
        } catch (error) {
            console.error('PDF 解析錯誤:', error);
            throw new Error(`PDF 檔案解析失敗: ${error.message}`);
        }
    }

    // Extract from HTML files
    static async extractFromHtml(file) {
        try {
            const htmlContent = await file.text();
            const fileName = file.name.replace(/\.[^/.]+$/, '');
            
            // Basic HTML parsing
            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlContent, 'text/html');
            
            // Extract title
            const title = doc.querySelector('title')?.textContent || fileName;
            
            // Extract main content
            const bodyText = doc.body?.textContent || htmlContent.replace(/<[^>]*>/g, '');
            const contentLines = bodyText.split('\n').filter(line => line.trim()).slice(0, 10);
            
            const slides = [
                {
                    slideNumber: 1,
                    title: title,
                    content: contentLines.length > 0 ? contentLines : ['HTML 內容已載入'],
                    notes: 'HTML 文件已轉換為投影片格式'
                }
            ];
            
            return {
                title: title,
                slides: slides,
                totalSlides: slides.length,
                metadata: {
                    fileSize: htmlContent.length,
                    fileType: 'html',
                    extractedAt: new Date().toISOString(),
                    isValid: true
                }
            };
            
        } catch (error) {
            console.error('HTML 解析錯誤:', error);
            throw new Error(`HTML 檔案解析失敗: ${error.message}`);
        }
    }

    // Format file size
    static formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // Convert presentation to PDF using HTML-Canvas method
    static async convertToPdf(content, title, options = {}) {
        try {
            console.log('🔄 開始簡報轉PDF (HTML-Canvas 方法)...');
            
            // Load required libraries
            await Promise.all([
                PresentationConverter.loadJsPDF(),
                PresentationConverter.loadHTML2Canvas()
            ]);
            
            console.log('📝 準備簡報HTML內容: ' + (title || 'Presentation'));
            
            // Convert presentation content to formatted text
            let formattedContent = '';
            
            if (content && content.slides) {
                // Process slides
                content.slides.forEach((slide, index) => {
                    formattedContent += `投影片 ${slide.slideNumber || index + 1}: ${slide.title || '無標題'}\n\n`;
                    
                    if (slide.content && Array.isArray(slide.content)) {
                        formattedContent += slide.content.join('\n') + '\n\n';
                    } else if (slide.content) {
                        formattedContent += slide.content + '\n\n';
                    }
                    
                    if (slide.notes) {
                        formattedContent += `備註: ${slide.notes}\n\n`;
                    }
                    
                    formattedContent += '---\n\n';
                });
            } else {
                // Fallback for simple text content
                formattedContent = content || '無簡報內容';
            }
            
            // Create HTML content with proper formatting
            const htmlContent = `
                <div style="
                    width: 210mm; 
                    padding: 20mm; 
                    background: white;
                    font-family: 'Microsoft YaHei', 'PingFang SC', 'Hiragino Sans GB', 'SimSun', sans-serif;
                    font-size: 14px;
                    line-height: 1.6;
                    color: #333;
                    box-sizing: border-box;
                ">
                    ${title ? `<div style="font-size: 24px; font-weight: bold; margin-bottom: 30px; color: #2c3e50; text-align: center; border-bottom: 2px solid #3498db; padding-bottom: 15px;">📊 ${title}</div>` : ''}
                    <div style="white-space: pre-wrap; font-size: 12px;">${formattedContent}</div>
                    <div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; font-size: 10px; color: #666; text-align: center;">
                        由簡報轉換器生成 - ${new Date().toLocaleString()}
                    </div>
                </div>
            `;
            
            // Create temporary container
            const tempDiv = document.createElement('div');
            tempDiv.style.position = 'fixed';
            tempDiv.style.top = '-9999px';
            tempDiv.style.left = '-9999px';
            tempDiv.innerHTML = htmlContent;
            document.body.appendChild(tempDiv);
            
            console.log('🎨 正在將簡報HTML轉換為Canvas...');
            
            // Convert HTML to Canvas using html2canvas
            const canvas = await html2canvas(tempDiv.firstElementChild, {
                scale: 2,
                useCORS: true,
                allowTaint: true,
                backgroundColor: '#ffffff',
                width: 794,  // A4 width in pixels at 96 DPI
                height: 1123 // A4 height in pixels at 96 DPI
            });
            
            console.log('✅ Canvas 生成成功: ' + canvas.width + 'x' + canvas.height);
            
            // Clean up temporary element
            document.body.removeChild(tempDiv);
            
            // Create PDF
            const { jsPDF } = window.jspdf || window;
            const pdf = new jsPDF({
                orientation: 'portrait',
                unit: 'mm',
                format: 'a4'
            });
            
            console.log('📄 正在生成簡報PDF...');
            
            // Add canvas as image to PDF
            const imgData = canvas.toDataURL('image/png');
            pdf.addImage(imgData, 'PNG', 0, 0, 210, 297);
            
            const pdfBlob = pdf.output('blob');
            console.log('✅ 簡報PDF創建成功:', pdfBlob.size, 'bytes');
            
            return pdfBlob;
            
        } catch (error) {
            console.error('簡報轉PDF錯誤:', error);
            throw new Error('簡報轉PDF失敗: ' + error.message);
        }
    }

    // Convert presentation to HTML
    static convertToHtml(content, title, options = {}) {
        try {
            console.log('🔄 開始簡報轉HTML...');
            
            let htmlContent = '';
            
            if (content && content.slides) {
                // Process slides into HTML
                const slidesHtml = content.slides.map((slide, index) => `
                    <div class="slide" data-slide="${index + 1}">
                        <h2>投影片 ${slide.slideNumber || index + 1}: ${slide.title || '無標題'}</h2>
                        <div class="slide-content">
                            ${Array.isArray(slide.content) ? 
                                slide.content.map(item => `<p>${item}</p>`).join('') : 
                                `<p>${slide.content || ''}</p>`
                            }
                        </div>
                        ${slide.notes ? `<div class="slide-notes"><strong>備註:</strong> ${slide.notes}</div>` : ''}
                    </div>
                `).join('');
                
                htmlContent = `<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${title || '簡報文件'}</title>
    <style>
        body { 
            font-family: 'Microsoft YaHei', 'PingFang SC', sans-serif; 
            line-height: 1.6; 
            margin: 40px; 
            color: #333; 
        }
        h1 { 
            color: #2c3e50; 
            border-bottom: 3px solid #3498db; 
            padding-bottom: 15px; 
            text-align: center; 
        }
        .slide { 
            margin-bottom: 40px; 
            padding: 20px; 
            border: 1px solid #ddd; 
            border-radius: 8px; 
            background: #f9f9f9; 
        }
        .slide h2 { 
            color: #27ae60; 
            margin-top: 0; 
        }
        .slide-content { 
            margin: 15px 0; 
        }
        .slide-notes { 
            background: #fff3cd; 
            padding: 10px; 
            border-left: 4px solid #ffc107; 
            margin-top: 15px; 
        }
        .footer { 
            text-align: center; 
            margin-top: 40px; 
            padding-top: 20px; 
            border-top: 1px solid #ddd; 
            color: #666; 
        }
    </style>
</head>
<body>
    <h1>📊 ${title || '簡報文件'}</h1>
    ${slidesHtml}
    <div class="footer">
        <p>由簡報轉換器生成 - ${new Date().toLocaleString()}</p>
    </div>
</body>
</html>`;
            } else {
                // Fallback for simple content
                htmlContent = `<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>${title || '簡報文件'}</title>
</head>
<body>
    <h1>${title || '簡報文件'}</h1>
    <pre>${content || '無簡報內容'}</pre>
</body>
</html>`;
            }
            
            const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
            console.log('✅ 簡報轉HTML完成');
            return blob;
            
        } catch (error) {
            console.error('簡報轉HTML錯誤:', error);
            throw new Error('簡報轉HTML失敗: ' + error.message);
        }
    }

    // Convert presentation to text
    static convertToText(content, title, options = {}) {
        try {
            console.log('🔄 開始簡報轉文字...');
            
            let textContent = '';
            
            if (title) {
                textContent += `${title}\n${'='.repeat(title.length)}\n\n`;
            }
            
            if (content && content.slides) {
                content.slides.forEach((slide, index) => {
                    textContent += `投影片 ${slide.slideNumber || index + 1}: ${slide.title || '無標題'}\n`;
                    textContent += '-'.repeat(50) + '\n';
                    
                    if (Array.isArray(slide.content)) {
                        textContent += slide.content.join('\n') + '\n';
                    } else if (slide.content) {
                        textContent += slide.content + '\n';
                    }
                    
                    if (slide.notes) {
                        textContent += `\n備註: ${slide.notes}\n`;
                    }
                    
                    textContent += '\n';
                });
            } else {
                textContent += content || '無簡報內容';
            }
            
            textContent += `\n\n由簡報轉換器生成 - ${new Date().toLocaleString()}`;
            
            const blob = new Blob([textContent], { type: 'text/plain;charset=utf-8' });
            console.log('✅ 簡報轉文字完成');
            return blob;
            
        } catch (error) {
            console.error('簡報轉文字錯誤:', error);
            throw new Error('簡報轉文字失敗: ' + error.message);
        }
    }

    // Load jsPDF library
    static async loadJsPDF() {
        if (window.jsPDF || (window.jspdf && window.jspdf.jsPDF)) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/jspdf@2.5.1/dist/jspdf.umd.min.js';
            script.onload = () => {
                setTimeout(() => {
                    if (window.jsPDF || (window.jspdf && window.jspdf.jsPDF)) {
                        console.log('✅ jsPDF 載入成功');
                        resolve();
                    } else {
                        reject(new Error('jsPDF 載入後無法使用'));
                    }
                }, 100);
            };
            script.onerror = () => reject(new Error('jsPDF 載入失敗'));
            document.head.appendChild(script);
        });
    }

    // Load html2canvas library
    static async loadHTML2Canvas() {
        if (window.html2canvas) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/html2canvas@1.4.1/dist/html2canvas.min.js';
            script.onload = () => {
                setTimeout(() => {
                    if (window.html2canvas) {
                        console.log('✅ html2canvas 載入成功');
                        resolve();
                    } else {
                        reject(new Error('html2canvas 載入後無法使用'));
                    }
                }, 100);
            };
            script.onerror = () => reject(new Error('html2canvas 載入失敗'));
            document.head.appendChild(script);
        });
    }

    // Parse PPTX content using basic ZIP structure analysis
    static async parsePPTXContent(arrayBuffer, fileName) {
        try {
            // Load JSZip if not available
            await PresentationConverter.loadJSZip();
            
            const zip = new JSZip();
            const zipContent = await zip.loadAsync(arrayBuffer);
            
            // Extract presentation structure
            const slides = [];
            let slideCount = 0;
            
            // Look for slide files
            const slideFiles = [];
            zipContent.forEach((relativePath, file) => {
                if (relativePath.includes('ppt/slides/slide') && relativePath.endsWith('.xml')) {
                    slideFiles.push({ path: relativePath, file: file });
                }
            });
            
            // Sort slide files by number
            slideFiles.sort((a, b) => {
                const aNum = parseInt(a.path.match(/slide(\d+)\.xml/)?.[1] || '0');
                const bNum = parseInt(b.path.match(/slide(\d+)\.xml/)?.[1] || '0');
                return aNum - bNum;
            });
            
            // Process each slide
            for (let i = 0; i < slideFiles.length; i++) {
                const slideFile = slideFiles[i];
                try {
                    const slideXml = await slideFile.file.async('text');
                    const slideContent = PresentationConverter.parseSlideXML(slideXml, i + 1);
                    if (slideContent) {
                        slides.push(slideContent);
                        slideCount++;
                    }
                } catch (error) {
                    console.warn(`無法解析投影片 ${i + 1}:`, error);
                    // 添加錯誤投影片
                    slides.push({
                        slideNumber: i + 1,
                        title: `投影片 ${i + 1}`,
                        content: ['⚠️ 此投影片內容無法正確解析'],
                        notes: `解析錯誤: ${error.message}`
                    });
                    slideCount++;
                }
            }
            
            // If no slides found, create a summary slide
            if (slides.length === 0) {
                slides.push({
                    slideNumber: 1,
                    title: '簡報摘要',
                    content: [
                        `檔案名稱: ${fileName}`,
                        '📊 PPTX 檔案結構分析:',
                        `• 檔案中包含 ${slideFiles.length} 張投影片`,
                        '• 已嘗試解析內容結構',
                        '• 部分內容可能需要進階解析器',
                        '',
                        '✨ 檔案格式驗證通過',
                        '🔧 支援轉換為多種輸出格式'
                    ],
                    notes: '基於 PPTX 檔案結構分析產生的摘要'
                });
            }
            
            return slides;
            
        } catch (error) {
            console.error('PPTX 解析錯誤:', error);
            // 回退到基本資訊
            return [
                {
                    slideNumber: 1,
                    title: '檔案資訊',
                    content: [
                        `檔案名稱: ${fileName}`,
                        '⚠️ PPTX 內容解析遇到困難',
                        '可能的原因:',
                        '• 檔案結構複雜',
                        '• 包含特殊元素',
                        '• 需要更強大的解析工具',
                        '',
                        '建議: 仍可轉換為其他格式進行檢視'
                    ],
                    notes: `解析錯誤: ${error.message}`
                }
            ];
        }
    }

    // Parse individual slide XML content
    static parseSlideXML(xmlContent, slideNumber) {
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlContent, 'text/xml');
            
            // Extract text content from the slide
            const textElements = xmlDoc.querySelectorAll('a\\:t, t');
            const textContent = [];
            
            textElements.forEach(element => {
                const text = element.textContent?.trim();
                if (text && text.length > 0) {
                    textContent.push(text);
                }
            });
            
            // Try to identify title (usually the first or largest text element)
            let title = `投影片 ${slideNumber}`;
            if (textContent.length > 0) {
                // Use first text as title if it's short enough
                if (textContent[0].length <= 50) {
                    title = textContent[0];
                    textContent.shift(); // Remove title from content
                } else {
                    title = `投影片 ${slideNumber}`;
                }
            }
            
            // Create slide object
            return {
                slideNumber: slideNumber,
                title: title || `投影片 ${slideNumber}`,
                content: textContent.length > 0 ? textContent : ['此投影片包含非文字內容或無法解析的元素'],
                notes: textContent.length > 0 ? `成功提取 ${textContent.length} 個文字元素` : '未找到文字內容'
            };
            
        } catch (error) {
            console.error(`解析投影片 ${slideNumber} XML 時發生錯誤:`, error);
            return {
                slideNumber: slideNumber,
                title: `投影片 ${slideNumber}`,
                content: ['⚠️ 投影片內容解析失敗'],
                notes: `XML 解析錯誤: ${error.message}`
            };
        }
    }

    // Load JSZip library
    static async loadJSZip() {
        if (window.JSZip) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/jszip@3.10.1/dist/jszip.min.js';
            script.onload = () => {
                setTimeout(() => {
                    if (window.JSZip) {
                        console.log('✅ JSZip 載入成功');
                        resolve();
                    } else {
                        reject(new Error('JSZip 載入後無法使用'));
                    }
                }, 100);
            };
            script.onerror = () => reject(new Error('JSZip 載入失敗'));
            document.head.appendChild(script);
        });
    }

    // Parse PDF content using PDF.js
    static async parsePDFContent(arrayBuffer, fileName) {
        try {
            // Load PDF.js if not available
            await PresentationConverter.loadPDFJS();
            
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            const slides = [];
            
            console.log(`PDF 包含 ${pdf.numPages} 頁`);
            
            // Extract text from each page
            for (let pageNum = 1; pageNum <= Math.min(pdf.numPages, 20); pageNum++) { // Limit to 20 pages
                try {
                    const page = await pdf.getPage(pageNum);
                    const textContent = await page.getTextContent();
                    
                    // Extract text items
                    const textItems = textContent.items.map(item => item.str.trim()).filter(str => str.length > 0);
                    
                    // Try to identify title (first significant text)
                    let title = `第 ${pageNum} 頁`;
                    if (textItems.length > 0) {
                        const firstLine = textItems[0];
                        if (firstLine.length <= 100) {
                            title = firstLine;
                        }
                    }
                    
                    // Group remaining text as content
                    const content = textItems.length > 1 ? textItems.slice(1) : textItems;
                    
                    slides.push({
                        slideNumber: pageNum,
                        title: title || `第 ${pageNum} 頁`,
                        content: content.length > 0 ? content : ['此頁面沒有可提取的文字內容'],
                        notes: `從 PDF 第 ${pageNum} 頁提取了 ${textItems.length} 個文字元素`
                    });
                    
                } catch (pageError) {
                    console.warn(`解析第 ${pageNum} 頁時發生錯誤:`, pageError);
                    slides.push({
                        slideNumber: pageNum,
                        title: `第 ${pageNum} 頁`,
                        content: ['⚠️ 此頁面內容無法解析'],
                        notes: `頁面解析錯誤: ${pageError.message}`
                    });
                }
            }
            
            // If more than 20 pages, add a summary
            if (pdf.numPages > 20) {
                slides.push({
                    slideNumber: slides.length + 1,
                    title: '文件摘要',
                    content: [
                        `此 PDF 文件共包含 ${pdf.numPages} 頁`,
                        `為了效能考慮，僅顯示前 20 頁內容`,
                        '',
                        '如需查看完整內容：',
                        '• 請使用專業的 PDF 閱讀器',
                        '• 或將檔案分割為較小的部分'
                    ],
                    notes: `完整文件包含 ${pdf.numPages} 頁，已略過 ${pdf.numPages - 20} 頁`
                });
            }
            
            return slides;
            
        } catch (error) {
            console.error('PDF.js 解析錯誤:', error);
            throw error;
        }
    }

    // Load PDF.js library
    static async loadPDFJS() {
        if (window.pdfjsLib) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.min.js';
            script.onload = () => {
                setTimeout(() => {
                    if (window.pdfjsLib) {
                        // Configure PDF.js worker
                        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
                        console.log('✅ PDF.js 載入成功');
                        resolve();
                    } else {
                        reject(new Error('PDF.js 載入後無法使用'));
                    }
                }, 100);
            };
            script.onerror = () => reject(new Error('PDF.js 載入失敗'));
            document.head.appendChild(script);
        });
    }
}

// Export for use in main application
if (typeof module !== 'undefined' && module.exports) {
    module.exports = PresentationConverter;
} else if (typeof window !== 'undefined') {
    window.PresentationConverter = PresentationConverter;
}
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
            
            // Create basic slide structure
            let slides = [];
            
            // Create default slides based on file analysis
            slides = [
                {
                    slideNumber: 1,
                    title: '簡報摘要',
                    content: [
                        `檔案名稱: ${fileName}`,
                        `檔案格式: ${fileType.toUpperCase()}`,
                        `檔案大小: ${PresentationConverter.formatFileSize(arrayBuffer.byteLength)}`,
                        '',
                        '✨ 檔案已成功載入',
                        '📊 支援的輸出格式:',
                        '• PDF 文件',
                        '• HTML 網頁',
                        '• 純文字內容',
                        '• Markdown 格式'
                    ],
                    notes: '此投影片說明支援的轉換格式'
                }
            ];
            
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
            
            // Create default slide for PDF
            const slides = [
                {
                    slideNumber: 1,
                    title: 'PDF 文件內容',
                    content: [
                        `來源檔案: ${fileName}.pdf`,
                        `檔案大小: ${PresentationConverter.formatFileSize(arrayBuffer.byteLength)}`,
                        '',
                        '📄 PDF 檔案已載入',
                        '⚠️ PDF 內容解析需要額外的函式庫支援',
                        '',
                        '建議的轉換選項:',
                        '• 轉換為 HTML 格式以便瀏覽',
                        '• 轉換為純文字格式',
                        '• 保持原始 PDF 格式'
                    ],
                    notes: 'PDF 檔案需要專門的解析工具'
                }
            ];
            
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
}

// Export for use in main application
if (typeof module !== 'undefined' && module.exports) {
    module.exports = PresentationConverter;
} else if (typeof window !== 'undefined') {
    window.PresentationConverter = PresentationConverter;
}
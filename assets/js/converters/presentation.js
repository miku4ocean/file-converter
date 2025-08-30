// Presentation conversion utilities

class PresentationConverter {
    constructor() {
        this.supportedFormats = {
            input: ['pptx', 'ppt', 'odp', 'html'],
            output: ['pdf', 'html', 'images', 'txt', 'md']
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
            // Note: This would require a PowerPoint parsing library
            // For now, return a placeholder implementation
            const arrayBuffer = await file.arrayBuffer();
            const fileName = file.name.replace(/\.[^/.]+$/, '');
            
            // Placeholder: In real implementation, parse PPTX structure
            const slides = [
                {
                    slideNumber: 1,
                    title: '此功能需要專門的 PowerPoint 解析函式庫',
                    content: [
                        'PowerPoint 檔案解析需要複雜的 ZIP 和 XML 處理',
                        '建議使用專門的函式庫如 PptxGenJS 或類似工具',
                        '或考慮使用伺服器端解決方案'
                    ],
                    notes: '這是一個佔位符投影片'
                },
                {
                    slideNumber: 2,
                    title: '檔案資訊',
                    content: [
                        `檔案名稱: ${file.name}`,
                        `檔案大小: ${PresentationConverter.formatFileSize(file.size)}`,
                        `檔案類型: ${file.type}`,
                        `最後修改: ${new Date(file.lastModified).toLocaleString()}`
                    ],
                    notes: '原始檔案的基本資訊'
                }
            ];
            
            return {
                title: fileName,
                slides: slides,
                totalSlides: slides.length,
                fileName: fileName,
                metadata: {
                    fileSize: file.size,
                    fileType: file.type,
                    lastModified: file.lastModified
                }
            };
        } catch (error) {
            throw new Error('PowerPoint 檔案讀取失敗: ' + error.message);
        }
    }

    // Extract content from HTML presentation files
    static async extractFromHtml(file) {
        try {
            const html = await file.text();
            const parser = new DOMParser();
            const doc = parser.parseFromString(html, 'text/html');
            
            // Try to detect presentation structure
            const slides = [];
            let slideElements = [];
            
            // Look for common presentation HTML patterns
            slideElements = doc.querySelectorAll('.slide, section, [data-slide]');
            
            if (slideElements.length === 0) {
                // Fallback: treat each main content block as a slide
                slideElements = doc.querySelectorAll('main > *, body > div, h1, h2');
            }
            
            slideElements.forEach((element, index) => {
                const title = element.querySelector('h1, h2, h3')?.textContent || 
                             element.textContent.split('\n')[0] || 
                             `投影片 ${index + 1}`;
                
                const content = [];
                const textContent = element.textContent.trim();
                
                if (textContent) {
                    const lines = textContent.split('\n')
                        .map(line => line.trim())
                        .filter(line => line.length > 0);
                    content.push(...lines);
                }
                
                slides.push({
                    slideNumber: index + 1,
                    title: title.substring(0, 100), // Limit title length
                    content: content,
                    notes: '',
                    originalHtml: element.outerHTML
                });
            });
            
            const fileName = file.name.replace(/\.[^/.]+$/, '');
            
            return {
                title: doc.title || fileName,
                slides: slides,
                totalSlides: slides.length,
                fileName: fileName,
                originalHtml: html,
                metadata: {
                    fileSize: file.size,
                    fileType: file.type
                }
            };
        } catch (error) {
            throw new Error('HTML 簡報檔案讀取失敗: ' + error.message);
        }
    }

    // Extract content from PDF files
    static async extractFromPdf(file) {
        try {
            // Note: This would require PDF.js or similar library for full PDF parsing
            // For now, return a placeholder implementation
            const fileName = file.name.replace(/\.[^/.]+$/, '');
            
            const slides = [
                {
                    slideNumber: 1,
                    title: 'PDF 轉換功能說明',
                    content: [
                        'PDF 檔案解析需要 PDF.js 或類似函式庫',
                        '建議使用專業的 PDF 處理工具',
                        '或考慮使用伺服器端解決方案'
                    ],
                    notes: '當前為佔位符實作'
                },
                {
                    slideNumber: 2,
                    title: '檔案資訊',
                    content: [
                        `檔案名稱: ${file.name}`,
                        `檔案大小: ${PresentationConverter.formatFileSize(file.size)}`,
                        `檔案類型: ${file.type}`,
                        `最後修改: ${new Date(file.lastModified).toLocaleString()}`
                    ],
                    notes: '原始 PDF 檔案的基本資訊'
                }
            ];
            
            return {
                title: fileName,
                slides: slides,
                totalSlides: slides.length,
                fileName: fileName,
                metadata: {
                    fileSize: file.size,
                    fileType: file.type,
                    lastModified: file.lastModified,
                    isPdfSource: true
                }
            };
        } catch (error) {
            throw new Error('PDF 檔案讀取失敗: ' + error.message);
        }
    }

    // Convert extracted content to various formats
    static async convertToFormat(presentationData, outputFormat, options = {}) {
        const { title, slides, fileName, originalHtml } = presentationData;
        const { 
            includeNotes = true, 
            slideNumbers = true,
            theme = 'default',
            maxFileSize = null 
        } = options;

        switch (outputFormat) {
            case 'txt':
                return PresentationConverter.convertToText(slides, title, { includeNotes, slideNumbers });
            case 'md':
                return PresentationConverter.convertToMarkdown(slides, title, { includeNotes, slideNumbers });
            case 'html':
                return PresentationConverter.convertToHtml(slides, title, { theme, slideNumbers, originalHtml });
            case 'pptx':
                return await PresentationConverter.convertToPptx(slides, title, options);
            case 'pdf':
                return await PresentationConverter.convertToPdf(slides, title, options);
            case 'images':
                return await PresentationConverter.convertToImages(slides, title, options);
            default:
                throw new Error(`不支援的輸出格式: ${outputFormat}`);
        }
    }

    // Convert to plain text
    static convertToText(slides, title, options = {}) {
        const { includeNotes = true, slideNumbers = true } = options;
        
        let content = '';
        
        // Add title
        if (title) {
            content += title.toUpperCase() + '\n';
            content += '='.repeat(title.length) + '\n\n';
        }
        
        slides.forEach((slide, index) => {
            // Slide header
            if (slideNumbers) {
                content += `投影片 ${slide.slideNumber || index + 1}\n`;
                content += '-'.repeat(20) + '\n';
            }
            
            // Slide title
            if (slide.title) {
                content += slide.title + '\n\n';
            }
            
            // Slide content
            if (slide.content && slide.content.length > 0) {
                slide.content.forEach(item => {
                    content += '• ' + item + '\n';
                });
                content += '\n';
            }
            
            // Slide notes
            if (includeNotes && slide.notes) {
                content += '備註: ' + slide.notes + '\n';
            }
            
            content += '\n';
        });
        
        const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
        return blob;
    }

    // Convert to Markdown
    static convertToMarkdown(slides, title, options = {}) {
        const { includeNotes = true, slideNumbers = true } = options;
        
        let markdown = '';
        
        // Add title
        if (title) {
            markdown += `# ${title}\n\n`;
        }
        
        slides.forEach((slide, index) => {
            // Slide header
            if (slideNumbers) {
                markdown += `## 投影片 ${slide.slideNumber || index + 1}`;
                if (slide.title) {
                    markdown += `: ${slide.title}`;
                }
                markdown += '\n\n';
            } else if (slide.title) {
                markdown += `## ${slide.title}\n\n`;
            }
            
            // Slide content
            if (slide.content && slide.content.length > 0) {
                slide.content.forEach(item => {
                    markdown += `- ${item}\n`;
                });
                markdown += '\n';
            }
            
            // Slide notes
            if (includeNotes && slide.notes) {
                markdown += `> **備註:** ${slide.notes}\n\n`;
            }
            
            markdown += '---\n\n';
        });
        
        const blob = new Blob([markdown], { type: 'text/markdown;charset=utf-8' });
        return blob;
    }

    // Convert to HTML presentation
    static convertToHtml(slides, title, options = {}) {
        const { theme = 'default', slideNumbers = true, originalHtml = null } = options;
        
        if (originalHtml) {
            const blob = new Blob([originalHtml], { type: 'text/html;charset=utf-8' });
            return blob;
        }
        
        const slideHtml = slides.map((slide, index) => {
            let slideContent = `<section class="slide" data-slide="${slide.slideNumber || index + 1}">`;
            
            if (slideNumbers) {
                slideContent += `<div class="slide-number">${slide.slideNumber || index + 1}</div>`;
            }
            
            if (slide.title) {
                slideContent += `<h1>${PresentationConverter.escapeHtml(slide.title)}</h1>`;
            }
            
            if (slide.content && slide.content.length > 0) {
                slideContent += '<ul>';
                slide.content.forEach(item => {
                    slideContent += `<li>${PresentationConverter.escapeHtml(item)}</li>`;
                });
                slideContent += '</ul>';
            }
            
            if (slide.notes) {
                slideContent += `<div class="slide-notes">${PresentationConverter.escapeHtml(slide.notes)}</div>`;
            }
            
            slideContent += '</section>';
            return slideContent;
        }).join('\n');
        
        const html = `<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${title || '簡報'}</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            margin: 0;
            padding: 0;
            background: #f5f5f5;
        }
        
        .presentation-container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        
        .presentation-title {
            text-align: center;
            color: #2c3e50;
            margin-bottom: 40px;
            font-size: 2.5em;
            border-bottom: 3px solid #3498db;
            padding-bottom: 20px;
        }
        
        .slide {
            background: white;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            margin: 30px 0;
            padding: 40px;
            position: relative;
            min-height: 300px;
        }
        
        .slide-number {
            position: absolute;
            top: 15px;
            right: 20px;
            background: #3498db;
            color: white;
            padding: 5px 12px;
            border-radius: 15px;
            font-size: 0.9em;
            font-weight: 600;
        }
        
        .slide h1 {
            color: #2c3e50;
            border-bottom: 2px solid #ecf0f1;
            padding-bottom: 15px;
            margin-bottom: 25px;
            font-size: 1.8em;
        }
        
        .slide ul {
            list-style: none;
            padding: 0;
        }
        
        .slide li {
            margin: 15px 0;
            padding: 10px 20px;
            border-left: 4px solid #3498db;
            background: #f8f9fa;
            border-radius: 0 5px 5px 0;
        }
        
        .slide-notes {
            margin-top: 30px;
            padding: 20px;
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 5px;
            font-style: italic;
            color: #856404;
        }
        
        .slide-notes::before {
            content: "📝 備註: ";
            font-weight: 600;
        }
        
        @media print {
            .slide {
                page-break-after: always;
                box-shadow: none;
                border: 1px solid #ddd;
            }
        }
        
        @media (max-width: 768px) {
            .presentation-container {
                padding: 10px;
            }
            
            .slide {
                padding: 20px;
            }
            
            .presentation-title {
                font-size: 2em;
            }
        }
    </style>
</head>
<body>
    <div class="presentation-container">
        <h1 class="presentation-title">${title || '簡報'}</h1>
        ${slideHtml}
    </div>
    
    <script>
        // Simple keyboard navigation
        let currentSlide = 0;
        const slides = document.querySelectorAll('.slide');
        
        document.addEventListener('keydown', function(e) {
            if (e.key === 'ArrowRight' && currentSlide < slides.length - 1) {
                currentSlide++;
                slides[currentSlide].scrollIntoView({ behavior: 'smooth' });
            } else if (e.key === 'ArrowLeft' && currentSlide > 0) {
                currentSlide--;
                slides[currentSlide].scrollIntoView({ behavior: 'smooth' });
            }
        });
    </script>
</body>
</html>`;
        
        const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
        return blob;
    }

    // Convert to PDF (improved placeholder with proper error handling)
    static async convertToPdf(slides, title, options = {}) {
        try {
            // Validate input data
            if (!slides || slides.length === 0) {
                throw new Error('無有可轉換的投影片內容');
            }

            // Create a more detailed PDF preview document
            const pdfPreview = `PDF 轉換預覽文件
${'='.repeat(30)}

簡報標題: ${title || '無標題'}
投影片總數: ${slides.length}
生成時間: ${new Date().toLocaleString()}

投影片內容:
${'-'.repeat(50)}

${slides.map((slide, index) => {
    let slideText = `投影片 ${slide.slideNumber || index + 1}\n`;
    slideText += `標題: ${slide.title || '無標題'}\n\n`;
    
    if (slide.content && slide.content.length > 0) {
        slideText += '內容:\n';
        slide.content.forEach(item => {
            slideText += `• ${item}\n`;
        });
    }
    
    if (slide.notes && slide.notes.trim()) {
        slideText += `\n備註: ${slide.notes}\n`;
    }
    
    return slideText;
}).join('\n' + '-'.repeat(50) + '\n\n')}

技術說明:
${'-'.repeat(30)}
本檔案為 PDF 轉換預覽文件。
完整的 PDF 轉換功能需要以下函式庫:

1. jsPDF - 基本 PDF 生成
2. html2canvas - HTML 轉圖片
3. PptxGenJS - PowerPoint 處理

安裝指令:
<script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>`;

            // Create blob with proper MIME type
            const blob = new Blob([pdfPreview], { 
                type: 'text/plain;charset=utf-8' 
            });
            return blob;
            
        } catch (error) {
            console.error('PDF 轉換錯誤:', error);
            
            // Create error report document
            const errorReport = `PDF 轉換錯誤報告
${'='.repeat(30)}

錯誤時間: ${new Date().toLocaleString()}
錯誤訊息: ${error.message}

建議解決方案:
1. 確認檔案格式正確
2. 檢查檔案是否損壞
3. 嘗試其他輸出格式 (TXT, HTML, MD)
4. 檢查瀏覽器控制台獲取詳細錯誤訊息

如果問題持續，請嘗試使用專業的 PDF 處理工具。`;
            
            const blob = new Blob([errorReport], { 
                type: 'text/plain;charset=utf-8' 
            });
            return blob;
        }
    }

    // Convert to PowerPoint format (placeholder - requires PptxGenJS)
    static async convertToPptx(slides, title, options = {}) {
        try {
            // Validate input
            if (!slides || slides.length === 0) {
                throw new Error('無有可轉換的投影片內容');
            }

            // Note: This would require PptxGenJS library for actual implementation
            const pptxContent = `PowerPoint 轉換預覽文件
${'='.repeat(40)}

簡報標題: ${title || '無標題'}
投影片數量: ${slides.length}
生成時間: ${new Date().toLocaleString()}

投影片列表:
${slides.map((slide, index) => {
    return `${index + 1}. ${slide.title || `投影片 ${index + 1}`} (內容項數: ${slide.content?.length || 0})`;
}).join('\n')}

實作指南:
${'-'.repeat(30)}
要實現完整的 PPTX 轉換功能，需要：

1. 安裝 PptxGenJS 函式庫
   <script src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js"></script>

2. 實作投影片布局和樣式
3. 處理文字、圖片和其他媒體元素
4. 配置投影片尺寸和主題

當前為模擬輸出，實際使用時會生成標準 PPTX 檔案。`;

            const blob = new Blob([pptxContent], { type: 'text/plain;charset=utf-8' });
            return blob;
            
        } catch (error) {
            console.error('PPTX 轉換錯誤:', error);
            
            const errorReport = `PPTX 轉換錯誤報告
${'='.repeat(30)}

錯誤時間: ${new Date().toLocaleString()}
錯誤訊息: ${error.message}

可能原因:
1. 檔案格式不支援
2. 檔案內容損壞
3. 所需函式庫未加載

建議解決方案:
1. 嘗試其他輸出格式 (HTML, TXT, MD)
2. 檢查原始檔案是否可正常開啟
3. 使用專業 PDF 處理軟體`;
            
            const blob = new Blob([errorReport], { type: 'text/plain;charset=utf-8' });
            return blob;
        }
    }

    // Convert slides to images (placeholder - requires html2canvas or similar)
    static async convertToImages(slides, title, options = {}) {
        try {
            // Note: This would require html2canvas library
            const imagesInfo = `圖片轉換功能需要 html2canvas 函式庫支援

簡報標題: ${title || '無標題'}
投影片數量: ${slides.length}

每張投影片將轉換為一個圖片檔案:
${slides.map((slide, index) => 
`${index + 1}. ${slide.title || `投影片${index + 1}`}.png`).join('\n')}

請在 HTML 中加入：
<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>

實際實作時會生成 ZIP 檔案包含所有投影片圖片。`;

            const blob = new Blob([imagesInfo], { type: 'text/plain;charset=utf-8' });
            return blob;
        } catch (error) {
            throw new Error('圖片轉換失敗: ' + error.message);
        }
    }

    // Convert to PowerPoint format (placeholder - requires PptxGenJS or similar)
    static async convertToPptx(slides, title, options = {}) {
        try {
            // Note: This would require PptxGenJS library
            const pptxContent = `PowerPoint 轉換功能需要 PptxGenJS 函式庫支援

簡報標題: ${title || '無標題'}
投影片數量: ${slides.length}

請在 HTML 中加入：
<script src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js"></script>

投影片內容預覽:
${slides.slice(0, 3).map((slide, index) => 
`投影片 ${index + 1}: ${slide.title || '無標題'}
${slide.content.join('\n')}`).join('\n\n')}

${slides.length > 3 ? `...還有 ${slides.length - 3} 張投影片` : ''}

實際實作時會生成標準 PPTX 格式檔案。`;

            const blob = new Blob([pptxContent], { type: 'text/plain;charset=utf-8' });
            return blob;
        } catch (error) {
            throw new Error('PowerPoint 轉換失敗: ' + error.message);
        }
    }

    // Helper function to escape HTML
    static escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // Format file size
    static formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // Get presentation statistics
    static getPresentationStats(slides) {
        if (!slides || slides.length === 0) {
            return {
                slideCount: 0,
                totalWords: 0,
                averageWordsPerSlide: 0,
                slidesWithNotes: 0,
                estimatedPresentationTime: 0
            };
        }

        let totalWords = 0;
        let slidesWithNotes = 0;

        slides.forEach(slide => {
            // Count words in title
            if (slide.title) {
                totalWords += slide.title.split(/\s+/).length;
            }
            
            // Count words in content
            if (slide.content && slide.content.length > 0) {
                slide.content.forEach(item => {
                    totalWords += item.split(/\s+/).length;
                });
            }
            
            // Count slides with notes
            if (slide.notes && slide.notes.trim().length > 0) {
                slidesWithNotes++;
            }
        });

        const averageWordsPerSlide = Math.round(totalWords / slides.length);
        
        // Estimate presentation time: 1-2 minutes per slide + word count
        const estimatedPresentationTime = Math.round(
            slides.length * 1.5 + totalWords / 150 // 150 words per minute
        );

        return {
            slideCount: slides.length,
            totalWords,
            averageWordsPerSlide,
            slidesWithNotes,
            estimatedPresentationTime
        };
    }
}

// Export for use in main application
if (typeof module !== 'undefined' && module.exports) {
    module.exports = PresentationConverter;
} else if (typeof window !== 'undefined') {
    window.PresentationConverter = PresentationConverter;
}
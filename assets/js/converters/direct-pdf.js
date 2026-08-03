// Direct PDF conversion - Print-like functionality
// Converts files to PDF preserving original appearance like browser print-to-PDF

class DirectPDFConverter {
    constructor() {
        this.supportedFormats = [
            'docx', 'doc', 'pptx', 'ppt', 'html', 'htm', 
            'txt', 'md', 'rtf', 'odt', 'odp', 'pdf'
        ];
    }

    // Main conversion method - like printing to PDF
    static async convertToPDF(file, options = {}) {
        try {
            console.log('🖨️ 開始直接轉PDF (類似列印功能):', file.name);
            
            const fileType = DirectPDFConverter.getFileType(file);
            
            switch (fileType) {
                case 'pdf':
                    // PDF files: return as-is
                    return await DirectPDFConverter.handlePDFFile(file);
                
                case 'html':
                case 'htm':
                    // HTML files: render in iframe and print to PDF
                    return await DirectPDFConverter.convertHTMLToPDF(file);
                
                case 'docx':
                case 'doc':
                case 'odt':
                    // Document files: render using document viewer
                    return await DirectPDFConverter.convertDocumentToPDF(file);
                
                case 'pptx':
                case 'ppt':
                case 'odp':
                    // Presentation files: render slides directly
                    return await DirectPDFConverter.convertPresentationToPDF(file);
                
                case 'txt':
                case 'md':
                case 'rtf':
                    // Text files: render with proper formatting
                    return await DirectPDFConverter.convertTextToPDF(file);
                
                default:
                    throw new Error(`不支援的檔案格式: ${fileType.toUpperCase()}`);
            }
            
        } catch (error) {
            console.error('直接PDF轉換失敗:', error);
            throw new Error(`PDF轉換失敗: ${error.message}`);
        }
    }

    // Handle PDF files (pass-through)
    static async handlePDFFile(file) {
        console.log('📄 PDF檔案直接回傳');
        return new Blob([await file.arrayBuffer()], { type: 'application/pdf' });
    }

    // Convert HTML files to PDF using browser's native print API
    static async convertHTMLToPDF(file) {
        try {
            console.log('🖨️ HTML檔案轉PDF (瀏覽器原生列印API)');
            
            const htmlContent = await file.text();
            
            // 直接使用 Puppeteer-style 轉換，避免觸發瀏覽器列印對話框
            return await DirectPDFConverter.convertHTMLWithPuppeteerStyle(htmlContent);
            
        } catch (error) {
            console.error('HTML轉PDF失敗:', error);
            throw error;
        }
    }

    // DISABLED: Use browser's native print functionality (causes print dialog)
    static async convertHTMLWithNativePrint(htmlContent) {
        console.warn('⚠️ 原生列印API已禁用，改用Puppeteer風格轉換');
        return await DirectPDFConverter.convertHTMLWithPuppeteerStyle(htmlContent);
    }

    // Puppeteer-style HTML to PDF conversion
    static async convertHTMLWithPuppeteerStyle(htmlContent) {
        try {
            console.log('🎭 使用Puppeteer風格轉換');
            
            // Create optimized HTML with proper page breaks
            const optimizedHTML = DirectPDFConverter.prepareHTMLForPDF(htmlContent);
            
            // Create iframe with exact A4 dimensions
            const iframe = document.createElement('iframe');
            iframe.style.position = 'fixed';
            iframe.style.top = '-9999px';
            iframe.style.left = '-9999px';
            iframe.style.width = '794px';   // A4 width at 96dpi
            iframe.style.height = '1123px'; // A4 height at 96dpi
            iframe.style.border = 'none';
            iframe.style.background = 'white';
            document.body.appendChild(iframe);
            
            // Load optimized content
            iframe.contentDocument.open();
            iframe.contentDocument.write(optimizedHTML);
            iframe.contentDocument.close();
            
            // Wait for content to render
            await new Promise(resolve => {
                iframe.onload = resolve;
                setTimeout(resolve, 2000); // Longer wait for complex content
            });
            
            // Calculate total content height for pagination
            const contentHeight = Math.max(
                iframe.contentDocument.body.scrollHeight,
                iframe.contentDocument.body.offsetHeight,
                iframe.contentDocument.documentElement.scrollHeight
            );
            
            const pageHeight = 1123; // A4 height in pixels
            const numPages = Math.max(1, Math.ceil(contentHeight / pageHeight));
            
            console.log(`📄 內容高度: ${contentHeight}px, 需要 ${numPages} 頁`);
            
            // Create PDF with proper pagination
            await DirectPDFConverter.loadJsPDF();
            const { jsPDF } = window.jspdf || window;
            const pdf = new jsPDF({
                orientation: 'portrait',
                unit: 'mm',
                format: 'a4',
                compress: true
            });
            
            // Capture each page separately
            for (let page = 0; page < numPages; page++) {
                const yOffset = page * pageHeight;
                
                console.log(`📄 渲染第 ${page + 1} 頁 (offset: ${yOffset}px)`);
                
                // Scroll to current page position
                iframe.contentWindow.scrollTo(0, yOffset);
                await new Promise(resolve => setTimeout(resolve, 300));
                
                // Capture current page
                await DirectPDFConverter.loadHTML2Canvas();
                const canvas = await html2canvas(iframe.contentDocument.body, {
                    scale: 2, // Balance between quality and performance
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff',
                    width: 794,
                    height: pageHeight,
                    scrollX: 0,
                    scrollY: yOffset,
                    windowWidth: 794,
                    windowHeight: pageHeight
                });
                
                // Add page to PDF
                if (page > 0) {
                    pdf.addPage();
                }
                
                const imgData = canvas.toDataURL('image/jpeg', 0.95);
                pdf.addImage(imgData, 'JPEG', 0, 0, 210, 297);
            }
            
            // Clean up
            document.body.removeChild(iframe);
            
            const pdfBlob = pdf.output('blob');
            console.log(`✅ 多頁HTML轉PDF完成: ${numPages} 頁, ${(pdfBlob.size/1024).toFixed(2)}KB`);
            return pdfBlob;
            
        } catch (error) {
            console.error('Puppeteer風格轉換失敗:', error);
            throw error;
        }
    }

    // Prepare HTML content for optimal PDF conversion
    static prepareHTMLForPDF(htmlContent) {
        return `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=794, initial-scale=1.0">
            <title>PDF Document</title>
            <style>
                @page {
                    size: A4;
                    margin: 20mm;
                }
                
                * {
                    box-sizing: border-box;
                    -webkit-print-color-adjust: exact;
                    color-adjust: exact;
                }
                
                body {
                    font-family: 'Times New Roman', 'PingFang TC', 'Microsoft JhengHei', 'Noto Serif CJK TC', 'SimSun', serif;
                    font-size: 12pt;
                    line-height: 1.5;
                    color: #000;
                    background: white;
                    margin: 20mm;
                    padding: 0;
                    width: 754px; /* A4 width minus margins */
                }
                
                h1, h2, h3, h4, h5, h6 {
                    page-break-after: avoid;
                    margin-top: 1.5em;
                    margin-bottom: 0.5em;
                    color: #2c3e50;
                }
                
                p {
                    margin: 0.5em 0;
                    orphans: 2;
                    widows: 2;
                }
                
                table {
                    width: 100%;
                    border-collapse: collapse;
                    page-break-inside: avoid;
                    margin: 1em 0;
                }
                
                table td, table th {
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: left;
                }
                
                img {
                    max-width: 100%;
                    height: auto;
                    page-break-inside: avoid;
                }
                
                ul, ol {
                    margin: 0.5em 0;
                    padding-left: 2em;
                }
                
                li {
                    margin: 0.25em 0;
                }
                
                .page-break {
                    page-break-before: always;
                }
                
                /* Avoid breaking these elements */
                .no-break {
                    page-break-inside: avoid;
                }
            </style>
        </head>
        <body>
            ${htmlContent}
        </body>
        </html>`;
    }

    // Convert document files to PDF
    static async convertDocumentToPDF(file) {
        try {
            console.log('📝 文書檔案轉PDF (直接渲染)');
            
            // For DOCX files, we need to render them properly
            // This is a complex task that typically requires server-side conversion
            // For now, we'll use mammoth.js if available, otherwise show the file content
            
            const fileType = DirectPDFConverter.getFileType(file);
            
            if (fileType === 'docx') {
                return await DirectPDFConverter.convertDOCXToPDF(file);
            } else {
                // For other document formats, extract text and format nicely
                const textContent = await DirectPDFConverter.extractTextFromDocument(file);
                return await DirectPDFConverter.renderTextAsPDF(textContent, file.name);
            }
            
        } catch (error) {
            console.error('文書轉PDF失敗:', error);
            throw error;
        }
    }

    // Convert DOCX to PDF using mammoth.js
    static async convertDOCXToPDF(file) {
        try {
            console.log('📄 開始DOCX轉PDF轉換');
            console.log('📁 檔案資訊:', { name: file.name, size: file.size, type: file.type });
            
            // Try to load mammoth.js
            await DirectPDFConverter.loadMammoth();
            
            if (typeof mammoth !== 'undefined') {
                console.log('📄 使用 Mammoth.js 解析 DOCX');
                
                const arrayBuffer = await file.arrayBuffer();
                const result = await mammoth.convertToHtml({ arrayBuffer });
                
                // Check if we got meaningful content
                if (result.html && result.html.trim().length > 50) {
                    console.log('✅ Mammoth.js 解析成功');
                    
                    // Create HTML document with proper styling
                    const htmlDoc = `
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset="utf-8">
                        <title>文書轉換</title>
                        <style>
                            body { 
                                font-family: 'Times New Roman', 'PingFang TC', 'Microsoft JhengHei', 'Noto Serif CJK TC', 'SimSun', serif; 
                                font-size: 12pt; 
                                line-height: 1.5; 
                                margin: 2.5cm;
                                background: white;
                                color: #333;
                            }
                            h1, h2, h3, h4, h5, h6 { 
                                color: #2c3e50; 
                                margin-top: 20px; 
                                margin-bottom: 10px;
                            }
                            p { margin-bottom: 10px; }
                            ul, ol { margin-bottom: 10px; }
                            table { 
                                border-collapse: collapse; 
                                width: 100%; 
                                margin-bottom: 20px; 
                            }
                            table td, table th { 
                                border: 1px solid #ddd; 
                                padding: 8px; 
                            }
                            @media print {
                                @page { 
                                    margin: 2.5cm; 
                                    size: A4;
                                }
                            }
                        </style>
                    </head>
                    <body>
                        ${result.html}
                    </body>
                    </html>`;
                    
                    // Create temporary HTML file and convert
                    const htmlBlob = new Blob([htmlDoc], { type: 'text/html' });
                    const htmlFile = new File([htmlBlob], 'temp.html', { type: 'text/html' });
                    
                    return await DirectPDFConverter.convertHTMLToPDF(htmlFile);
                } else {
                    console.warn('Mammoth.js 解析結果不完整，使用備用方法');
                    throw new Error('Mammoth.js 解析內容不足');
                }
                
            } else {
                throw new Error('Mammoth.js 未載入');
            }
            
        } catch (mammothError) {
            console.warn('Mammoth.js 轉換失敗，嘗試直接解析 DOCX:', mammothError);
            
            // Try direct DOCX text extraction as fallback
            try {
                const textContent = await DirectPDFConverter.extractDOCXText(file);
                if (textContent && textContent.trim()) {
                    console.log('📄 使用直接文字提取方法');
                    return await DirectPDFConverter.renderTextAsPDF(textContent, file.name);
                }
            } catch (extractError) {
                console.warn('直接文字提取也失敗:', extractError);
            }
            
            // Final fallback to generic text extraction
            console.log('📄 使用最基本的備用方案');
            const textContent = await DirectPDFConverter.extractTextFromDocument(file);
            return await DirectPDFConverter.renderTextAsPDF(textContent, file.name);
        }
    }

    // Direct DOCX text extraction (backup method)
    static async extractDOCXText(file) {
        try {
            // Load JSZip to parse DOCX structure
            await DirectPDFConverter.loadJSZip();
            
            const arrayBuffer = await file.arrayBuffer();
            const zip = new JSZip();
            const zipContent = await zip.loadAsync(arrayBuffer);
            
            // Look for the main document XML
            const documentXml = zipContent.file('word/document.xml');
            if (documentXml) {
                const xmlContent = await documentXml.async('text');
                
                // Extract text from XML (basic approach)
                const textMatches = xmlContent.match(/<w:t[^>]*>([^<]+)<\/w:t>/g) || [];
                const paragraphMatches = xmlContent.match(/<w:p[^>]*>/g) || [];
                
                let extractedText = '';
                let currentParagraph = '';
                
                textMatches.forEach(match => {
                    const text = match.replace(/<[^>]+>/g, '').trim();
                    if (text) {
                        currentParagraph += text + ' ';
                    }
                });
                
                // Add paragraph breaks
                const paragraphs = currentParagraph.split(/\s{3,}/).filter(p => p.trim());
                extractedText = paragraphs.join('\n\n');
                
                if (extractedText.trim()) {
                    return `文書檔案: ${file.name}\n\n${extractedText}`;
                }
            }
            
            return `文書檔案: ${file.name}\n\n檔案大小: ${DirectPDFConverter.formatFileSize(file.size)}\n⚠️ 此DOCX檔案的內容無法完全解析，建議使用Microsoft Word或相容軟體開啟`;
            
        } catch (error) {
            console.warn('DOCX直接解析失敗:', error);
            return `文書檔案: ${file.name}\n檔案大小: ${DirectPDFConverter.formatFileSize(file.size)}\n❌ 檔案解析失敗: ${error.message}`;
        }
    }

    // Convert presentation files to PDF
    static async convertPresentationToPDF(file) {
        try {
            console.log('🎯 簡報檔案轉PDF (投影片模式)');
            
            const fileType = DirectPDFConverter.getFileType(file);
            
            if (fileType === 'pptx') {
                // For PPTX, try to extract slide images if possible
                return await DirectPDFConverter.convertPPTXToPDF(file);
            } else {
                // For other formats, show file info
                return await DirectPDFConverter.createPresentationPlaceholderPDF(file);
            }
            
        } catch (error) {
            console.error('簡報轉PDF失敗:', error);
            throw error;
        }
    }

    // Convert PPTX to PDF by extracting text content (simplified approach)
    static async convertPPTXToPDF(file) {
        try {
            console.log('🎯 開始PPTX轉PDF (文字提取模式)');
            console.log('📁 檔案資訊:', { name: file.name, size: file.size, type: file.type });
            
            // Load JSZip for PPTX processing
            await DirectPDFConverter.loadJSZip();
            
            const arrayBuffer = await file.arrayBuffer();
            console.log('📦 檔案載入完成，大小:', arrayBuffer.byteLength, '位元組');
            
            const zip = new JSZip();
            const zipContent = await zip.loadAsync(arrayBuffer);
            console.log('🗂️ ZIP內容載入完成，檔案數量:', Object.keys(zipContent.files).length);
            
            // Extract text content from slides
            const textContent = await DirectPDFConverter.extractPPTXText(zipContent);
            console.log('📄 提取的文字內容長度:', textContent?.length || 0);
            
            if (textContent && textContent.trim()) {
                console.log('📊 成功提取PPTX文字內容');
                
                // Convert extracted text to HTML format for better presentation
                const formattedHTML = DirectPDFConverter.formatPPTXTextAsHTML(textContent, file.name);
                
                // Create HTML file and convert to PDF
                const htmlBlob = new Blob([formattedHTML], { type: 'text/html' });
                const htmlFile = new File([htmlBlob], 'temp.html', { type: 'text/html' });
                
                return await DirectPDFConverter.convertHTMLToPDF(htmlFile);
            } else {
                console.warn('無法提取PPTX文字內容，建立預覽版本');
                return await DirectPDFConverter.createPresentationPlaceholderPDF(file);
            }
            
        } catch (error) {
            console.error('PPTX轉PDF失敗:', error);
            return await DirectPDFConverter.createPresentationPlaceholderPDF(file);
        }
    }

    // Format PPTX text content as HTML for better PDF presentation
    static formatPPTXTextAsHTML(textContent, fileName) {
        const slides = textContent.split('--- 投影片').filter(slide => slide.trim());
        
        let htmlContent = '';
        
        slides.forEach((slideText, index) => {
            if (index > 0) {
                htmlContent += '<div class="page-break"></div>';
            }
            
            const lines = slideText.split('\n').filter(line => line.trim());
            
            htmlContent += `
            <div class="slide">
                <h2 style="color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px;">
                    投影片 ${index + 1}
                </h2>`;
            
            lines.forEach((line, lineIndex) => {
                const cleanLine = line.trim();
                if (cleanLine) {
                    if (lineIndex === 0 && cleanLine.length < 100) {
                        // Likely a title
                        htmlContent += `<h3 style="color: #2980b9; margin: 15px 0;">${cleanLine}</h3>`;
                    } else {
                        // Content
                        htmlContent += `<p style="margin: 10px 0; line-height: 1.6;">${cleanLine}</p>`;
                    }
                }
            });
            
            htmlContent += '</div>';
        });
        
        return DirectPDFConverter.createStyledHTML(htmlContent, `簡報: ${fileName}`);
    }

    // Create placeholder PDF for presentations
    static async createPresentationPlaceholderPDF(file) {
        const content = `
        <div style="text-align: center; padding: 50px; font-family: 'PingFang TC', 'Microsoft JhengHei', 'Noto Sans CJK TC', 'Microsoft YaHei', Arial, sans-serif;">
            <h1>📊 簡報檔案</h1>
            <h2>${file.name}</h2>
            <p>檔案大小: ${DirectPDFConverter.formatFileSize(file.size)}</p>
            <p>檔案類型: ${DirectPDFConverter.getFileType(file).toUpperCase()}</p>
            <br>
            <p>⚠️ 此為簡報檔案的預覽版本</p>
            <p>完整內容需要專業的簡報軟體開啟</p>
            <br>
            <p>建議使用以下軟體開啟原檔案：</p>
            <ul style="text-align: left; display: inline-block;">
                <li>Microsoft PowerPoint</li>
                <li>LibreOffice Impress</li>
                <li>Google Slides</li>
            </ul>
        </div>`;
        
        const htmlDoc = DirectPDFConverter.createStyledHTML(content, '簡報檔案預覽');
        const htmlBlob = new Blob([htmlDoc], { type: 'text/html' });
        const htmlFile = new File([htmlBlob], 'temp.html', { type: 'text/html' });
        
        return await DirectPDFConverter.convertHTMLToPDF(htmlFile);
    }

    // Convert text files to PDF
    static async convertTextToPDF(file) {
        try {
            console.log('📝 文字檔案轉PDF');
            
            const textContent = await file.text();
            return await DirectPDFConverter.renderTextAsPDF(textContent, file.name);
            
        } catch (error) {
            console.error('文字轉PDF失敗:', error);
            throw error;
        }
    }

    // Render text content as PDF with proper formatting
    static async renderTextAsPDF(textContent, filename) {
        // Create formatted HTML
        const formattedContent = textContent
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/\n/g, '<br>');
        
        const htmlDoc = DirectPDFConverter.createStyledHTML(
            `<div style="white-space: pre-wrap; font-family: 'Courier New', 'PingFang TC', 'Microsoft JhengHei', monospace;">${formattedContent}</div>`,
            filename || '文字文件'
        );
        
        const htmlBlob = new Blob([htmlDoc], { type: 'text/html' });
        const htmlFile = new File([htmlBlob], 'temp.html', { type: 'text/html' });
        
        return await DirectPDFConverter.convertHTMLToPDF(htmlFile);
    }

    // Create styled HTML document
    static createStyledHTML(content, title) {
        return `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>${title}</title>
            <style>
                body { 
                    font-family: 'PingFang TC', 'Microsoft JhengHei', 'Noto Sans CJK TC', 'Microsoft YaHei', Arial, sans-serif; 
                    font-size: 12pt; 
                    line-height: 1.6; 
                    margin: 2.5cm;
                    background: white;
                    color: #333;
                }
                h1, h2, h3 { color: #2c3e50; }
                @media print {
                    @page { 
                        margin: 2.5cm; 
                        size: A4;
                    }
                }
            </style>
        </head>
        <body>
            ${content}
        </body>
        </html>`;
    }

    // Extract text from document files
    static async extractTextFromDocument(file) {
        try {
            const fileType = DirectPDFConverter.getFileType(file);
            
            switch (fileType) {
                case 'txt':
                case 'md':
                    return await file.text();
                
                case 'rtf':
                    const rtfText = await file.text();
                    // Simple RTF to text conversion
                    return rtfText
                        .replace(/\\[a-z]+\d*\s?/gi, '') // Remove RTF commands
                        .replace(/[{}]/g, '') // Remove braces
                        .replace(/\s+/g, ' ') // Normalize whitespace
                        .trim();
                
                default:
                    return `檔案類型: ${fileType.toUpperCase()}\n檔案名稱: ${file.name}\n檔案大小: ${DirectPDFConverter.formatFileSize(file.size)}\n\n⚠️ 此檔案格式需要專業軟體才能完整顯示內容`;
            }
        } catch (error) {
            return `檔案讀取失敗: ${error.message}`;
        }
    }

    // Utility methods
    static getFileType(file) {
        return file.name.toLowerCase().substring(file.name.lastIndexOf('.') + 1);
    }

    static formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // PPTX specific methods
    static async extractPPTXSlides(zipContent) {
        try {
            const slides = [];
            
            // Look for slide XML files
            zipContent.forEach((relativePath, zipEntry) => {
                if (relativePath.startsWith('ppt/slides/slide') && relativePath.endsWith('.xml')) {
                    slides.push({
                        path: relativePath,
                        entry: zipEntry,
                        number: parseInt(relativePath.match(/slide(\d+)\.xml/)?.[1] || '0')
                    });
                }
            });
            
            // Sort slides by number
            slides.sort((a, b) => a.number - b.number);
            
            // Extract content from each slide
            const slideContents = [];
            for (const slide of slides) {
                try {
                    const xmlContent = await slide.entry.async('text');
                    const slideContent = DirectPDFConverter.parseSlideXML(xmlContent, slide.number);
                    slideContents.push(slideContent);
                } catch (error) {
                    console.warn(`解析投影片 ${slide.number} 失敗:`, error);
                    slideContents.push({
                        number: slide.number,
                        title: `投影片 ${slide.number}`,
                        content: '此投影片內容無法解析',
                        textElements: []
                    });
                }
            }
            
            return slideContents;
        } catch (error) {
            console.error('PPTX投影片提取失敗:', error);
            return null;
        }
    }

    static parseSlideXML(xmlContent, slideNumber) {
        try {
            // Basic XML text extraction for PPTX slides
            // This is a simplified parser - real PPTX parsing is complex
            
            const textElements = [];
            let title = '';
            
            // Extract text content using regex (simplified approach)
            const textMatches = xmlContent.match(/<a:t[^>]*>([^<]+)<\/a:t>/g) || [];
            
            textMatches.forEach((match, index) => {
                const text = match.replace(/<[^>]+>/g, '').trim();
                if (text) {
                    if (index === 0 && text.length < 100) {
                        title = text; // First short text is likely the title
                    }
                    textElements.push({
                        text: text,
                        type: index === 0 ? 'title' : 'content'
                    });
                }
            });
            
            return {
                number: slideNumber,
                title: title || `投影片 ${slideNumber}`,
                content: textElements.map(el => el.text).join('\n'),
                textElements: textElements
            };
        } catch (error) {
            console.warn(`解析投影片 ${slideNumber} XML失敗:`, error);
            return {
                number: slideNumber,
                title: `投影片 ${slideNumber}`,
                content: '投影片內容解析失敗',
                textElements: []
            };
        }
    }

    static async renderSlideAsHTML(slide, slideNumber) {
        const titleText = slide.title || `投影片 ${slideNumber}`;
        const contentElements = slide.textElements || [];
        
        let contentHTML = '';
        
        if (contentElements.length > 0) {
            contentElements.forEach(element => {
                if (element.type === 'title') {
                    contentHTML += `<h1 style="font-size: 24px; color: #2c3e50; margin-bottom: 20px; text-align: center;">${element.text}</h1>`;
                } else {
                    contentHTML += `<p style="font-size: 16px; line-height: 1.6; margin-bottom: 15px;">${element.text}</p>`;
                }
            });
        } else {
            contentHTML = `<h1 style="font-size: 24px; color: #2c3e50; text-align: center; margin-top: 50px;">${titleText}</h1>`;
            if (slide.content && slide.content !== titleText) {
                contentHTML += `<div style="font-size: 16px; line-height: 1.6; margin-top: 30px; white-space: pre-wrap;">${slide.content}</div>`;
            }
        }
        
        return `
        <div style="
            width: 297mm; 
            height: 210mm; 
            padding: 20mm; 
            background: white; 
            font-family: 'PingFang TC', 'Microsoft JhengHei', 'Noto Sans CJK TC', 'Microsoft YaHei', Arial, sans-serif;
            display: flex;
            flex-direction: column;
            justify-content: center;
        ">
            ${contentHTML}
            <div style="position: absolute; bottom: 10mm; right: 15mm; font-size: 12px; color: #666;">
                ${slideNumber}
            </div>
        </div>`;
    }

    static async addHTMLToPDFPage(pdf, htmlContent) {
        try {
            // Create temporary element to render HTML
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = htmlContent;
            tempDiv.style.position = 'fixed';
            tempDiv.style.top = '-9999px';
            tempDiv.style.left = '-9999px';
            tempDiv.style.width = '297mm';
            tempDiv.style.height = '210mm';
            document.body.appendChild(tempDiv);
            
            // Load html2canvas if not already loaded
            await DirectPDFConverter.loadHTML2Canvas();
            
            // Capture the HTML as canvas
            const canvas = await html2canvas(tempDiv, {
                scale: 3, // 提高解析度 (從1提升到3)
                useCORS: true,
                allowTaint: true,
                backgroundColor: '#ffffff',
                width: 1123,  // A4 landscape width in pixels
                height: 794   // A4 landscape height in pixels  
            });
            
            // Add canvas image to PDF
            const imgData = canvas.toDataURL('image/png');
            pdf.addImage(imgData, 'PNG', 0, 0, 297, 210);
            
            // Clean up
            document.body.removeChild(tempDiv);
            
        } catch (error) {
            console.warn('HTML轉PDF頁面失敗:', error);
            // Add text fallback
            pdf.setFontSize(16);
            pdf.text('投影片渲染失敗', 20, 30);
        }
    }

    static async extractPPTXText(zipContent) {
        try {
            let allText = '';
            const slideFiles = [];
            
            // Collect slide XML files
            zipContent.forEach((relativePath, zipEntry) => {
                if (relativePath.startsWith('ppt/slides/slide') && relativePath.endsWith('.xml')) {
                    const slideNumber = parseInt(relativePath.match(/slide(\d+)\.xml/)?.[1] || '0');
                    slideFiles.push({ path: relativePath, entry: zipEntry, number: slideNumber });
                }
            });
            
            // Sort by slide number
            slideFiles.sort((a, b) => a.number - b.number);
            
            // Extract text from each slide
            for (const slideFile of slideFiles) {
                try {
                    const xmlContent = await slideFile.entry.async('text');
                    const textMatches = xmlContent.match(/<a:t[^>]*>([^<]+)<\/a:t>/g) || [];
                    
                    const slideTexts = textMatches.map(match => 
                        match.replace(/<[^>]+>/g, '').trim()
                    ).filter(text => text.length > 0);
                    
                    if (slideTexts.length > 0) {
                        allText += `\n\n--- 投影片 ${slideFile.number} ---\n`;
                        allText += slideTexts.join('\n');
                    }
                } catch (error) {
                    console.warn(`提取投影片文字失敗: ${slideFile.path}`, error);
                }
            }
            
            return allText.trim();
        } catch (error) {
            console.error('PPTX文字提取失敗:', error);
            return '';
        }
    }

    // Load required libraries
    static async loadJsPDF() {
        if (window.jsPDF || (window.jspdf && window.jspdf.jsPDF)) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/jspdf@2.5.1/dist/jspdf.umd.min.js';
            script.onload = () => {
                setTimeout(() => {
                    if (window.jsPDF || (window.jspdf && window.jspdf.jsPDF)) {
                        resolve();
                    } else {
                        reject(new Error('jsPDF 載入失敗'));
                    }
                }, 100);
            };
            script.onerror = () => reject(new Error('jsPDF 載入失敗'));
            document.head.appendChild(script);
        });
    }

    static async loadHTML2Canvas() {
        if (window.html2canvas) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/html2canvas@1.4.1/dist/html2canvas.min.js';
            script.onload = () => {
                setTimeout(() => {
                    if (window.html2canvas) {
                        resolve();
                    } else {
                        reject(new Error('html2canvas 載入失敗'));
                    }
                }, 100);
            };
            script.onerror = () => reject(new Error('html2canvas 載入失敗'));
            document.head.appendChild(script);
        });
    }

    static async loadJSZip() {
        if (window.JSZip) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/jszip@3.10.1/dist/jszip.min.js';
            script.onload = () => {
                setTimeout(() => {
                    if (window.JSZip) {
                        resolve();
                    } else {
                        reject(new Error('JSZip 載入失敗'));
                    }
                }, 100);
            };
            script.onerror = () => reject(new Error('JSZip 載入失敗'));
            document.head.appendChild(script);
        });
    }

    static async loadMammoth() {
        if (window.mammoth) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://unpkg.com/mammoth@1.4.2/mammoth.browser.min.js';
            script.onload = () => {
                setTimeout(() => {
                    if (window.mammoth) {
                        resolve();
                    } else {
                        reject(new Error('Mammoth.js 載入失敗'));
                    }
                }, 100);
            };
            script.onerror = () => reject(new Error('Mammoth.js 載入失敗'));
            document.head.appendChild(script);
        });
    }
}

// Export for use
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DirectPDFConverter;
} else if (typeof window !== 'undefined') {
    window.DirectPDFConverter = DirectPDFConverter;
}
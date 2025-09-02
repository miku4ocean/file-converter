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

    // Convert HTML files to PDF (like browser print)
    static async convertHTMLToPDF(file) {
        try {
            console.log('🌐 HTML檔案轉PDF (瀏覽器列印模式)');
            
            const htmlContent = await file.text();
            
            // Create a hidden iframe to render the HTML
            const iframe = document.createElement('iframe');
            iframe.style.position = 'fixed';
            iframe.style.top = '-9999px';
            iframe.style.left = '-9999px';
            iframe.style.width = '210mm';  // A4 width
            iframe.style.height = '297mm'; // A4 height
            iframe.style.border = 'none';
            document.body.appendChild(iframe);
            
            // Load content into iframe
            iframe.contentDocument.open();
            iframe.contentDocument.write(htmlContent);
            iframe.contentDocument.close();
            
            // Wait for content to load
            await new Promise(resolve => {
                iframe.onload = resolve;
                setTimeout(resolve, 1000); // Fallback timeout
            });
            
            // Use html2canvas to capture the rendered content
            await DirectPDFConverter.loadHTML2Canvas();
            const canvas = await html2canvas(iframe.contentDocument.body, {
                scale: 3, // 提高解析度 (從2提升到3)
                useCORS: true,
                allowTaint: true,
                backgroundColor: '#ffffff',
                width: 794,  // A4 width in pixels
                height: 1123 // A4 height in pixels
            });
            
            // Convert canvas to PDF
            await DirectPDFConverter.loadJsPDF();
            const { jsPDF } = window.jspdf || window;
            const pdf = new jsPDF({
                orientation: 'portrait',
                unit: 'mm',
                format: 'a4'
            });
            
            const imgData = canvas.toDataURL('image/png');
            pdf.addImage(imgData, 'PNG', 0, 0, 210, 297);
            
            // Clean up
            document.body.removeChild(iframe);
            
            const pdfBlob = pdf.output('blob');
            console.log('✅ HTML轉PDF完成:', pdfBlob.size, 'bytes');
            return pdfBlob;
            
        } catch (error) {
            console.error('HTML轉PDF失敗:', error);
            throw error;
        }
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
                                font-family: 'Times New Roman', '微軟正黑體', serif; 
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

    // Convert PPTX to PDF by extracting and rendering slide content
    static async convertPPTXToPDF(file) {
        try {
            console.log('🎯 開始PPTX轉PDF (內容解析模式)');
            
            // Load JSZip for PPTX processing
            await DirectPDFConverter.loadJSZip();
            
            const arrayBuffer = await file.arrayBuffer();
            const zip = new JSZip();
            const zipContent = await zip.loadAsync(arrayBuffer);
            
            // Extract slide content from PPTX structure
            const slides = await DirectPDFConverter.extractPPTXSlides(zipContent);
            
            if (slides && slides.length > 0) {
                console.log(`📊 解析到 ${slides.length} 張投影片`);
                
                // Create PDF with rendered slides
                await DirectPDFConverter.loadJsPDF();
                const { jsPDF } = window.jspdf || window;
                const pdf = new jsPDF({
                    orientation: 'landscape', // Presentations are typically landscape
                    unit: 'mm',
                    format: 'a4'
                });
                
                for (let i = 0; i < slides.length; i++) {
                    try {
                        const slide = slides[i];
                        
                        // Add page for each slide (except first)
                        if (i > 0) {
                            pdf.addPage();
                        }
                        
                        // Render slide content as HTML and convert to PDF page
                        const slideHTML = await DirectPDFConverter.renderSlideAsHTML(slide, i + 1);
                        await DirectPDFConverter.addHTMLToPDFPage(pdf, slideHTML);
                        
                    } catch (slideError) {
                        console.warn(`投影片 ${i + 1} 渲染失敗:`, slideError);
                        // Add error slide
                        pdf.setFontSize(16);
                        pdf.text(`投影片 ${i + 1} (載入失敗)`, 20, 30);
                        pdf.setFontSize(12);
                        pdf.text('此投影片內容無法正確解析', 20, 50);
                    }
                }
                
                const pdfBlob = pdf.output('blob');
                console.log(`✅ PPTX轉PDF完成: ${slides.length} 頁`);
                return pdfBlob;
            }
            
            // Fallback: try to extract basic text content
            console.log('🔄 使用備用方法解析PPTX');
            const textContent = await DirectPDFConverter.extractPPTXText(zipContent);
            if (textContent && textContent.trim()) {
                return await DirectPDFConverter.renderTextAsPDF(textContent, file.name);
            }
            
            // Final fallback: create placeholder
            return await DirectPDFConverter.createPresentationPlaceholderPDF(file);
            
        } catch (error) {
            console.warn('PPTX 解析失敗，建立備用版本:', error);
            // Try to extract any text content as last resort
            try {
                const arrayBuffer = await file.arrayBuffer();
                const zip = new JSZip();
                const zipContent = await zip.loadAsync(arrayBuffer);
                const textContent = await DirectPDFConverter.extractPPTXText(zipContent);
                if (textContent && textContent.trim()) {
                    console.log('📝 使用文字內容作為備用方案');
                    return await DirectPDFConverter.renderTextAsPDF(textContent, file.name);
                }
            } catch (extractError) {
                console.warn('文字提取也失敗:', extractError);
            }
            
            return await DirectPDFConverter.createPresentationPlaceholderPDF(file);
        }
    }

    // Create placeholder PDF for presentations
    static async createPresentationPlaceholderPDF(file) {
        const content = `
        <div style="text-align: center; padding: 50px; font-family: Arial, sans-serif;">
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
            `<div style="white-space: pre-wrap; font-family: 'Courier New', monospace;">${formattedContent}</div>`,
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
                    font-family: Arial, '微軟正黑體', sans-serif; 
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
            font-family: Arial, '微軟正黑體', sans-serif;
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
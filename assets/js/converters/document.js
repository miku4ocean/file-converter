// Document conversion utilities

class DocumentConverter {
    constructor() {
        this.supportedFormats = {
            input: ['docx', 'doc', 'rtf', 'odt', 'html', 'txt', 'md'],
            output: ['txt', 'html', 'md', 'pdf', 'docx', 'rtf']
        };
    }

    // Validate document file
    static isValidDocumentFile(file) {
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // docx
            'application/msword', // doc
            'application/rtf', // rtf
            'text/rtf',
            'application/vnd.oasis.opendocument.text', // odt
            'text/html',
            'text/plain',
            'text/markdown',
            'text/x-markdown'
        ];
        
        const validExtensions = ['.docx', '.doc', '.rtf', '.odt', '.html', '.htm', '.txt', '.md'];
        const fileExtension = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));
        
        return validTypes.includes(file.type) || validExtensions.includes(fileExtension);
    }

    // Extract text content from various document formats
    static async extractTextContent(file) {
        const fileType = DocumentConverter.getFileType(file);
        
        switch (fileType) {
            case 'docx':
                return await DocumentConverter.extractFromDocx(file);
            case 'txt':
                return await DocumentConverter.extractFromText(file);
            case 'html':
                return await DocumentConverter.extractFromHtml(file);
            case 'md':
                return await DocumentConverter.extractFromMarkdown(file);
            case 'rtf':
                return await DocumentConverter.extractFromRtf(file);
            default:
                throw new Error(`不支援的文件格式: ${fileType}`);
        }
    }

    // Get file type from file object
    static getFileType(file) {
        const extension = file.name.toLowerCase().substring(file.name.lastIndexOf('.') + 1);
        return extension;
    }

    // Extract content from DOCX files (requires mammoth.js CDN)
    static async extractFromDocx(file) {
        try {
            // Note: This would require mammoth.js library
            // For now, return a placeholder implementation
            const arrayBuffer = await file.arrayBuffer();
            
            // Placeholder: In real implementation, use mammoth.js
            return {
                content: '此功能需要 mammoth.js 函式庫支援\n請在 HTML 中加入：\n<script src="https://cdn.jsdelivr.net/npm/mammoth@1.4.2/mammoth.browser.min.js"></script>',
                title: file.name.replace(/\.[^/.]+$/, ''),
                wordCount: 0
            };
        } catch (error) {
            throw new Error('DOCX 檔案讀取失敗: ' + error.message);
        }
    }

    // Extract content from plain text files
    static async extractFromText(file) {
        try {
            const text = await file.text();
            return {
                content: text,
                title: file.name.replace(/\.[^/.]+$/, ''),
                wordCount: text.split(/\s+/).filter(word => word.length > 0).length
            };
        } catch (error) {
            throw new Error('文字檔案讀取失敗: ' + error.message);
        }
    }

    // Extract content from HTML files
    static async extractFromHtml(file) {
        try {
            const html = await file.text();
            const parser = new DOMParser();
            const doc = parser.parseFromString(html, 'text/html');
            
            // Extract title
            const titleElement = doc.querySelector('title');
            const title = titleElement ? titleElement.textContent : file.name.replace(/\.[^/.]+$/, '');
            
            // Extract content (remove scripts and styles)
            const scripts = doc.querySelectorAll('script, style');
            scripts.forEach(el => el.remove());
            
            const content = doc.body ? doc.body.textContent || doc.body.innerText : doc.documentElement.textContent;
            
            return {
                content: content.trim(),
                title: title,
                wordCount: content.split(/\s+/).filter(word => word.length > 0).length,
                originalHtml: html
            };
        } catch (error) {
            throw new Error('HTML 檔案讀取失敗: ' + error.message);
        }
    }

    // Extract content from Markdown files
    static async extractFromMarkdown(file) {
        try {
            const markdown = await file.text();
            
            // Simple markdown to text conversion
            let content = markdown
                .replace(/^#+ /gm, '') // Remove headers
                .replace(/\*\*(.*?)\*\*/g, '$1') // Remove bold
                .replace(/\*(.*?)\*/g, '$1') // Remove italic
                .replace(/\[(.*?)\]\(.*?\)/g, '$1') // Remove links, keep text
                .replace(/`(.*?)`/g, '$1') // Remove code backticks
                .replace(/```[\s\S]*?```/g, '') // Remove code blocks
                .trim();
            
            return {
                content: content,
                title: file.name.replace(/\.[^/.]+$/, ''),
                wordCount: content.split(/\s+/).filter(word => word.length > 0).length,
                originalMarkdown: markdown
            };
        } catch (error) {
            throw new Error('Markdown 檔案讀取失敗: ' + error.message);
        }
    }

    // Basic RTF text extraction
    static async extractFromRtf(file) {
        try {
            const rtf = await file.text();
            
            // Simple RTF to text conversion (basic implementation)
            let content = rtf
                .replace(/\\[a-z]+\d*\s?/gi, '') // Remove RTF commands
                .replace(/[{}]/g, '') // Remove braces
                .replace(/\s+/g, ' ') // Normalize whitespace
                .trim();
            
            return {
                content: content,
                title: file.name.replace(/\.[^/.]+$/, ''),
                wordCount: content.split(/\s+/).filter(word => word.length > 0).length
            };
        } catch (error) {
            throw new Error('RTF 檔案讀取失敗: ' + error.message);
        }
    }

    // Convert extracted content to various formats
    static async convertToFormat(extractedContent, outputFormat, options = {}) {
        const { title, content, originalHtml, originalMarkdown } = extractedContent;
        const { maxFileSize = null, compression = 'standard' } = options;

        switch (outputFormat) {
            case 'txt':
                return DocumentConverter.convertToText(content, title);
            case 'html':
                return DocumentConverter.convertToHtml(content, title, originalHtml);
            case 'md':
                return DocumentConverter.convertToMarkdown(content, title, originalMarkdown);
            case 'pdf':
                return await DocumentConverter.convertToPdf(content, title, options);
            case 'docx':
                return await DocumentConverter.convertToDocx(content, title, options);
            case 'rtf':
                return DocumentConverter.convertToRtf(content, title, options);
            default:
                throw new Error(`不支援的輸出格式: ${outputFormat}`);
        }
    }

    // Convert to plain text
    static convertToText(content, title) {
        let output = '';
        if (title) {
            output += title + '\n';
            output += '='.repeat(title.length) + '\n\n';
        }
        output += content;
        
        const blob = new Blob([output], { type: 'text/plain;charset=utf-8' });
        return blob;
    }

    // Convert to HTML
    static convertToHtml(content, title, originalHtml = null) {
        let html;
        
        if (originalHtml) {
            html = originalHtml;
        } else {
            const paragraphs = content.split('\n\n').map(p => 
                p.trim() ? `<p>${p.replace(/\n/g, '<br>')}</p>` : ''
            ).filter(p => p).join('\n');
            
            html = `<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${title || '轉換後的文件'}</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.6;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            color: #333;
        }
        h1 { color: #2c3e50; }
        p { margin-bottom: 1em; }
    </style>
</head>
<body>
    <h1>${title || '轉換後的文件'}</h1>
    ${paragraphs}
</body>
</html>`;
        }
        
        const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
        return blob;
    }

    // Convert to Markdown
    static convertToMarkdown(content, title, originalMarkdown = null) {
        let markdown;
        
        if (originalMarkdown) {
            markdown = originalMarkdown;
        } else {
            markdown = '';
            if (title) {
                markdown += `# ${title}\n\n`;
            }
            
            // Convert paragraphs to markdown
            const paragraphs = content.split('\n\n');
            markdown += paragraphs.join('\n\n');
        }
        
        const blob = new Blob([markdown], { type: 'text/markdown;charset=utf-8' });
        return blob;
    }

    // Convert to PDF (requires jsPDF)
    static async convertToPdf(content, title, options = {}) {
        try {
            console.log('開始 PDF 轉換...');
            
            // Load jsPDF library if not already loaded
            try {
                await window.libLoader.loadLibrary('jspdf');
                return await DocumentConverter.createPdfWithJsPDF(content, title, options);
            } catch (libError) {
                console.warn('jsPDF 載入失敗，使用 HTML 回退方式:', libError.message);
                return DocumentConverter.createHtmlToPdf(content, title, options);
            }
            
        } catch (error) {
            console.error('PDF 轉換錯誤:', error);
            throw new Error('PDF 轉換失敗: ' + error.message);
        }
    }

    // Create PDF using jsPDF library
    static async createPdfWithJsPDF(content, title, options = {}) {
        try {
            // Ensure jsPDF is properly loaded
            if (!window.jsPDF && !(window.jspdf && window.jspdf.jsPDF)) {
                await DocumentConverter.loadJsPDF();
            }
            
            // 支援不同的 jsPDF 匯入方式
            const jsPDF = window.jsPDF || (window.jspdf && window.jspdf.jsPDF);
            if (!jsPDF) {
                throw new Error('jsPDF 函式庫未正確載入');
            }
            
            const doc = new jsPDF({
                orientation: 'portrait',
                unit: 'mm',
                format: 'a4'
            });
            
            // Enable UTF-8 support
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(12);
            
            let yPosition = 20;
            const pageHeight = 280;
            const margin = 20;
            const lineHeight = 7;
            const pageWidth = 170;
            
            // Add title if provided
            if (title) {
                doc.setFontSize(18);
                doc.setFont('helvetica', 'bold');
                
                const titleLines = doc.splitTextToSize(title, pageWidth);
                titleLines.forEach(line => {
                    doc.text(line, margin, yPosition);
                    yPosition += lineHeight + 2;
                });
                
                yPosition += 5; // Extra space after title
                doc.setFont('helvetica', 'normal');
                doc.setFontSize(12);
            }
            
            // Process content paragraph by paragraph
            const paragraphs = content.split('\n\n');
            
            paragraphs.forEach(paragraph => {
                if (!paragraph.trim()) {
                    yPosition += lineHeight; // Empty line
                    return;
                }
                
                const lines = doc.splitTextToSize(paragraph.trim(), pageWidth);
                
                lines.forEach(line => {
                    // Check if we need a new page
                    if (yPosition > pageHeight - margin) {
                        doc.addPage();
                        yPosition = margin;
                    }
                    
                    doc.text(line, margin, yPosition);
                    yPosition += lineHeight;
                });
                
                yPosition += 3; // Space between paragraphs
            });
            
            // Add footer
            const pageCount = doc.internal.getNumberOfPages();
            for (let i = 1; i <= pageCount; i++) {
                doc.setPage(i);
                doc.setFontSize(10);
                doc.setTextColor(128, 128, 128);
                doc.text(`頁面 ${i} / ${pageCount}`, margin, pageHeight + 10);
                doc.text(`由檔案轉換器生成 - ${new Date().toLocaleDateString()}`, pageWidth - 50, pageHeight + 10);
            }
            
            // Generate PDF blob with proper headers
            const pdfBlob = doc.output('blob');
            
            // Verify the PDF is properly formatted
            if (pdfBlob.size < 1000) {
                throw new Error('生成的 PDF 檔案過小，可能有問題');
            }
            
            console.log('✅ 真正的 PDF 檔案創建完成:', pdfBlob.size, 'bytes');
            return pdfBlob;
            
        } catch (error) {
            console.error('jsPDF 轉換錯誤:', error);
            throw error;
        }
    }

    // Load jsPDF library
    static async loadJsPDF() {
        if (window.jsPDF || (window.jspdf && window.jspdf.jsPDF)) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
            script.onload = () => {
                // Wait a moment for the library to initialize
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

    // Fallback: Create HTML-based PDF
    static createHtmlToPdf(content, title, options = {}) {
        try {
            console.log('使用 HTML 回退方式創建 PDF 格式文件...');
            
            // Create printable HTML structure
            const htmlContent = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>${title || '文件'}</title>
    <style>
        @media print {
            @page { margin: 2cm; size: A4; }
        }
        body {
            font-family: Arial, "微軟正黑體", sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
        }
        h1 {
            color: #2c3e50;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
            margin-bottom: 30px;
        }
        p {
            margin-bottom: 15px;
            text-align: justify;
        }
        .footer {
            margin-top: 50px;
            text-align: center;
            font-size: 12px;
            color: #666;
            border-top: 1px solid #eee;
            padding-top: 20px;
        }
    </style>
</head>
<body>
    ${title ? `<h1>${DocumentConverter.escapeHtml(title)}</h1>` : ''}
    ${DocumentConverter.contentToHtmlParagraphs(content)}
    
    <div class="footer">
        <p>本文件由檔案格式轉換器生成 - ${new Date().toLocaleDateString()}</p>
        <p>提示：使用瀏覽器的「列印」功能可將此文件儲存為 PDF</p>
    </div>
</body>
</html>`;
            
            const blob = new Blob([htmlContent], { 
                type: 'text/html;charset=utf-8' 
            });
            
            console.log('HTML 格式轉換完成 (可列印為 PDF)');
            return blob;
            
        } catch (error) {
            console.error('HTML 轉換錯誤:', error);
            throw error;
        }
    }

    // Helper: Convert content to HTML paragraphs
    static contentToHtmlParagraphs(content) {
        return content
            .split('\n\n')
            .filter(p => p.trim())
            .map(p => `<p>${DocumentConverter.escapeHtml(p.trim())}</p>`)
            .join('\n');
    }

    // Helper: Escape HTML characters
    static escapeHtml(text) {
        if (!text) return '';
        return text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    // Compress content for file size optimization
    static compressContent(content, compressionLevel = 'standard') {
        switch (compressionLevel) {
            case 'high':
                // Aggressive compression
                return content
                    .replace(/\s+/g, ' ') // Multiple spaces to single
                    .replace(/\n\s*\n/g, '\n') // Multiple newlines to single
                    .trim();
            case 'medium':
                // Moderate compression
                return content
                    .replace(/[ \t]+/g, ' ') // Multiple spaces/tabs to single space
                    .replace(/\n{3,}/g, '\n\n') // Multiple newlines to double
                    .trim();
            case 'low':
            case 'standard':
            default:
                // Light compression
                return content
                    .replace(/[ \t]+/g, ' ') // Multiple spaces/tabs to single
                    .trim();
        }
    }

    // Get document statistics
    static getDocumentStats(content) {
        const words = content.split(/\s+/).filter(word => word.length > 0);
        const chars = content.length;
        const charsNoSpaces = content.replace(/\s/g, '').length;
        const lines = content.split('\n').length;
        const paragraphs = content.split(/\n\s*\n/).filter(p => p.trim().length > 0).length;
        
        return {
            wordCount: words.length,
            charCount: chars,
            charCountNoSpaces: charsNoSpaces,
            lineCount: lines,
            paragraphCount: paragraphs,
            avgWordsPerParagraph: Math.round(words.length / paragraphs),
            readingTimeMinutes: Math.ceil(words.length / 200) // Assume 200 WPM
        };
    }

    // Convert to DOCX format - create a real DOCX file using ZIP structure
    static async convertToDocx(content, title, options = {}) {
        try {
            console.log('開始轉換真正的 DOCX 格式...');
            
            const documentTitle = title || '文件';
            const documentContent = content || '';
            
            // Try to load JSZip for creating proper DOCX structure
            try {
                await DocumentConverter.loadJSZip();
            } catch (zipError) {
                console.warn('JSZip 載入失敗，使用 RTF 替代方案');
                return DocumentConverter.convertToRtf(content, title, options);
            }
            
            // Create proper DOCX structure using JSZip
            const zip = new JSZip();
            
            // Create [Content_Types].xml
            const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
            zip.file('[Content_Types].xml', contentTypes);
            
            // Create _rels/.rels
            const mainRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
            zip.folder('_rels').file('.rels', mainRels);
            
            // Create word/_rels/document.xml.rels
            const documentRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
            zip.folder('word').folder('_rels').file('document.xml.rels', documentRels);
            
            // Create the main document content
            const paragraphs = documentContent.split('\n\n').map(paragraph => {
                if (!paragraph.trim()) return '';
                return `<w:p><w:r><w:t xml:space="preserve">${DocumentConverter.escapeXml(paragraph.trim())}</w:t></w:r></w:p>`;
            }).join('');
            
            const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:sz w:val="32"/>
                </w:rPr>
                <w:t>${DocumentConverter.escapeXml(documentTitle)}</w:t>
            </w:r>
        </w:p>
        <w:p><w:r><w:t> </w:t></w:r></w:p>
        ${paragraphs}
        <w:p><w:r><w:t> </w:t></w:r></w:p>
        <w:p><w:r><w:rPr><w:sz w:val="18"/><w:color w:val="666666"/></w:rPr><w:t>由檔案轉換器生成 - ${new Date().toLocaleDateString()}</w:t></w:r></w:p>
        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>
            <w:cols w:space="708"/>
        </w:sectPr>
    </w:body>
</w:document>`;
            
            zip.folder('word').file('document.xml', documentXml);
            
            // Generate the ZIP file as blob
            const docxBlob = await zip.generateAsync({ 
                type: 'blob', 
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                compression: 'DEFLATE',
                compressionOptions: { level: 6 }
            });
            
            console.log('真正的 DOCX 檔案創建完成:', docxBlob.size, 'bytes');
            return docxBlob;
            
        } catch (error) {
            console.error('DOCX 轉換錯誤:', error);
            // Fallback to RTF format
            console.log('回退到 RTF 格式');
            return this.convertToRtf(content, title, options);
        }
    }

    // Load JSZip library for DOCX creation
    static async loadJSZip() {
        if (window.JSZip) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
            script.onload = () => {
                if (window.JSZip) {
                    console.log('✅ JSZip 載入成功');
                    resolve();
                } else {
                    reject(new Error('JSZip 載入後無法使用'));
                }
            };
            script.onerror = () => reject(new Error('JSZip 載入失敗'));
            document.head.appendChild(script);
        });
    }

    // Convert to RTF format
    static convertToRtf(content, title, options = {}) {
        try {
            console.log('轉換為 RTF 格式...');
            
            const documentTitle = title || '文件';
            const documentContent = content || '';
            
            // Create RTF document
            let rtfContent = '{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Times New Roman;}}';
            
            // Add title
            rtfContent += `\\f0\\fs28\\b ${DocumentConverter.escapeRtf(documentTitle)}\\b0\\par\\par`;
            
            // Add content with paragraph breaks
            const paragraphs = documentContent.split('\n\n');
            paragraphs.forEach(paragraph => {
                if (paragraph.trim()) {
                    rtfContent += `\\f0\\fs24 ${DocumentConverter.escapeRtf(paragraph.trim())}\\par\\par`;
                }
            });
            
            rtfContent += '}';
            
            const blob = new Blob([rtfContent], { 
                type: 'application/rtf;charset=utf-8' 
            });
            
            console.log('RTF 轉換完成');
            return blob;
            
        } catch (error) {
            console.error('RTF 轉換錯誤:', error);
            throw new Error('RTF 轉換失敗: ' + error.message);
        }
    }

    // Helper: Convert content to Word XML paragraphs
    static contentToWordXml(content) {
        const paragraphs = content.split('\n\n');
        return paragraphs.map(paragraph => {
            if (!paragraph.trim()) return '';
            
            return `
        <w:p>
            <w:r>
                <w:t>${DocumentConverter.escapeXml(paragraph.trim())}</w:t>
            </w:r>
        </w:p>`;
        }).join('');
    }

    // Helper: Create DOCX file structure (simplified)
    static async createDocxFile(documentXml, title) {
        // For simplicity, we'll create a Word XML document
        // This is readable by most Word processors including LibreOffice
        const wordXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">
    <w:docPr>
        <w:view w:val="print"/>
        <w:zoom w:percent="100"/>
    </w:docPr>
    <w:body>
        <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:sz w:val="28"/>
                </w:rPr>
                <w:t>${DocumentConverter.escapeXml(title)}</w:t>
            </w:r>
        </w:p>
        <w:p><w:r><w:t></w:t></w:r></w:p>
        ${documentXml.replace(/<?xml[^>]*>/, '').replace(/<w:document[^>]*>|<\/w:document>|<w:body>|<\/w:body>/g, '')}
    </w:body>
</w:wordDocument>`;

        const blob = new Blob([wordXml], { 
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document;charset=utf-8' 
        });
        
        return blob;
    }

    // Helper: Escape XML characters
    static escapeXml(text) {
        if (!text) return '';
        return text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    // Helper: Escape RTF characters
    static escapeRtf(text) {
        if (!text) return '';
        return text
            .replace(/\\/g, '\\\\')
            .replace(/{/g, '\\{')
            .replace(/}/g, '\\}')
            .replace(/\n/g, '\\par ')
            .replace(/\r/g, '');
    }

    // Convert document content to PDF format
    static async convertToPdf(content, title = '文件', options = {}) {
        try {
            console.log('🔄 開始 PDF 轉換...');
            
            // Load jsPDF if not already loaded
            if (typeof jsPDF === 'undefined') {
                await DocumentConverter.loadJsPdf();
            }
            
            // Create new PDF document - handle different jsPDF structures
            let doc;
            if (typeof jsPDF === 'function') {
                doc = new jsPDF('p', 'mm', 'a4');
            } else if (typeof jsPDF !== 'undefined' && typeof jsPDF.jsPDF === 'function') {
                doc = new jsPDF.jsPDF('p', 'mm', 'a4');
            } else if (typeof window.jspdf !== 'undefined' && typeof window.jspdf.jsPDF === 'function') {
                doc = new window.jspdf.jsPDF('p', 'mm', 'a4');
            } else {
                throw new Error('無法找到可用的 jsPDF 構造函數');
            }
            
            // Configure font and basic settings
            const pageWidth = doc.internal.pageSize.getWidth();
            const pageHeight = doc.internal.pageSize.getHeight();
            const margin = 20;
            const contentWidth = pageWidth - 2 * margin;
            let yPosition = margin;
            
            // Add title if provided
            if (title && title.trim()) {
                doc.setFontSize(18);
                doc.setFont('helvetica', 'bold');
                const titleLines = doc.splitTextToSize(title.trim(), contentWidth);
                
                titleLines.forEach((line, index) => {
                    if (yPosition > pageHeight - margin) {
                        doc.addPage();
                        yPosition = margin;
                    }
                    doc.text(line, margin, yPosition);
                    yPosition += 8;
                });
                
                yPosition += 10; // Add extra space after title
            }
            
            // Process content
            doc.setFontSize(12);
            doc.setFont('helvetica', 'normal');
            
            const processedContent = content || '(無內容)';
            const paragraphs = processedContent.split(/\n\s*\n/); // Split by double newlines
            
            paragraphs.forEach((paragraph, index) => {
                const trimmedParagraph = paragraph.trim();
                if (!trimmedParagraph) return;
                
                // Handle line breaks within paragraphs
                const lines = trimmedParagraph.split('\n');
                
                lines.forEach(line => {
                    if (!line.trim()) {
                        yPosition += 6; // Add space for empty lines
                        return;
                    }
                    
                    // Split long lines to fit page width
                    const wrappedLines = doc.splitTextToSize(line.trim(), contentWidth);
                    
                    wrappedLines.forEach(wrappedLine => {
                        // Check if we need a new page
                        if (yPosition > pageHeight - margin - 10) {
                            doc.addPage();
                            yPosition = margin;
                        }
                        
                        doc.text(wrappedLine, margin, yPosition);
                        yPosition += 6;
                    });
                });
                
                // Add space between paragraphs
                if (index < paragraphs.length - 1) {
                    yPosition += 4;
                }
            });
            
            // Generate PDF blob
            const pdfBlob = doc.output('blob');
            
            console.log(`✅ PDF 轉換完成 - 檔案大小: ${pdfBlob.size} bytes`);
            return pdfBlob;
            
        } catch (error) {
            console.error('❌ PDF 轉換失敗:', error);
            throw new Error(`PDF 轉換失敗: ${error.message}`);
        }
    }

    // Load jsPDF library dynamically
    static async loadJsPdf() {
        // Check multiple possible global names for jsPDF
        if (typeof jsPDF !== 'undefined' || 
            typeof window.jsPDF !== 'undefined' ||
            (typeof window.jspdf !== 'undefined' && window.jspdf.jsPDF)) {
            return Promise.resolve();
        }
        
        return new Promise((resolve, reject) => {
            console.log('📚 載入 jsPDF 函式庫...');
            const script = document.createElement('script');
            script.src = 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js';
            
            script.onload = () => {
                // Wait a bit for the script to initialize
                setTimeout(() => {
                    if (typeof jsPDF !== 'undefined') {
                        console.log('✅ jsPDF 載入成功 (全域變數 jsPDF)');
                        resolve();
                    } else if (typeof window.jsPDF !== 'undefined') {
                        console.log('✅ jsPDF 載入成功 (window.jsPDF)');
                        window.jsPDF = window.jsPDF;
                        resolve();
                    } else if (typeof window.jspdf !== 'undefined' && window.jspdf.jsPDF) {
                        console.log('✅ jsPDF 載入成功 (window.jspdf.jsPDF)');
                        window.jsPDF = window.jspdf.jsPDF;
                        resolve();
                    } else {
                        console.error('❌ jsPDF 載入後無法找到可用的 jsPDF 物件');
                        console.log('可用的全域變數:', Object.keys(window).filter(key => key.toLowerCase().includes('pdf')));
                        reject(new Error('jsPDF 載入失敗：無法找到 jsPDF 物件'));
                    }
                }, 100);
            };
            
            script.onerror = (error) => {
                console.error('❌ jsPDF 腳本載入失敗:', error);
                reject(new Error('無法載入 jsPDF 函式庫'));
            };
            
            document.head.appendChild(script);
        });
    }

    // Main conversion method that routes to specific converters
    static async convert(content, outputFormat, title = '', options = {}) {
        try {
            console.log(`🔄 轉換為 ${outputFormat.toUpperCase()} 格式...`);
            
            switch (outputFormat.toLowerCase()) {
                case 'pdf':
                    return await DocumentConverter.convertToPdf(content, title, options);
                case 'docx':
                    return await DocumentConverter.convertToDocx(content, title, options);
                case 'rtf':
                    return DocumentConverter.convertToRtf(content, title, options);
                case 'html':
                    return DocumentConverter.convertToHtml(content, title, options);
                case 'txt':
                    return DocumentConverter.convertToText(content, title, options);
                case 'md':
                    return DocumentConverter.convertToMarkdown(content, title, options);
                default:
                    throw new Error(`不支援的輸出格式: ${outputFormat}`);
            }
        } catch (error) {
            console.error(`❌ 轉換失敗 (${outputFormat}):`, error);
            throw error;
        }
    }

    // Convert to HTML format
    static convertToHtml(content, title = '', options = {}) {
        const documentTitle = title || '文件';
        const processedContent = content || '';
        
        // Convert paragraphs to HTML
        const paragraphs = processedContent.split(/\n\s*\n/)
            .filter(p => p.trim())
            .map(p => `<p>${DocumentConverter.escapeXml(p.trim().replace(/\n/g, '<br>'))}</p>`)
            .join('\n');
        
        const htmlContent = `<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${DocumentConverter.escapeXml(documentTitle)}</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 40px; }
        h1 { color: #333; border-bottom: 2px solid #333; padding-bottom: 10px; }
        p { margin: 16px 0; }
    </style>
</head>
<body>
    <h1>${DocumentConverter.escapeXml(documentTitle)}</h1>
    ${paragraphs}
</body>
</html>`;
        
        return new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    }

    // Convert to plain text
    static convertToText(content, title = '', options = {}) {
        const documentTitle = title || '文件';
        const processedContent = content || '';
        
        let textContent = '';
        if (title && title.trim()) {
            textContent += `${documentTitle}\n${'='.repeat(documentTitle.length)}\n\n`;
        }
        textContent += processedContent;
        
        return new Blob([textContent], { type: 'text/plain;charset=utf-8' });
    }

    // Convert to Markdown format
    static convertToMarkdown(content, title = '', options = {}) {
        const documentTitle = title || '文件';
        const processedContent = content || '';
        
        let markdownContent = '';
        if (title && title.trim()) {
            markdownContent += `# ${documentTitle}\n\n`;
        }
        
        // Convert paragraphs to markdown (simple conversion)
        const paragraphs = processedContent.split(/\n\s*\n/)
            .filter(p => p.trim())
            .map(p => p.trim().replace(/\n/g, '  \n')) // Preserve line breaks in markdown
            .join('\n\n');
        
        markdownContent += paragraphs;
        
        return new Blob([markdownContent], { type: 'text/markdown;charset=utf-8' });
    }
}

// Export for use in main application
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DocumentConverter;
} else if (typeof window !== 'undefined') {
    window.DocumentConverter = DocumentConverter;
}
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
            // Note: This would require jsPDF library
            // For now, return a placeholder implementation
            const pdfContent = `PDF 轉換功能需要 jsPDF 函式庫支援
            
標題: ${title || '無標題'}
內容長度: ${content.length} 字元

請在 HTML 中加入：
<script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js"></script>

實際實作時會將文字內容轉換為 PDF 格式。`;

            const blob = new Blob([pdfContent], { type: 'text/plain;charset=utf-8' });
            return blob;
        } catch (error) {
            throw new Error('PDF 轉換失敗: ' + error.message);
        }
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
}

// Export for use in main application
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DocumentConverter;
} else if (typeof window !== 'undefined') {
    window.DocumentConverter = DocumentConverter;
}
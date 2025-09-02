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
                scale: 2,
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
            // Try to load mammoth.js
            await DirectPDFConverter.loadMammoth();
            
            if (typeof mammoth !== 'undefined') {
                console.log('📄 使用 Mammoth.js 轉換 DOCX');
                
                const arrayBuffer = await file.arrayBuffer();
                const result = await mammoth.convertToHtml({ arrayBuffer });
                
                // Create HTML document
                const htmlDoc = `
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <style>
                        body { 
                            font-family: 'Times New Roman', serif; 
                            font-size: 12pt; 
                            line-height: 1.2; 
                            margin: 2.5cm;
                            background: white;
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
                throw new Error('Mammoth.js 未載入，無法處理 DOCX');
            }
            
        } catch (error) {
            console.warn('DOCX 專用轉換失敗，使用一般文字模式:', error);
            // Fallback to text extraction
            const textContent = await DirectPDFConverter.extractTextFromDocument(file);
            return await DirectPDFConverter.renderTextAsPDF(textContent, file.name);
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

    // Convert PPTX to PDF by extracting slide images
    static async convertPPTXToPDF(file) {
        try {
            // Load JSZip for PPTX processing
            await DirectPDFConverter.loadJSZip();
            
            const arrayBuffer = await file.arrayBuffer();
            const zip = new JSZip();
            const zipContent = await zip.loadAsync(arrayBuffer);
            
            // Look for slide images in the ZIP
            const slideImages = [];
            zipContent.forEach((relativePath, file) => {
                if (relativePath.includes('ppt/media/') && 
                    (relativePath.endsWith('.png') || relativePath.endsWith('.jpg') || relativePath.endsWith('.jpeg'))) {
                    slideImages.push({ path: relativePath, file: file });
                }
            });
            
            if (slideImages.length > 0) {
                console.log(`🖼️ 找到 ${slideImages.length} 張投影片圖片`);
                
                // Create PDF with slide images
                await DirectPDFConverter.loadJsPDF();
                const { jsPDF } = window.jspdf || window;
                const pdf = new jsPDF({
                    orientation: 'landscape', // Presentations are typically landscape
                    unit: 'mm',
                    format: 'a4'
                });
                
                let addedPages = 0;
                
                for (let i = 0; i < slideImages.length; i++) {
                    try {
                        const imageFile = slideImages[i];
                        const imageData = await imageFile.file.async('blob');
                        const imageUrl = URL.createObjectURL(imageData);
                        
                        // Add page for each slide (except first)
                        if (addedPages > 0) {
                            pdf.addPage();
                        }
                        
                        // Add image to PDF
                        pdf.addImage(imageUrl, 'JPEG', 10, 10, 277, 190); // A4 landscape dimensions
                        
                        URL.revokeObjectURL(imageUrl);
                        addedPages++;
                        
                    } catch (imgError) {
                        console.warn(`投影片 ${i + 1} 處理失敗:`, imgError);
                    }
                }
                
                if (addedPages > 0) {
                    const pdfBlob = pdf.output('blob');
                    console.log(`✅ PPTX轉PDF完成: ${addedPages} 頁`);
                    return pdfBlob;
                }
            }
            
            // Fallback: create placeholder
            return await DirectPDFConverter.createPresentationPlaceholderPDF(file);
            
        } catch (error) {
            console.warn('PPTX 解析失敗，建立預覽版本:', error);
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
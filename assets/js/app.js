class FileConverter {
    constructor() {
        this.uploadedFiles = [];
        this.convertedFiles = [];
        this.isConverting = false;
        
        this.initializeElements();
        this.bindEvents();
    }

    initializeElements() {
        // Upload elements
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.uploadBtn = document.getElementById('uploadBtn');
        
        // File list elements
        this.fileListSection = document.getElementById('fileListSection');
        this.fileList = document.getElementById('fileList');
        
        // Conversion elements
        this.conversionSection = document.getElementById('conversionSection');
        this.outputFormat = document.getElementById('outputFormat');
        this.quality = document.getElementById('quality');
        this.qualityValue = document.getElementById('qualityValue');
        this.maxWidth = document.getElementById('maxWidth');
        this.convertBtn = document.getElementById('convertBtn');
        
        // Progress elements
        this.progressSection = document.getElementById('progressSection');
        this.progressFill = document.getElementById('progressFill');
        this.progressText = document.getElementById('progressText');
        this.progressDetails = document.getElementById('progressDetails');
        
        // Download elements
        this.downloadSection = document.getElementById('downloadSection');
        this.downloadList = document.getElementById('downloadList');
        this.downloadAllBtn = document.getElementById('downloadAllBtn');
        this.resetBtn = document.getElementById('resetBtn');
    }

    bindEvents() {
        // Upload events
        this.uploadBtn.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        
        // Drag and drop events
        this.uploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
        this.uploadArea.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        this.uploadArea.addEventListener('drop', (e) => this.handleDrop(e));
        
        // Conversion events
        this.quality.addEventListener('input', (e) => {
            this.qualityValue.textContent = Math.round(e.target.value * 100) + '%';
        });
        this.outputFormat.addEventListener('change', () => this.updateConvertButton());
        this.convertBtn.addEventListener('click', () => this.startConversion());
        
        // Download events
        this.downloadAllBtn.addEventListener('click', () => this.downloadAll());
        this.resetBtn.addEventListener('click', () => this.reset());
    }

    handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadArea.classList.add('dragover');
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadArea.classList.remove('dragover');
    }

    handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadArea.classList.remove('dragover');
        
        const files = Array.from(e.dataTransfer.files);
        this.processFiles(files);
    }

    handleFileSelect(e) {
        const files = Array.from(e.target.files);
        this.processFiles(files);
    }

    processFiles(files) {
        const imageFiles = files.filter(file => this.isValidImageFile(file));
        
        if (imageFiles.length === 0) {
            alert('請選擇有效的圖片檔案 (JPG, PNG, GIF, BMP, WebP)');
            return;
        }

        if (imageFiles.length !== files.length) {
            alert(`只能處理圖片檔案，已篩選出 ${imageFiles.length} 個圖片檔案`);
        }

        this.uploadedFiles = imageFiles;
        this.displayFileList();
        this.showConversionOptions();
    }

    isValidImageFile(file) {
        const validTypes = [
            'image/jpeg',
            'image/jpg', 
            'image/png',
            'image/gif',
            'image/bmp',
            'image/webp'
        ];
        return validTypes.includes(file.type);
    }

    displayFileList() {
        this.fileList.innerHTML = '';
        
        this.uploadedFiles.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            
            // Create preview
            const preview = document.createElement('img');
            preview.className = 'file-preview';
            preview.src = URL.createObjectURL(file);
            preview.onload = () => URL.revokeObjectURL(preview.src);
            
            // File info
            const fileInfo = document.createElement('div');
            fileInfo.className = 'file-info';
            
            const fileDetails = document.createElement('div');
            fileDetails.className = 'file-details';
            fileDetails.innerHTML = `
                <h4>${file.name}</h4>
                <p>${this.formatFileSize(file.size)} | ${file.type}</p>
            `;
            
            fileInfo.appendChild(preview);
            fileInfo.appendChild(fileDetails);
            
            // Remove button
            const removeBtn = document.createElement('button');
            removeBtn.className = 'file-remove';
            removeBtn.textContent = '移除';
            removeBtn.addEventListener('click', () => this.removeFile(index));
            
            fileItem.appendChild(fileInfo);
            fileItem.appendChild(removeBtn);
            this.fileList.appendChild(fileItem);
        });
        
        this.fileListSection.style.display = 'block';
    }

    removeFile(index) {
        this.uploadedFiles.splice(index, 1);
        
        if (this.uploadedFiles.length === 0) {
            this.fileListSection.style.display = 'none';
            this.conversionSection.style.display = 'none';
        } else {
            this.displayFileList();
        }
        
        this.updateConvertButton();
    }

    showConversionOptions() {
        // Update format options based on uploaded files
        this.updateFormatOptions();
        this.conversionSection.style.display = 'block';
        this.updateConvertButton();
    }

    updateFormatOptions() {
        // Get unique input formats
        const inputFormats = [...new Set(this.uploadedFiles.map(file => {
            return file.type.split('/')[1].toLowerCase();
        }))];
        
        // Reset options
        this.outputFormat.innerHTML = '<option value="">請選擇格式</option>';
        
        // Add available output formats
        const formats = [
            { value: 'jpeg', label: 'JPEG' },
            { value: 'png', label: 'PNG' },
            { value: 'webp', label: 'WebP' }
        ];
        
        formats.forEach(format => {
            const option = document.createElement('option');
            option.value = format.value;
            option.textContent = format.label;
            this.outputFormat.appendChild(option);
        });
    }

    updateConvertButton() {
        const hasFiles = this.uploadedFiles.length > 0;
        const hasFormat = this.outputFormat.value !== '';
        
        this.convertBtn.disabled = !hasFiles || !hasFormat || this.isConverting;
    }

    async startConversion() {
        if (this.isConverting) return;
        
        this.isConverting = true;
        this.convertBtn.disabled = true;
        this.convertedFiles = [];
        
        // Show progress section
        this.progressSection.style.display = 'block';
        this.downloadSection.style.display = 'none';
        
        const totalFiles = this.uploadedFiles.length;
        let processedFiles = 0;
        
        try {
            for (let i = 0; i < this.uploadedFiles.length; i++) {
                const file = this.uploadedFiles[i];
                
                // Update progress
                this.updateProgress(processedFiles, totalFiles, `正在處理: ${file.name}`);
                
                try {
                    const convertedBlob = await this.convertImage(file);
                    const convertedFile = {
                        name: this.generateOutputFileName(file.name),
                        blob: convertedBlob,
                        originalFile: file
                    };
                    
                    this.convertedFiles.push(convertedFile);
                    processedFiles++;
                    
                } catch (error) {
                    console.error(`轉換失敗: ${file.name}`, error);
                    this.updateProgress(processedFiles, totalFiles, `轉換失敗: ${file.name}`);
                }
            }
            
            // Complete
            this.updateProgress(totalFiles, totalFiles, '轉換完成！');
            this.showDownloadSection();
            
        } catch (error) {
            console.error('轉換過程發生錯誤:', error);
            alert('轉換過程發生錯誤，請重試');
        } finally {
            this.isConverting = false;
            this.updateConvertButton();
        }
    }

    async convertImage(file) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            
            img.onload = () => {
                try {
                    // Calculate dimensions
                    let { width, height } = img;
                    const maxWidth = parseInt(this.maxWidth.value) || null;
                    
                    if (maxWidth && width > maxWidth) {
                        height = (height * maxWidth) / width;
                        width = maxWidth;
                    }
                    
                    // Set canvas size
                    canvas.width = width;
                    canvas.height = height;
                    
                    // Draw image
                    ctx.drawImage(img, 0, 0, width, height);
                    
                    // Convert to blob
                    const outputFormat = this.outputFormat.value;
                    const quality = parseFloat(this.quality.value);
                    const mimeType = `image/${outputFormat}`;
                    
                    canvas.toBlob((blob) => {
                        if (blob) {
                            resolve(blob);
                        } else {
                            reject(new Error('轉換失敗'));
                        }
                    }, mimeType, quality);
                    
                } catch (error) {
                    reject(error);
                }
            };
            
            img.onerror = () => reject(new Error('圖片載入失敗'));
            img.src = URL.createObjectURL(file);
        });
    }

    generateOutputFileName(originalName) {
        const nameWithoutExt = originalName.replace(/\.[^/.]+$/, '');
        const outputFormat = this.outputFormat.value;
        return `${nameWithoutExt}.${outputFormat}`;
    }

    updateProgress(processed, total, message) {
        const percentage = Math.round((processed / total) * 100);
        this.progressFill.style.width = percentage + '%';
        this.progressText.textContent = percentage + '%';
        this.progressDetails.textContent = message;
    }

    showDownloadSection() {
        this.downloadList.innerHTML = '';
        
        this.convertedFiles.forEach((file, index) => {
            const downloadItem = document.createElement('div');
            downloadItem.className = 'download-item';
            
            // Create preview
            const preview = document.createElement('img');
            preview.className = 'download-preview';
            preview.src = URL.createObjectURL(file.blob);
            preview.onload = () => URL.revokeObjectURL(preview.src);
            
            // Download info
            const downloadInfo = document.createElement('div');
            downloadInfo.className = 'download-info';
            
            const fileDetails = document.createElement('div');
            fileDetails.innerHTML = `
                <h4>${file.name}</h4>
                <p>${this.formatFileSize(file.blob.size)}</p>
            `;
            
            downloadInfo.appendChild(preview);
            downloadInfo.appendChild(fileDetails);
            
            // Download button
            const downloadBtn = document.createElement('button');
            downloadBtn.className = 'download-btn';
            downloadBtn.textContent = '下載';
            downloadBtn.addEventListener('click', () => this.downloadFile(file));
            
            downloadItem.appendChild(downloadInfo);
            downloadItem.appendChild(downloadBtn);
            this.downloadList.appendChild(downloadItem);
        });
        
        this.downloadSection.style.display = 'block';
    }

    downloadFile(file) {
        const url = URL.createObjectURL(file.blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = file.name;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    async downloadAll() {
        if (this.convertedFiles.length === 1) {
            this.downloadFile(this.convertedFiles[0]);
            return;
        }
        
        // Create ZIP for multiple files (if we had JSZip)
        // For now, download individually
        for (const file of this.convertedFiles) {
            await new Promise(resolve => {
                setTimeout(() => {
                    this.downloadFile(file);
                    resolve();
                }, 500);
            });
        }
    }

    reset() {
        this.uploadedFiles = [];
        this.convertedFiles = [];
        this.isConverting = false;
        
        // Reset form
        this.fileInput.value = '';
        this.outputFormat.value = '';
        this.quality.value = 0.8;
        this.qualityValue.textContent = '80%';
        this.maxWidth.value = '';
        
        // Hide sections
        this.fileListSection.style.display = 'none';
        this.conversionSection.style.display = 'none';
        this.progressSection.style.display = 'none';
        this.downloadSection.style.display = 'none';
        
        this.updateConvertButton();
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    new FileConverter();
});
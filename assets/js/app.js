class FileConverter {
    constructor() {
        this.uploadedFiles = [];
        this.convertedFiles = [];
        this.isConverting = false;
        this.currentFileType = 'image'; // image, document, spreadsheet, presentation
        
        this.initializeElements();
        this.bindEvents();
        this.setupFileTypeSelector();
    }

    initializeElements() {
        // File type selector
        this.typeTabs = document.querySelectorAll('.tab-btn');
        
        // Upload elements
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.uploadBtn = document.getElementById('uploadBtn');
        this.uploadIcon = document.getElementById('uploadIcon');
        this.uploadTitle = document.getElementById('uploadTitle');
        this.uploadDescription = document.getElementById('uploadDescription');
        
        // File list elements
        this.fileListSection = document.getElementById('fileListSection');
        this.fileList = document.getElementById('fileList');
        
        // Conversion elements
        this.conversionSection = document.getElementById('conversionSection');
        this.conversionOptions = document.getElementById('conversionOptions');
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
        this.convertBtn.addEventListener('click', () => this.startConversion());
        
        // Download events
        this.downloadAllBtn.addEventListener('click', () => this.downloadAll());
        this.resetBtn.addEventListener('click', () => this.reset());
    }

    setupFileTypeSelector() {
        this.typeTabs.forEach(tab => {
            tab.addEventListener('click', (e) => {
                const fileType = e.target.dataset.type;
                this.switchFileType(fileType);
            });
        });
    }

    switchFileType(fileType) {
        console.log('Switching file type from', this.currentFileType, 'to', fileType);
        this.currentFileType = fileType;
        
        // Update active tab
        this.typeTabs.forEach(tab => {
            tab.classList.remove('active');
            if (tab.dataset.type === fileType) {
                tab.classList.add('active');
                console.log('Activated tab for:', fileType);
            }
        });
        
        // Update UI based on file type
        this.updateUploadInterface(fileType);
        
        // Clear current files if switching type
        this.reset();
        
        console.log('File type switched successfully to:', this.currentFileType);
    }

    updateUploadInterface(fileType) {
        const configs = {
            image: {
                icon: '🖼️',
                title: '拖拉圖片檔案到此處或點擊上傳',
                description: '支援 JPG、PNG、GIF、BMP、WebP 格式',
                accept: 'image/*'
            },
            document: {
                icon: '📄',
                title: '拖拉文書檔案到此處或點擊上傳',
                description: '支援 DOCX、TXT、HTML、MD、RTF 格式',
                accept: '.docx,.doc,.txt,.html,.htm,.md,.rtf'
            },
            spreadsheet: {
                icon: '📊',
                title: '拖拉表單檔案到此處或點擊上傳',
                description: '支援 XLSX、CSV、TSV 格式',
                accept: '.xlsx,.xls,.csv,.tsv'
            },
            presentation: {
                icon: '🎯',
                title: '拖拉簡報檔案到此處或點擊上傳',
                description: '支援 PPTX、PDF、HTML 格式',
                accept: '.pptx,.ppt,.pdf,.html,.htm'
            }
        };
        
        const config = configs[fileType];
        if (config) {
            this.uploadIcon.textContent = config.icon;
            this.uploadTitle.textContent = config.title;
            this.uploadDescription.textContent = config.description;
            this.fileInput.setAttribute('accept', config.accept);
        }
    }

    isValidFile(file) {
        switch (this.currentFileType) {
            case 'image':
                return ImageConverter.isValidImageFile(file);
            case 'document':
                return DocumentConverter.isValidDocumentFile(file);
            case 'spreadsheet':
                return SpreadsheetConverter.isValidSpreadsheetFile(file);
            case 'presentation':
                return PresentationConverter.isValidPresentationFile(file);
            default:
                return false;
        }
    }

    getFileTypeErrorMessage() {
        const messages = {
            image: '請選擇有效的圖片檔案 (JPG, PNG, GIF, BMP, WebP)',
            document: '請選擇有效的文書檔案 (DOCX, TXT, HTML, MD, RTF)',
            spreadsheet: '請選擇有效的表單檔案 (XLSX, CSV, TSV)',
            presentation: '請選擇有效的簡報檔案 (PPTX, HTML)'
        };
        return messages[this.currentFileType] || '請選擇有效的檔案';
    }

    getFileTypeName() {
        const names = {
            image: '圖片',
            document: '文書',
            spreadsheet: '表單',
            presentation: '簡報'
        };
        return names[this.currentFileType] || '檔案';
    }

    createErrorReport(file, error) {
        const report = `檔案轉換錯誤報告
${'='.repeat(40)}

錯誤時間: ${new Date().toLocaleString()}
檔案名稱: ${file.name}
檔案大小: ${this.formatFileSize(file.size)}
檔案類型: ${file.type}
轉換類型: ${this.currentFileType}
目標格式: ${document.getElementById('outputFormat')?.value || '未設定'}

錯誤訊息:
${error.message}

可能原因:
1. 檔案格式不支援或損壞
2. 檔案內容異常
3. 所需的第三方函式庫未加載
4. 轉換參數設定錯誤
5. 瀏覽器相容性問題

建議解決方案:
1. 檢查檔案是否可正常開啟
2. 嘗試其他輸出格式
3. 減小檔案大小後再試
4. 更新瀏覽器版本
5. 使用專業轉換軟體

技術支援:
如果問題持續，請將此錯誤報告提供給技術人員。`;

        return new Blob([report], { type: 'text/plain;charset=utf-8' });
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
        const validFiles = files.filter(file => this.isValidFile(file));
        
        if (validFiles.length === 0) {
            alert(this.getFileTypeErrorMessage());
            return;
        }

        if (validFiles.length !== files.length) {
            const fileTypeName = this.getFileTypeName();
            alert(`只能處理${fileTypeName}檔案，已篩選出 ${validFiles.length} 個有效檔案`);
        }

        this.uploadedFiles = validFiles;
        this.displayFileList();
        this.showConversionOptions();
    }

    displayFileList() {
        console.log('Displaying file list. Current file type:', this.currentFileType, 'Files count:', this.uploadedFiles.length);
        this.fileList.innerHTML = '';
        
        this.uploadedFiles.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            
            // Create preview based on file type
            const preview = this.createFilePreview(file);
            
            // File info
            const fileInfo = document.createElement('div');
            fileInfo.className = 'file-info';
            
            const fileDetails = document.createElement('div');
            fileDetails.className = 'file-details';
            fileDetails.innerHTML = `
                <h4>${file.name}</h4>
                <p>${this.formatFileSize(file.size)} | ${this.getFileTypeName()}</p>
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
        // Create conversion options based on file type
        this.createConversionOptions();
        this.conversionSection.style.display = 'block';
        this.updateConvertButton();
    }

    createConversionOptions() {
        this.conversionOptions.innerHTML = '';
        
        // Create options based on file type
        const optionConfigs = {
            image: this.createImageOptions(),
            document: this.createDocumentOptions(),
            spreadsheet: this.createSpreadsheetOptions(),
            presentation: this.createPresentationOptions()
        };
        
        const config = optionConfigs[this.currentFileType];
        if (config) {
            this.conversionOptions.appendChild(config);
        }
    }

    createImageOptions() {
        const container = document.createElement('div');
        
        // Output format
        const formatGroup = this.createOptionGroup('輸出格式', 'select', 'outputFormat', [
            { value: '', label: '請選擇格式' },
            { value: 'jpeg', label: 'JPEG' },
            { value: 'png', label: 'PNG' },
            { value: 'webp', label: 'WebP' }
        ]);
        
        // Quality setting
        const qualityGroup = this.createOptionGroup('品質設定', 'range', 'quality', null, {
            min: 0.1, max: 1, step: 0.1, value: 0.8
        });
        
        // Max width
        const widthGroup = this.createOptionGroup('最大寬度 (像素)', 'number', 'maxWidth', null, {
            placeholder: '不限制', min: 1
        });
        
        container.appendChild(formatGroup);
        container.appendChild(qualityGroup);
        container.appendChild(widthGroup);
        
        return container;
    }

    createDocumentOptions() {
        const container = document.createElement('div');
        
        // Output format
        const formatGroup = this.createOptionGroup('輸出格式', 'select', 'outputFormat', [
            { value: '', label: '請選擇格式' },
            { value: 'txt', label: '純文字 (TXT)' },
            { value: 'docx', label: 'Word 文件 (DOCX)' },
            { value: 'html', label: 'HTML 網頁' },
            { value: 'md', label: 'Markdown' },
            { value: 'pdf', label: 'PDF 文件' },
            { value: 'rtf', label: 'RTF 文件' }
        ]);
        
        // Compression level
        const compressionGroup = this.createOptionGroup('內容壓縮', 'select', 'compression', [
            { value: 'standard', label: '標準' },
            { value: 'low', label: '輕度壓縮' },
            { value: 'medium', label: '中度壓縮' },
            { value: 'high', label: '高度壓縮' }
        ]);
        
        container.appendChild(formatGroup);
        container.appendChild(compressionGroup);
        
        return container;
    }

    createSpreadsheetOptions() {
        const container = document.createElement('div');
        
        // Output format
        const formatGroup = this.createOptionGroup('輸出格式', 'select', 'outputFormat', [
            { value: '', label: '請選擇格式' },
            { value: 'csv', label: 'CSV 逗號分隔' },
            { value: 'xlsx', label: 'Excel 工作簿 (XLSX)' },
            { value: 'json', label: 'JSON 格式' },
            { value: 'html', label: 'HTML 表格' },
            { value: 'txt', label: '純文字' }
        ]);
        
        // Include headers
        const headersGroup = this.createOptionGroup('包含標題行', 'checkbox', 'includeHeaders', null, {
            checked: true
        });
        
        container.appendChild(formatGroup);
        container.appendChild(headersGroup);
        
        return container;
    }

    createPresentationOptions() {
        const container = document.createElement('div');
        
        // Output format
        const formatGroup = this.createOptionGroup('輸出格式', 'select', 'outputFormat', [
            { value: '', label: '請選擇格式' },
            { value: 'pptx', label: 'PowerPoint (PPTX)' },
            { value: 'html', label: 'HTML 簡報' },
            { value: 'txt', label: '純文字' },
            { value: 'md', label: 'Markdown' },
            { value: 'pdf', label: 'PDF 文件' }
        ]);
        
        // Include notes
        const notesGroup = this.createOptionGroup('包含備註', 'checkbox', 'includeNotes', null, {
            checked: true
        });
        
        container.appendChild(formatGroup);
        container.appendChild(notesGroup);
        
        return container;
    }

    createOptionGroup(label, type, id, options = null, attrs = {}) {
        const group = document.createElement('div');
        group.className = 'option-group';
        
        const labelEl = document.createElement('label');
        labelEl.textContent = label + '：';
        labelEl.setAttribute('for', id);
        
        let input;
        if (type === 'select') {
            input = document.createElement('select');
            if (options) {
                options.forEach(option => {
                    const optionEl = document.createElement('option');
                    optionEl.value = option.value;
                    optionEl.textContent = option.label;
                    input.appendChild(optionEl);
                });
            }
        } else {
            input = document.createElement('input');
            input.type = type;
        }
        
        input.id = id;
        Object.keys(attrs).forEach(attr => {
            if (attr === 'checked') {
                input.checked = attrs[attr];
            } else {
                input.setAttribute(attr, attrs[attr]);
            }
        });
        
        // Add event listeners
        if (type === 'range') {
            const valueDisplay = document.createElement('span');
            valueDisplay.id = id + 'Value';
            valueDisplay.textContent = Math.round(attrs.value * 100) + '%';
            
            input.addEventListener('input', (e) => {
                valueDisplay.textContent = Math.round(e.target.value * 100) + '%';
            });
            
            group.appendChild(labelEl);
            group.appendChild(input);
            group.appendChild(valueDisplay);
        } else {
            group.appendChild(labelEl);
            group.appendChild(input);
        }
        
        if (id === 'outputFormat') {
            input.addEventListener('change', () => this.updateConvertButton());
        }
        
        return group;
    }

    updateConvertButton() {
        const hasFiles = this.uploadedFiles.length > 0;
        const outputFormat = document.getElementById('outputFormat');
        const hasFormat = outputFormat && outputFormat.value !== '';
        
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
                    const convertedBlob = await this.convertFile(file);
                    const convertedFile = {
                        name: this.generateOutputFileName(file.name),
                        blob: convertedBlob,
                        originalFile: file
                    };
                    
                    this.convertedFiles.push(convertedFile);
                    processedFiles++;
                    
                } catch (error) {
                    console.error(`轉換失敗: ${file.name}`, error);
                    
                    // Create error report file
                    const errorReport = this.createErrorReport(file, error);
                    const errorFile = {
                        name: `錯誤報告_${file.name.replace(/\.[^/.]+$/, '')}.txt`,
                        blob: errorReport,
                        originalFile: file,
                        isError: true
                    };
                    
                    this.convertedFiles.push(errorFile);
                    this.updateProgress(processedFiles + 1, totalFiles, `轉換失敗: ${file.name} (已生成錯誤報告)`);
                    processedFiles++;
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

    async convertFile(file) {
        const outputFormat = document.getElementById('outputFormat').value;
        
        switch (this.currentFileType) {
            case 'image':
                return await this.convertImageFile(file, outputFormat);
            case 'document':
                return await this.convertDocumentFile(file, outputFormat);
            case 'spreadsheet':
                return await this.convertSpreadsheetFile(file, outputFormat);
            case 'presentation':
                return await this.convertPresentationFile(file, outputFormat);
            default:
                throw new Error(`不支援的檔案類型: ${this.currentFileType}`);
        }
    }

    async convertImageFile(file, outputFormat) {
        const maxWidth = document.getElementById('maxWidth')?.value;
        const quality = document.getElementById('quality')?.value || 0.8;
        
        return new Promise((resolve, reject) => {
            const img = new Image();
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            
            img.onload = () => {
                try {
                    // Calculate dimensions
                    let { width, height } = img;
                    const maxW = parseInt(maxWidth) || null;
                    
                    if (maxW && width > maxW) {
                        height = (height * maxW) / width;
                        width = maxW;
                    }
                    
                    // Set canvas size
                    canvas.width = width;
                    canvas.height = height;
                    
                    // Draw image
                    ctx.drawImage(img, 0, 0, width, height);
                    
                    // Convert to blob
                    const mimeType = `image/${outputFormat}`;
                    
                    canvas.toBlob((blob) => {
                        if (blob) {
                            resolve(blob);
                        } else {
                            reject(new Error('圖片轉換失敗'));
                        }
                    }, mimeType, parseFloat(quality));
                    
                } catch (error) {
                    reject(error);
                }
            };
            
            img.onerror = () => reject(new Error('圖片載入失敗'));
            img.src = URL.createObjectURL(file);
        });
    }

    async convertDocumentFile(file, outputFormat) {
        try {
            // Use direct PDF conversion for PDF output
            if (outputFormat.toLowerCase() === 'pdf') {
                console.log('📄 使用直接檔案轉PDF功能');
                return await DocumentConverter.convertToPdf(file, file.name, {});
            }
            
            // For other formats, extract content first
            const extractedContent = await DocumentConverter.extractTextContent(file);
            const compression = document.getElementById('compression')?.value || 'standard';
            
            // Apply compression if needed
            if (compression !== 'standard') {
                extractedContent.content = DocumentConverter.compressContent(
                    extractedContent.content, compression
                );
            }
            
            return await DocumentConverter.convertToFormat(extractedContent, outputFormat, {
                compression
            });
        } catch (error) {
            throw new Error(`文書轉換失敗: ${error.message}`);
        }
    }

    async convertSpreadsheetFile(file, outputFormat) {
        try {
            const parsedData = await SpreadsheetConverter.parseSpreadsheetData(file);
            // Use ?? (not ||) so an unchecked box (checked === false) is respected;
            // only fall back to true when the element itself is missing (undefined).
            const includeHeaders = document.getElementById('includeHeaders')?.checked ?? true;
            
            return await SpreadsheetConverter.convertToFormat(parsedData, outputFormat, {
                includeHeaders
            });
        } catch (error) {
            throw new Error(`表單轉換失敗: ${error.message}`);
        }
    }

    async convertPresentationFile(file, outputFormat) {
        try {
            const title = file.name.replace(/\.[^/.]+$/, '');
            
            // Use direct PDF conversion for PDF output
            if (outputFormat.toLowerCase() === 'pdf') {
                console.log('📊 使用直接檔案轉PDF功能');
                return await PresentationConverter.convertToPdf(file, title, {});
            }
            
            // For other formats, extract content first
            const presentationData = await PresentationConverter.extractPresentationContent(file);
            // Use ?? (not ||) so an unchecked box (checked === false) is respected;
            // only fall back to true when the element itself is missing (undefined).
            const includeNotes = document.getElementById('includeNotes')?.checked ?? true;
            
            let blob;
            switch (outputFormat.toLowerCase()) {
                case 'html':
                    blob = PresentationConverter.convertToHtml(presentationData, title, { includeNotes });
                    break;
                case 'txt':
                case 'text':
                    blob = PresentationConverter.convertToText(presentationData, title, { includeNotes });
                    break;
                default:
                    throw new Error(`不支援的輸出格式: ${outputFormat}`);
            }
            
            return blob;
        } catch (error) {
            throw new Error(`簡報轉換失敗: ${error.message}`);
        }
    }

    generateOutputFileName(originalName) {
        const nameWithoutExt = originalName.replace(/\.[^/.]+$/, '');
        const outputFormat = document.getElementById('outputFormat').value;
        
        // Handle special cases
        if (outputFormat === 'html') {
            return `${nameWithoutExt}.html`;
        } else if (outputFormat === 'json') {
            return `${nameWithoutExt}.json`;
        } else if (outputFormat === 'md') {
            return `${nameWithoutExt}.md`;
        }
        
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
            
            // Create preview based on file type
            const preview = this.createDownloadPreview(file);
            
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
        
        // Clear dynamic options
        if (this.conversionOptions) {
            this.conversionOptions.innerHTML = '';
        }
        
        // Hide sections
        this.fileListSection.style.display = 'none';
        this.conversionSection.style.display = 'none';
        this.progressSection.style.display = 'none';
        this.downloadSection.style.display = 'none';
        
        this.updateConvertButton();
    }

    createFilePreview(file) {
        console.log('Creating file preview for:', file.name, 'Type:', this.currentFileType);
        
        const preview = document.createElement('div');
        preview.className = 'file-preview';
        
        // Only show thumbnails for images and videos
        if (this.currentFileType === 'image' || this.currentFileType === 'video') {
            if (this.currentFileType === 'image') {
                const img = document.createElement('img');
                img.style.width = '100%';
                img.style.height = '100%';
                img.style.objectFit = 'cover';
                img.style.borderRadius = '8px';
                img.src = URL.createObjectURL(file);
                img.onload = () => URL.revokeObjectURL(img.src);
                preview.appendChild(img);
                console.log('Created image preview');
            }
        } else {
            // For documents, spreadsheets, presentations - show type icon
            const icons = {
                document: '📄',
                spreadsheet: '📊',
                presentation: '🎯',
                audio: '🎧'
            };
            const icon = icons[this.currentFileType] || '📁';
            preview.textContent = icon;
            preview.classList.add('icon-preview');
            
            // Apply styles to ensure proper display
            preview.style.display = 'flex';
            preview.style.alignItems = 'center';
            preview.style.justifyContent = 'center';
            preview.style.fontSize = '2rem';
            preview.style.backgroundColor = '#f8f9fa';
            
            console.log('Created icon preview with icon:', icon, 'for type:', this.currentFileType);
        }
        
        return preview;
    }

    createDownloadPreview(file) {
        // Get file extension to determine icon
        const extension = file.name.toLowerCase().split('.').pop();
        
        // Check if it's an image file that can be previewed
        const imageExtensions = ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp'];
        
        if (imageExtensions.includes(extension)) {
            // Show actual image preview for image files
            const preview = document.createElement('img');
            preview.className = 'download-preview';
            preview.src = URL.createObjectURL(file.blob);
            preview.onload = () => URL.revokeObjectURL(preview.src);
            return preview;
        } else {
            // Show file type icon for other files
            const preview = document.createElement('div');
            preview.className = 'download-preview file-preview icon-preview';
            
            const icons = {
                // Document formats
                'pdf': '📄', 'doc': '📄', 'docx': '📄', 'txt': '📄', 'rtf': '📄', 'md': '📄',
                // Spreadsheet formats  
                'xlsx': '📊', 'xls': '📊', 'csv': '📊', 'tsv': '📊',
                // Presentation formats
                'pptx': '🎯', 'ppt': '🎯', 'html': '🎯',
                // Other formats
                'json': '📋', 'xml': '📋'
            };
            
            const icon = icons[extension] || '📁';
            preview.textContent = icon;
            preview.style.display = 'flex';
            preview.style.alignItems = 'center';
            preview.style.justifyContent = 'center';
            preview.style.fontSize = '2rem';
            preview.style.backgroundColor = '#f8f9fa';
            preview.style.border = '1px solid #ddd';
            preview.style.borderRadius = '8px';
            preview.style.width = '50px';
            preview.style.height = '50px';
            
            return preview;
        }
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
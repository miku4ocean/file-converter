// Image conversion utilities and enhancements

class ImageConverter {
    constructor() {
        this.supportedFormats = {
            input: ['jpeg', 'jpg', 'png', 'gif', 'bmp', 'webp'],
            output: ['jpeg', 'png', 'webp']
        };
    }

    // Advanced image processing functions
    static async resizeImage(file, maxWidth, maxHeight, quality = 0.8) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');

            img.onload = () => {
                try {
                    let { width, height } = img;

                    // Calculate new dimensions
                    if (maxWidth && width > maxWidth) {
                        height = (height * maxWidth) / width;
                        width = maxWidth;
                    }
                    
                    if (maxHeight && height > maxHeight) {
                        width = (width * maxHeight) / height;
                        height = maxHeight;
                    }

                    // Set canvas size
                    canvas.width = width;
                    canvas.height = height;

                    // Enable image smoothing for better quality
                    ctx.imageSmoothingEnabled = true;
                    ctx.imageSmoothingQuality = 'high';

                    // Draw image
                    ctx.drawImage(img, 0, 0, width, height);

                    // Convert to blob
                    canvas.toBlob((blob) => {
                        if (blob) {
                            resolve(blob);
                        } else {
                            reject(new Error('圖片處理失敗'));
                        }
                    }, file.type, quality);

                } catch (error) {
                    reject(error);
                }
            };

            img.onerror = () => reject(new Error('圖片載入失敗'));
            img.src = URL.createObjectURL(file);
        });
    }

    // Convert image format with advanced options
    static async convertFormat(file, outputFormat, options = {}) {
        const {
            quality = 0.8,
            maxWidth = null,
            maxHeight = null,
            backgroundColor = null
        } = options;

        return new Promise((resolve, reject) => {
            const img = new Image();
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');

            img.onload = () => {
                try {
                    let { width, height } = img;

                    // Calculate dimensions
                    if (maxWidth && width > maxWidth) {
                        height = (height * maxWidth) / width;
                        width = maxWidth;
                    }
                    
                    if (maxHeight && height > maxHeight) {
                        width = (width * maxHeight) / height;
                        height = maxHeight;
                    }

                    // Set canvas size
                    canvas.width = width;
                    canvas.height = height;

                    // Set background color for formats that don't support transparency
                    if (backgroundColor && (outputFormat === 'jpeg' || outputFormat === 'jpg')) {
                        ctx.fillStyle = backgroundColor;
                        ctx.fillRect(0, 0, width, height);
                    }

                    // Enable high-quality rendering
                    ctx.imageSmoothingEnabled = true;
                    ctx.imageSmoothingQuality = 'high';

                    // Draw image
                    ctx.drawImage(img, 0, 0, width, height);

                    // Convert to blob
                    const mimeType = `image/${outputFormat}`;
                    canvas.toBlob((blob) => {
                        if (blob) {
                            resolve(blob);
                        } else {
                            reject(new Error('格式轉換失敗'));
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

    // Get image metadata
    static async getImageInfo(file) {
        return new Promise((resolve, reject) => {
            const img = new Image();

            img.onload = () => {
                const info = {
                    width: img.width,
                    height: img.height,
                    aspectRatio: img.width / img.height,
                    size: file.size,
                    type: file.type,
                    name: file.name,
                    lastModified: file.lastModified
                };
                resolve(info);
            };

            img.onerror = () => reject(new Error('無法讀取圖片資訊'));
            img.src = URL.createObjectURL(file);
        });
    }

    // Compress image while maintaining quality
    static async compressImage(file, targetSizeKB, maxIterations = 10) {
        let quality = 0.8;
        let iteration = 0;
        
        while (iteration < maxIterations) {
            const compressed = await ImageConverter.convertFormat(file, 'jpeg', { quality });
            
            if (compressed.size <= targetSizeKB * 1024) {
                return compressed;
            }
            
            quality -= 0.1;
            if (quality <= 0.1) break;
            iteration++;
        }
        
        // Return best effort if target size not achieved
        return ImageConverter.convertFormat(file, 'jpeg', { quality: 0.1 });
    }

    // Batch process images
    static async batchProcess(files, processFunction, onProgress = null) {
        const results = [];
        const total = files.length;
        
        for (let i = 0; i < files.length; i++) {
            try {
                const result = await processFunction(files[i], i);
                results.push({ success: true, result, file: files[i] });
                
                if (onProgress) {
                    onProgress(i + 1, total, files[i].name);
                }
            } catch (error) {
                results.push({ success: false, error, file: files[i] });
                
                if (onProgress) {
                    onProgress(i + 1, total, `錯誤: ${files[i].name}`);
                }
            }
        }
        
        return results;
    }

    // Validate image file
    static isValidImageFile(file) {
        const validTypes = [
            'image/jpeg',
            'image/jpg',
            'image/png',
            'image/gif',
            'image/bmp',
            'image/webp',
            'image/svg+xml'
        ];
        
        return validTypes.includes(file.type) && file.size > 0;
    }

    // Get optimal output format suggestion
    static suggestOutputFormat(inputFormat, hasTransparency = false) {
        const format = inputFormat.toLowerCase();
        
        if (hasTransparency) {
            return 'png'; // Preserve transparency
        }
        
        switch (format) {
            case 'png':
            case 'bmp':
                return 'jpeg'; // Convert to JPEG for smaller size
            case 'gif':
                return 'png'; // Preserve quality
            case 'webp':
                return 'webp'; // Keep modern format
            default:
                return 'jpeg';
        }
    }
}

// Export for use in main application
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ImageConverter;
} else if (typeof window !== 'undefined') {
    window.ImageConverter = ImageConverter;
}
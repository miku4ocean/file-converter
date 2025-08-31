// External library loader for advanced conversion features

class LibraryLoader {
    constructor() {
        this.loadedLibraries = new Set();
        this.loadingPromises = new Map();
    }

    // Available CDN libraries
    static libraries = {
        mammoth: {
            url: 'https://cdn.jsdelivr.net/npm/mammoth@1.4.2/mammoth.browser.min.js',
            globalName: 'mammoth',
            description: 'DOCX 文件解析'
        },
        sheetjs: {
            url: 'https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js',
            globalName: 'XLSX',
            description: 'Excel 文件處理'
        },
        jspdf: {
            url: 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',
            globalName: 'jsPDF',
            description: 'PDF 文件生成'
        },
        jszip: {
            url: 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
            globalName: 'JSZip',
            description: 'ZIP 檔案處理 (用於 DOCX/XLSX 創建)'
        },
        html2canvas: {
            url: 'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js',
            globalName: 'html2canvas',
            description: 'HTML 轉圖片'
        },
        pptxgenjs: {
            url: 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js',
            globalName: 'PptxGenJS',
            description: 'PowerPoint 文件生成'
        },
        pdfjslib: {
            url: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js',
            globalName: 'pdfjsLib',
            description: 'PDF 文件解析'
        }
    };

    // Load a library dynamically
    async loadLibrary(libraryName) {
        if (this.loadedLibraries.has(libraryName)) {
            return true; // Already loaded
        }

        if (this.loadingPromises.has(libraryName)) {
            return await this.loadingPromises.get(libraryName); // Currently loading
        }

        const library = LibraryLoader.libraries[libraryName];
        if (!library) {
            throw new Error(`未知的函式庫: ${libraryName}`);
        }

        // Check if already available globally
        if (window[library.globalName]) {
            this.loadedLibraries.add(libraryName);
            return true;
        }

        // Load the library
        const loadPromise = this.loadScript(library.url, library.globalName);
        this.loadingPromises.set(libraryName, loadPromise);

        try {
            await loadPromise;
            this.loadedLibraries.add(libraryName);
            this.loadingPromises.delete(libraryName);
            return true;
        } catch (error) {
            this.loadingPromises.delete(libraryName);
            throw new Error(`載入 ${library.description} 函式庫失敗: ${error.message}`);
        }
    }

    // Load multiple libraries
    async loadLibraries(libraryNames) {
        const results = await Promise.allSettled(
            libraryNames.map(name => this.loadLibrary(name))
        );

        const failed = [];
        results.forEach((result, index) => {
            if (result.status === 'rejected') {
                failed.push({
                    library: libraryNames[index],
                    error: result.reason
                });
            }
        });

        if (failed.length > 0) {
            const errorMessages = failed.map(f => `${f.library}: ${f.error.message}`).join('\n');
            throw new Error(`部分函式庫載入失敗:\n${errorMessages}`);
        }

        return true;
    }

    // Load script tag
    loadScript(url, globalName) {
        return new Promise((resolve, reject) => {
            // Check if script already exists
            const existingScript = document.querySelector(`script[src="${url}"]`);
            if (existingScript) {
                if (window[globalName]) {
                    resolve();
                    return;
                }
            }

            const script = document.createElement('script');
            script.src = url;
            script.async = true;
            
            script.onload = () => {
                // Verify the library is available
                if (window[globalName]) {
                    resolve();
                } else {
                    reject(new Error(`函式庫載入後無法找到全域物件: ${globalName}`));
                }
            };
            
            script.onerror = () => {
                reject(new Error(`網路錯誤: 無法載入 ${url}`));
            };
            
            // Add timeout
            setTimeout(() => {
                if (!window[globalName]) {
                    reject(new Error(`載入逾時: ${url}`));
                }
            }, 10000); // 10 second timeout
            
            document.head.appendChild(script);
        });
    }

    // Check if library is available
    isLibraryLoaded(libraryName) {
        return this.loadedLibraries.has(libraryName);
    }

    // Get library status
    getLibraryStatus() {
        const status = {};
        Object.keys(LibraryLoader.libraries).forEach(name => {
            const library = LibraryLoader.libraries[name];
            status[name] = {
                loaded: this.loadedLibraries.has(name),
                available: !!window[library.globalName],
                description: library.description
            };
        });
        return status;
    }

    // Show library loading progress
    showLoadingProgress(libraryNames, onProgress = null) {
        return new Promise(async (resolve, reject) => {
            try {
                for (let i = 0; i < libraryNames.length; i++) {
                    const libraryName = libraryNames[i];
                    const library = LibraryLoader.libraries[libraryName];
                    
                    if (onProgress) {
                        onProgress(i, libraryNames.length, `載入 ${library.description}...`);
                    }
                    
                    await this.loadLibrary(libraryName);
                    
                    if (onProgress) {
                        onProgress(i + 1, libraryNames.length, `${library.description} 載入完成`);
                    }
                }
                
                resolve();
            } catch (error) {
                reject(error);
            }
        });
    }
}

// Global instance
window.libLoader = new LibraryLoader();
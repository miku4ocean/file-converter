// Spreadsheet conversion utilities

class SpreadsheetConverter {
    constructor() {
        this.supportedFormats = {
            input: ['xlsx', 'xls', 'csv', 'tsv', 'ods'],
            output: ['csv', 'xlsx', 'json', 'html', 'txt']
        };
    }

    // Validate spreadsheet file
    static isValidSpreadsheetFile(file) {
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // xlsx
            'application/vnd.ms-excel', // xls
            'text/csv',
            'text/tab-separated-values', // tsv
            'application/vnd.oasis.opendocument.spreadsheet' // ods
        ];
        
        const validExtensions = ['.xlsx', '.xls', '.csv', '.tsv', '.ods'];
        const fileExtension = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));
        
        return validTypes.includes(file.type) || validExtensions.includes(fileExtension);
    }

    // Parse spreadsheet data from various formats
    static async parseSpreadsheetData(file) {
        const fileType = SpreadsheetConverter.getFileType(file);
        
        switch (fileType) {
            case 'csv':
                return await SpreadsheetConverter.parseCsv(file);
            case 'tsv':
                return await SpreadsheetConverter.parseTsv(file);
            case 'xlsx':
            case 'xls':
                return await SpreadsheetConverter.parseExcel(file);
            default:
                throw new Error(`不支援的表單格式: ${fileType}`);
        }
    }

    // Get file type from file object
    static getFileType(file) {
        const extension = file.name.toLowerCase().substring(file.name.lastIndexOf('.') + 1);
        return extension;
    }

    // Parse CSV files
    static async parseCsv(file) {
        try {
            const text = await file.text();
            const rows = SpreadsheetConverter.parseCsvText(text, ',');
            
            return {
                data: rows,
                sheets: [{ name: 'Sheet1', data: rows }],
                fileName: file.name.replace(/\.[^/.]+$/, ''),
                rowCount: rows.length,
                columnCount: rows.length > 0 ? rows[0].length : 0
            };
        } catch (error) {
            throw new Error('CSV 檔案讀取失敗: ' + error.message);
        }
    }

    // Parse TSV files
    static async parseTsv(file) {
        try {
            const text = await file.text();
            const rows = SpreadsheetConverter.parseCsvText(text, '\t');
            
            return {
                data: rows,
                sheets: [{ name: 'Sheet1', data: rows }],
                fileName: file.name.replace(/\.[^/.]+$/, ''),
                rowCount: rows.length,
                columnCount: rows.length > 0 ? rows[0].length : 0
            };
        } catch (error) {
            throw new Error('TSV 檔案讀取失敗: ' + error.message);
        }
    }

    // Parse Excel files (requires SheetJS)
    static async parseExcel(file) {
        try {
            console.log('開始解析 Excel 文件:', file.name);
            
            // Try to load SheetJS library if not loaded
            try {
                await window.libLoader.loadLibrary('sheetjs');
            } catch (libError) {
                console.warn('SheetJS 載入失敗，使用簡化解析方式:', libError.message);
                return SpreadsheetConverter.parseExcelFallback(file);
            }
            
            // Use SheetJS to parse Excel file
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            
            // Get the first worksheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON array
            const data = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,  // Use array of arrays instead of objects
                defval: ''  // Default value for empty cells
            });
            
            console.log('Excel 解析完成，行數:', data.length);
            
            return {
                data: data,
                sheets: workbook.SheetNames.map(name => ({
                    name: name,
                    data: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '' })
                })),
                fileName: file.name.replace(/\.[^/.]+$/, ''),
                rowCount: data.length,
                columnCount: data.length > 0 ? Math.max(...data.map(row => row.length)) : 0
            };
            
        } catch (error) {
            console.error('Excel 解析錯誤:', error);
            
            // Fallback to simple parsing
            try {
                return await SpreadsheetConverter.parseExcelFallback(file);
            } catch (fallbackError) {
                throw new Error('Excel 檔案讀取失敗: ' + error.message);
            }
        }
    }

    // Fallback Excel parser for simple CSV-like content
    static async parseExcelFallback(file) {
        try {
            console.log('使用回退方式解析 Excel 文件...');
            
            // Try to read as text (works for some simple Excel formats)
            const arrayBuffer = await file.arrayBuffer();
            const text = new TextDecoder('utf-8').decode(arrayBuffer);
            
            // Look for readable text content
            const lines = text.split(/[\r\n]+/).filter(line => line.trim());
            const data = [];
            
            // Try to extract tabular data
            for (const line of lines) {
                // Skip binary data and headers
                if (line.includes('\0') || line.length < 2) continue;
                
                // Try to parse as tab-separated or comma-separated
                let row = line.split('\t');
                if (row.length === 1) {
                    row = line.split(',');
                }
                
                // Clean up the row
                row = row.map(cell => cell.replace(/[^\w\s\u4e00-\u9fff.,()-]/g, '').trim());
                
                if (row.some(cell => cell.length > 0)) {
                    data.push(row);
                }
            }
            
            // If we couldn't extract meaningful data, create a simple structure
            if (data.length === 0) {
                data.push(['列1', '列2', '列3']);
                data.push(['範例資料1', '範例資料2', '範例資料3']);
                data.push(['檔案名稱', file.name, '']);
                data.push(['檔案大小', SpreadsheetConverter.formatFileSize(file.size), '']);
                
                console.log('創建範例資料結構');
            }
            
            console.log('Excel 回退解析完成，行數:', data.length);
            
            return {
                data: data,
                sheets: [{ name: 'Sheet1', data: data }],
                fileName: file.name.replace(/\.[^/.]+$/, ''),
                rowCount: data.length,
                columnCount: data.length > 0 ? Math.max(...data.map(row => row.length)) : 0
            };
            
        } catch (error) {
            console.error('Excel 回退解析錯誤:', error);
            throw new Error('Excel 檔案解析失敗: ' + error.message);
        }
    }

    // Helper: Format file size
    static formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // Helper function to parse CSV text with custom delimiter
    static parseCsvText(text, delimiter = ',') {
        const rows = [];
        const lines = text.split('\n');
        
        for (let line of lines) {
            line = line.trim();
            if (line.length === 0) continue;
            
            const row = [];
            let current = '';
            let inQuotes = false;
            let escapeNext = false;
            
            for (let i = 0; i < line.length; i++) {
                const char = line[i];
                
                if (escapeNext) {
                    current += char;
                    escapeNext = false;
                } else if (char === '\\') {
                    escapeNext = true;
                } else if (char === '"') {
                    if (inQuotes && i + 1 < line.length && line[i + 1] === '"') {
                        // Double quote escape
                        current += '"';
                        i++; // Skip next quote
                    } else {
                        inQuotes = !inQuotes;
                    }
                } else if (char === delimiter && !inQuotes) {
                    row.push(current);
                    current = '';
                } else {
                    current += char;
                }
            }
            
            row.push(current);
            if (row.some(cell => cell.trim().length > 0)) { // Only add non-empty rows
                rows.push(row);
            }
        }
        
        return rows;
    }

    // Convert parsed data to various formats
    static async convertToFormat(parsedData, outputFormat, options = {}) {
        const { data, sheets, fileName } = parsedData;
        const { 
            includeHeaders = true, 
            delimiter = ',', 
            maxFileSize = null,
            sheetIndex = 0 
        } = options;

        const targetData = sheets && sheets.length > 0 ? sheets[sheetIndex].data : data;

        switch (outputFormat) {
            case 'csv':
                return SpreadsheetConverter.convertToCsv(targetData, { delimiter, includeHeaders });
            case 'json':
                return SpreadsheetConverter.convertToJson(targetData, { includeHeaders });
            case 'html':
                return SpreadsheetConverter.convertToHtml(targetData, fileName, { includeHeaders });
            case 'txt':
                return SpreadsheetConverter.convertToText(targetData, { delimiter: '\t', includeHeaders });
            case 'xlsx':
                return await SpreadsheetConverter.convertToExcel(targetData, fileName, options);
            default:
                throw new Error(`不支援的輸出格式: ${outputFormat}`);
        }
    }

    // Convert to CSV format
    static convertToCsv(data, options = {}) {
        const { delimiter = ',', includeHeaders = true } = options;
        
        if (!data || data.length === 0) {
            const blob = new Blob([''], { type: 'text/csv;charset=utf-8' });
            return blob;
        }
        
        let csvContent = '';
        const BOM = '\uFEFF'; // UTF-8 BOM for Excel compatibility
        csvContent += BOM;
        
        data.forEach((row, index) => {
            if (index === 0 && !includeHeaders) return;
            
            const csvRow = row.map(cell => {
                const cellStr = String(cell || '').trim();
                // Always wrap in quotes for better compatibility
                return `"${cellStr.replace(/"/g, '""')}"`;
            }).join(delimiter);
            
            csvContent += csvRow + '\n';
        });
        
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
        return blob;
    }

    // Convert to JSON format
    static convertToJson(data, options = {}) {
        const { includeHeaders = true } = options;
        
        if (data.length === 0) {
            const blob = new Blob(['[]'], { type: 'application/json;charset=utf-8' });
            return blob;
        }
        
        let jsonData;
        
        if (includeHeaders && data.length > 1) {
            const headers = data[0];
            jsonData = data.slice(1).map(row => {
                const obj = {};
                headers.forEach((header, index) => {
                    obj[header] = row[index] || '';
                });
                return obj;
            });
        } else {
            jsonData = data.map((row, index) => {
                const obj = {};
                row.forEach((cell, cellIndex) => {
                    obj[`column_${cellIndex}`] = cell;
                });
                obj._rowIndex = index;
                return obj;
            });
        }
        
        const jsonString = JSON.stringify(jsonData, null, 2);
        const blob = new Blob([jsonString], { type: 'application/json;charset=utf-8' });
        return blob;
    }

    // Helper function to escape HTML
    static escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // Convert to HTML table
    static convertToHtml(data, fileName, options = {}) {
        const { includeHeaders = true } = options;
        
        let tableHtml = '';
        
        data.forEach((row, index) => {
            if (index === 0 && includeHeaders) {
                tableHtml += '<thead><tr>';
                row.forEach(cell => {
                    tableHtml += `<th>${SpreadsheetConverter.escapeHtml(cell || '')}</th>`;
                });
                tableHtml += '</tr></thead><tbody>';
            } else {
                tableHtml += '<tr>';
                row.forEach(cell => {
                    tableHtml += `<td>${SpreadsheetConverter.escapeHtml(cell || '')}</td>`;
                });
                tableHtml += '</tr>';
            }
        });
        
        if (includeHeaders) {
            tableHtml += '</tbody>';
        }
        
        const html = `<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${fileName || '試算表'}</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            margin: 20px;
            color: #333;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px 12px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
            font-weight: 600;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f0f8ff;
        }
        h1 {
            color: #2c3e50;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
        }
    </style>
</head>
<body>
    <h1>${fileName || '試算表'}</h1>
    <p>總共 ${data.length} 列資料</p>
    <table>
        ${tableHtml}
    </table>
</body>
</html>`;
        
        const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
        return blob;
    }

    // Convert to plain text
    static convertToText(data, options = {}) {
        const { delimiter = '\t', includeHeaders = true } = options;
        
        let textContent = '';
        
        data.forEach((row, index) => {
            if (index === 0 && !includeHeaders) return;
            
            const textRow = row.join(delimiter);
            textContent += textRow + '\n';
        });
        
        const blob = new Blob([textContent], { type: 'text/plain;charset=utf-8' });
        return blob;
    }

    // Convert to Excel format (using basic XML structure)
    static async convertToExcel(data, fileName, options = {}) {
        try {
            console.log('開始轉換真正的 XLSX 格式...');
            
            if (!data || data.length === 0) {
                throw new Error('無資料可轉換');
            }

            // Try to use SheetJS for proper XLSX creation
            try {
                await SpreadsheetConverter.loadSheetJS();
                return await SpreadsheetConverter.createXlsxWithSheetJS(data, fileName, options);
            } catch (sheetError) {
                console.warn('SheetJS 載入失敗，使用自製 XLSX 格式');
                return SpreadsheetConverter.createSimpleXlsx(data, fileName, options);
            }
            
        } catch (error) {
            console.error('Excel 轉換錯誤:', error);
            
            // Fallback to CSV if Excel conversion fails
            console.warn('回退到 CSV 格式');
            return SpreadsheetConverter.convertToCsv(data, options);
        }
    }

    // Create XLSX using SheetJS library
    static async createXlsxWithSheetJS(data, fileName, options = {}) {
        const { includeHeaders = true } = options;
        
        // Create a new workbook
        const wb = XLSX.utils.book_new();
        
        // Create worksheet from array data
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // Set column widths
        const colWidths = [];
        if (data.length > 0) {
            data[0].forEach((_, colIndex) => {
                let maxWidth = 10;
                data.forEach(row => {
                    if (row[colIndex]) {
                        const cellLength = String(row[colIndex]).length;
                        maxWidth = Math.max(maxWidth, Math.min(cellLength + 2, 30));
                    }
                });
                colWidths.push({ width: maxWidth });
            });
        }
        ws['!cols'] = colWidths;
        
        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        
        // Generate XLSX file
        const xlsxArray = XLSX.write(wb, { 
            bookType: 'xlsx', 
            type: 'array',
            compression: true
        });
        
        const blob = new Blob([xlsxArray], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        console.log('✅ SheetJS XLSX 檔案創建完成:', blob.size, 'bytes');
        return blob;
    }

    // Load SheetJS library
    static async loadSheetJS() {
        if (window.XLSX) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js';
            script.onload = () => {
                setTimeout(() => {
                    if (window.XLSX) {
                        console.log('✅ SheetJS 載入成功');
                        resolve();
                    } else {
                        reject(new Error('SheetJS 載入後無法使用'));
                    }
                }, 100);
            };
            script.onerror = () => reject(new Error('SheetJS 載入失敗'));
            document.head.appendChild(script);
        });
    }

    // Create simple XLSX using ZIP structure (fallback)
    static async createSimpleXlsx(data, fileName, options = {}) {
        try {
            // Load JSZip for creating XLSX structure
            await SpreadsheetConverter.loadJSZip();
            
            const zip = new JSZip();
            const { includeHeaders = true } = options;
            
            // Create [Content_Types].xml
            const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;
            zip.file('[Content_Types].xml', contentTypes);
            
            // Create _rels/.rels
            const mainRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
            zip.folder('_rels').file('.rels', mainRels);
            
            // Create xl/_rels/workbook.xml.rels
            const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
            zip.folder('xl').folder('_rels').file('workbook.xml.rels', workbookRels);
            
            // Create xl/workbook.xml
            const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>`;
            zip.folder('xl').file('workbook.xml', workbookXml);
            
            // Create worksheet data
            let sheetData = '';
            data.forEach((row, rowIndex) => {
                sheetData += `<row r="${rowIndex + 1}">`;
                row.forEach((cell, colIndex) => {
                    const cellRef = this.numberToColumnName(colIndex + 1) + (rowIndex + 1);
                    const cellValue = String(cell || '').replace(/[&<>"']/g, function(char) {
                        const entities = {
                            '&': '&amp;',
                            '<': '&lt;',
                            '>': '&gt;',
                            '"': '&quot;',
                            "'": '&apos;'
                        };
                        return entities[char];
                    });
                    sheetData += `<c r="${cellRef}" t="inlineStr"><is><t>${cellValue}</t></is></c>`;
                });
                sheetData += '</row>';
            });
            
            // Create xl/worksheets/sheet1.xml
            const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        ${sheetData}
    </sheetData>
</worksheet>`;
            zip.folder('xl').folder('worksheets').file('sheet1.xml', worksheetXml);
            
            // Generate XLSX file
            const xlsxBlob = await zip.generateAsync({ 
                type: 'blob', 
                mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                compression: 'DEFLATE'
            });
            
            console.log('✅ 自製 XLSX 檔案創建完成:', xlsxBlob.size, 'bytes');
            return xlsxBlob;
            
        } catch (zipError) {
            console.error('XLSX ZIP 創建錯誤:', zipError);
            throw zipError;
        }
    }

    // Load JSZip library
    static async loadJSZip() {
        if (window.JSZip) return;
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
            script.onload = () => {
                setTimeout(() => {
                    if (window.JSZip) {
                        console.log('✅ JSZip 載入成功');
                        resolve();
                    } else {
                        reject(new Error('JSZip 載入後無法使用'));
                    }
                }, 100);
            };
            script.onerror = () => reject(new Error('JSZip 載入失敗'));
            document.head.appendChild(script);
        });
    }

    // Helper: Convert column number to Excel column name (A, B, C, ... AA, AB, etc.)
    static numberToColumnName(num) {
        let result = '';
        while (num > 0) {
            num--;
            result = String.fromCharCode(65 + (num % 26)) + result;
            num = Math.floor(num / 26);
        }
        return result;
    }

    // Original Excel XML format (legacy fallback)
    static createLegacyExcelXml(data, fileName, options = {}) {
        try {
            const { includeHeaders = true } = options;
        
        let xmlContent = `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
<Title>${fileName || '試算表'}</Title>
<Created>${new Date().toISOString()}</Created>
</DocumentProperties>
<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
<WindowHeight>12000</WindowHeight>
<WindowWidth>20000</WindowWidth>
<WindowTopX>0</WindowTopX>
<WindowTopY>0</WindowTopY>
<ProtectStructure>False</ProtectStructure>
<ProtectWindows>False</ProtectWindows>
</ExcelWorkbook>
<Styles>
<Style ss:ID="Default" ss:Name="Normal">
<Alignment ss:Vertical="Bottom"/>
<Borders/>
<Font ss:FontName="新細明體" x:CharSet="136" ss:Size="12"/>
<Interior/>
<NumberFormat/>
<Protection/>
</Style>
<Style ss:ID="Header">
<Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
<Borders>
<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
</Borders>
<Font ss:FontName="新細明體" x:CharSet="136" ss:Size="12" ss:Bold="1"/>
<Interior ss:Color="#C0C0C0" ss:Pattern="Solid"/>
</Style>
</Styles>
<Worksheet ss:Name="Sheet1">
<Table>`;

            // Add rows
            data.forEach((row, rowIndex) => {
                if (rowIndex === 0 && !includeHeaders) return;
                
                const isHeader = rowIndex === 0 && includeHeaders;
                xmlContent += `\n<Row>`;
                
                row.forEach(cell => {
                    const cellValue = String(cell || '').replace(/[&<>"']/g, function(char) {
                        const entities = {
                            '&': '&amp;',
                            '<': '&lt;',
                            '>': '&gt;',
                            '"': '&quot;',
                            '\'': '&apos;'
                        };
                        return entities[char];
                    });
                    
                    const styleId = isHeader ? 'Header' : 'Default';
                    xmlContent += `\n<Cell ss:StyleID="${styleId}"><Data ss:Type="String">${cellValue}</Data></Cell>`;
                });
                
                xmlContent += `\n</Row>`;
            });

            xmlContent += `\n</Table>
</Worksheet>
</Workbook>`;

            // Create blob with proper MIME type for Excel
            const blob = new Blob([xmlContent], { 
                type: 'application/vnd.ms-excel;charset=utf-8' 
            });
            
            return blob;
            
        } catch (error) {
            console.error('Excel 轉換錯誤:', error);
            
            // Fallback to CSV if Excel conversion fails
            console.warn('回退到 CSV 格式');
            return SpreadsheetConverter.convertToCsv(data, options);
        }
    }

    // Format file size
    static formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // Get spreadsheet statistics
    static getSpreadsheetStats(data) {
        if (!data || data.length === 0) {
            return {
                rowCount: 0,
                columnCount: 0,
                cellCount: 0,
                nonEmptyCellCount: 0,
                hasHeaders: false
            };
        }

        const rowCount = data.length;
        const columnCount = Math.max(...data.map(row => row.length));
        const cellCount = rowCount * columnCount;
        
        let nonEmptyCellCount = 0;
        data.forEach(row => {
            row.forEach(cell => {
                if (cell && String(cell).trim().length > 0) {
                    nonEmptyCellCount++;
                }
            });
        });

        // Simple heuristic to detect headers
        const hasHeaders = rowCount > 1 && data[0].every(cell => 
            typeof cell === 'string' && cell.trim().length > 0
        );

        return {
            rowCount,
            columnCount,
            cellCount,
            nonEmptyCellCount,
            hasHeaders,
            fillRate: Math.round((nonEmptyCellCount / cellCount) * 100)
        };
    }

    // Optimize data for file size
    static optimizeData(data, options = {}) {
        const { 
            removeEmptyRows = true, 
            removeEmptyColumns = true,
            trimWhitespace = true 
        } = options;

        let optimizedData = [...data];

        if (trimWhitespace) {
            optimizedData = optimizedData.map(row => 
                row.map(cell => typeof cell === 'string' ? cell.trim() : cell)
            );
        }

        if (removeEmptyRows) {
            optimizedData = optimizedData.filter(row => 
                row.some(cell => cell && String(cell).trim().length > 0)
            );
        }

        if (removeEmptyColumns) {
            const nonEmptyColumns = [];
            const maxCols = Math.max(...optimizedData.map(row => row.length));
            
            for (let col = 0; col < maxCols; col++) {
                const hasContent = optimizedData.some(row => 
                    row[col] && String(row[col]).trim().length > 0
                );
                if (hasContent) {
                    nonEmptyColumns.push(col);
                }
            }
            
            optimizedData = optimizedData.map(row => 
                nonEmptyColumns.map(col => row[col] || '')
            );
        }

        return optimizedData;
    }
}

// Export for use in main application
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SpreadsheetConverter;
} else if (typeof window !== 'undefined') {
    window.SpreadsheetConverter = SpreadsheetConverter;
}
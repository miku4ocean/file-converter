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
            // Note: This would require SheetJS library
            // For now, return a placeholder implementation
            const arrayBuffer = await file.arrayBuffer();
            
            // Placeholder: In real implementation, use SheetJS
            return {
                data: [
                    ['此功能需要 SheetJS 函式庫支援'],
                    ['請在 HTML 中加入：'],
                    ['<script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>'],
                    [''],
                    ['檔案名稱', file.name],
                    ['檔案大小', SpreadsheetConverter.formatFileSize(file.size)],
                    ['檔案類型', file.type]
                ],
                sheets: [{ name: 'Info', data: [] }],
                fileName: file.name.replace(/\.[^/.]+$/, ''),
                rowCount: 7,
                columnCount: 2
            };
        } catch (error) {
            throw new Error('Excel 檔案讀取失敗: ' + error.message);
        }
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
            if (!data || data.length === 0) {
                throw new Error('無資料可轉換');
            }

            // Create a basic Excel XML format (works with Excel 2003+)
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

    // Helper function to escape HTML
    static escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
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
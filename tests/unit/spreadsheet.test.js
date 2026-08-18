// Unit tests for assets/js/converters/spreadsheet.js
// Run with: node --test tests/unit
'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');

const SpreadsheetConverter = require('../../assets/js/converters/spreadsheet.js');

describe('SpreadsheetConverter.parseCsvText', () => {
    test('splits simple comma rows', () => {
        const rows = SpreadsheetConverter.parseCsvText('a,b,c\n1,2,3', ',');
        assert.deepEqual(rows, [['a', 'b', 'c'], ['1', '2', '3']]);
    });

    test('handles double-quote escaping (RFC 4180)', () => {
        const rows = SpreadsheetConverter.parseCsvText('"She said ""hi""",ok', ',');
        assert.deepEqual(rows, [['She said "hi"', 'ok']]);
    });

    test('keeps a delimiter inside quotes as literal data', () => {
        const rows = SpreadsheetConverter.parseCsvText('"a,b",c', ',');
        assert.deepEqual(rows, [['a,b', 'c']]);
    });

    // REGRESSION (real bug): the parser used to treat backslash as an escape
    // character (escapeNext), which is not part of CSV/TSV (RFC 4180 only
    // doubles quotes). Any literal backslash in a cell — e.g. a Windows path,
    // a regex, or a LaTeX snippet — was silently deleted from the parsed
    // output, corrupting the data on every CSV import.
    test('preserves literal backslashes in unquoted cells (Windows paths, regex, etc.)', () => {
        const rows = SpreadsheetConverter.parseCsvText('C:\\Users\\test,data', ',');
        assert.deepEqual(rows, [['C:\\Users\\test', 'data']]);
    });

    test('preserves literal backslashes inside quoted cells', () => {
        const rows = SpreadsheetConverter.parseCsvText('"regex: \\d+\\s*",note', ',');
        assert.deepEqual(rows, [['regex: \\d+\\s*', 'note']]);
    });

    test('supports a custom delimiter (TSV)', () => {
        const rows = SpreadsheetConverter.parseCsvText('a\tb\n1\t2', '\t');
        assert.deepEqual(rows, [['a', 'b'], ['1', '2']]);
    });

    test('drops fully blank lines', () => {
        const rows = SpreadsheetConverter.parseCsvText('a,b\n\n1,2\n   \n', ',');
        assert.deepEqual(rows, [['a', 'b'], ['1', '2']]);
    });
});

describe('SpreadsheetConverter.convertToCsv', () => {
    test('includes header row by default and wraps cells in quotes', async () => {
        const blob = SpreadsheetConverter.convertToCsv([['Name', 'Age'], ['Alice', '30']], { includeHeaders: true });
        // Note: Blob.text() decodes as UTF-8 and strips a leading BOM (same as
        // browsers), so the BOM the converter prepends for Excel compatibility
        // isn't visible here - only the decoded row content is asserted.
        const text = await blob.text();
        assert.match(text, /^"Name","Age"\n"Alice","30"\n$/);
    });

    test('omits header row when includeHeaders is false', async () => {
        const blob = SpreadsheetConverter.convertToCsv([['Name', 'Age'], ['Alice', '30']], { includeHeaders: false });
        const text = await blob.text();
        assert.doesNotMatch(text, /Name/);
        assert.match(text, /"Alice","30"/);
    });

    test('doubles embedded quote characters', async () => {
        const blob = SpreadsheetConverter.convertToCsv([['He said "hi"']], { includeHeaders: true });
        const text = await blob.text();
        assert.match(text, /"He said ""hi"""/);
    });

    // SECURITY (OWASP CSV Injection / CSV Formula Injection): a cell value
    // that starts with =, +, -, @, tab, or CR can be interpreted as a
    // formula/command by Excel or Google Sheets when the CSV is opened,
    // e.g. =cmd|'/c calc'!A1 can trigger code execution. Such values must be
    // prefixed with a leading single quote so spreadsheet apps treat them as
    // literal text instead of evaluating them.
    test('a benign value passes through unchanged', async () => {
        const blob = SpreadsheetConverter.convertToCsv([['Alice']], { includeHeaders: true });
        const text = await blob.text();
        assert.match(text, /^"Alice"\n$/);
    });

    test('prefixes a formula-injection payload with a leading single quote', async () => {
        const blob = SpreadsheetConverter.convertToCsv([["=cmd|'/c calc'!A1"]], { includeHeaders: true });
        const text = await blob.text();
        assert.match(text, /^"'=cmd\|'/);
    });

    test('prefixes +/-/@-prefixed values with a leading single quote', async () => {
        const blob = SpreadsheetConverter.convertToCsv([
            ['+1+1', '-1+1', '@SUM(A1:A2)']
        ], { includeHeaders: true });
        const text = await blob.text();
        assert.match(text, /^"'\+1\+1","'-1\+1","'@SUM\(A1:A2\)"\n$/);
    });
});

describe('SpreadsheetConverter.sanitizeCsvField', () => {
    test('leaves normal text untouched', () => {
        assert.equal(SpreadsheetConverter.sanitizeCsvField('Alice'), 'Alice');
    });

    test('prepends a single quote to formula-triggering prefixes', () => {
        assert.equal(SpreadsheetConverter.sanitizeCsvField("=cmd|'/c calc'!A1"), "'=cmd|'/c calc'!A1");
        assert.equal(SpreadsheetConverter.sanitizeCsvField('+1+1'), "'+1+1");
        assert.equal(SpreadsheetConverter.sanitizeCsvField('-1+1'), "'-1+1");
        assert.equal(SpreadsheetConverter.sanitizeCsvField('@SUM(A1:A2)'), "'@SUM(A1:A2)");
    });

    // R2 behavior fix: the guard still *checks* the trimmed value (so leading
    // whitespace cannot bypass it), but the cell data itself is no longer
    // trimmed/mutated - the apostrophe is prefixed to the original value.
    // A cell of "'   =cmd" is still inert in Excel (leading ' = literal text).
    test('checks the trimmed value, so leading whitespace does not bypass the guard', () => {
        assert.equal(SpreadsheetConverter.sanitizeCsvField('   =cmd'), "'   =cmd");
    });

    // R2 behavior fix: a raw leading tab/CR is itself an Excel formula/command
    // trigger (OWASP CSV Injection prefix list). Previously the tab/CR was
    // silently DELETED by the unconditional trim (data mutation); now the
    // original value is preserved and neutralized with a leading apostrophe.
    test('a leading tab or CR is neutralized with a quote, not silently deleted', () => {
        assert.equal(SpreadsheetConverter.sanitizeCsvField('\tevil'), "'\tevil");
        assert.equal(SpreadsheetConverter.sanitizeCsvField('\revil'), "'\revil");
    });
});

describe('SpreadsheetConverter.sanitizeCsvField - 資料保真（R2 回歸）', () => {
    // REGRESSION (real bug): String(value || '') 把 falsy 但真實存在的儲存格值
    // 0 與 false 變成空字串 —— 每次 CSV 匯出都把「數值 0」欄位靜默清空。
    test('preserves numeric zero and boolean false instead of blanking them', () => {
        assert.equal(SpreadsheetConverter.sanitizeCsvField(0), '0');
        assert.equal(SpreadsheetConverter.sanitizeCsvField(false), 'false');
    });

    // REGRESSION (real bug): 每個儲存格都被 .trim()，合法的前後空白被靜默改寫。
    test('does not trim legitimate leading/trailing whitespace of ordinary cells', () => {
        assert.equal(SpreadsheetConverter.sanitizeCsvField('  padded  '), '  padded  ');
    });

    // REGRESSION (real bug): 純數字（尤其負數）被當公式觸發字元加上前導單引號，
    // 整欄負數輸出成 '-42 之類的文字，數值欄毀損且使用者不易察覺。
    // 純數字在 Excel 中求值就是它自己，無注入風險，必須原樣通過。
    test('plain numbers (incl. negative / signed / scientific) are not prefixed', () => {
        assert.equal(SpreadsheetConverter.sanitizeCsvField(-42), '-42');
        assert.equal(SpreadsheetConverter.sanitizeCsvField('-3.14'), '-3.14');
        assert.equal(SpreadsheetConverter.sanitizeCsvField('+7'), '+7');
        assert.equal(SpreadsheetConverter.sanitizeCsvField('-1e-5'), '-1e-5');
    });

    test('non-numeric formula triggers are still neutralized', () => {
        assert.equal(SpreadsheetConverter.sanitizeCsvField('-1+1'), "'-1+1");
        assert.equal(SpreadsheetConverter.sanitizeCsvField('=SUM(A1)'), "'=SUM(A1)");
    });

    test('convertToCsv keeps 0 and -5 as data', async () => {
        const blob = SpreadsheetConverter.convertToCsv([[0, -5]], { includeHeaders: true });
        const text = await blob.text();
        assert.match(text, /^"0","-5"\n$/);
    });
});

describe('SpreadsheetConverter.parseCsvText - RFC 4180 引號內換行（R2 回歸）', () => {
    // REGRESSION (real bug): 解析前先用 '\n' 把整份文字切行，引號內含換行的
    // 多行儲存格（RFC 4180 合法）被硬拆成兩列，之後所有列全部錯位，
    // 使用者不易察覺資料已毀損。
    test('keeps a newline inside quotes as part of the cell', () => {
        const rows = SpreadsheetConverter.parseCsvText('"a\nb",c', ',');
        assert.deepEqual(rows, [['a\nb', 'c']]);
    });

    test('rows after a multi-line cell stay aligned', () => {
        const rows = SpreadsheetConverter.parseCsvText('h1,h2\n"multi\nline",x\ny,z', ',');
        assert.deepEqual(rows, [['h1', 'h2'], ['multi\nline', 'x'], ['y', 'z']]);
    });

    test('handles CRLF row endings', () => {
        const rows = SpreadsheetConverter.parseCsvText('a,b\r\n1,2\r\n', ',');
        assert.deepEqual(rows, [['a', 'b'], ['1', '2']]);
    });
});

describe('SpreadsheetConverter.convertToJson - falsy 儲存格保真（R2 回歸）', () => {
    // REGRESSION (real bug): row[index] || '' 把 0 與 false 變成 ''，
    // Excel/SheetJS 解析出的數值 0 在 JSON 匯出中被清空。
    test('keeps 0 and false cell values', async () => {
        const blob = SpreadsheetConverter.convertToJson([['n', 'b'], [0, false]], { includeHeaders: true });
        const parsed = JSON.parse(await blob.text());
        assert.deepEqual(parsed, [{ n: 0, b: false }]);
    });
});

describe('SpreadsheetConverter.parseCsv - UTF-16 編碼偵測（R2 回歸）', () => {
    // REGRESSION (real bug): parseCsv 用 file.text()（固定 UTF-8 解碼）。
    // Excel「Unicode 文字」匯出與 Windows 記事本「Unicode」存檔皆為 UTF-16LE，
    // 解出來是夾滿 NUL 的亂碼且無任何報錯 —— 靜默半成品。
    test('decodes a UTF-16LE (BOM) CSV correctly', async () => {
        const payload = Buffer.from('名稱,值\nA,1', 'utf16le');
        const bytes = new Uint8Array(2 + payload.length);
        bytes.set([0xFF, 0xFE], 0);
        bytes.set(payload, 2);
        const file = new File([bytes], 'u16.csv', { type: 'text/csv' });
        const parsed = await SpreadsheetConverter.parseCsv(file);
        assert.deepEqual(parsed.data, [['名稱', '值'], ['A', '1']]);
    });

    test('decodes a UTF-16BE (BOM) CSV correctly', async () => {
        const le = Buffer.from('名稱,值\nA,1', 'utf16le');
        const be = Buffer.from(le);
        be.swap16();
        const bytes = new Uint8Array(2 + be.length);
        bytes.set([0xFE, 0xFF], 0);
        bytes.set(be, 2);
        const file = new File([bytes], 'u16be.csv', { type: 'text/csv' });
        const parsed = await SpreadsheetConverter.parseCsv(file);
        assert.deepEqual(parsed.data, [['名稱', '值'], ['A', '1']]);
    });

    test('still decodes plain UTF-8 with BOM stripped (guard)', async () => {
        const body = Buffer.from('a,b\n1,2', 'utf8');
        const bytes = new Uint8Array(3 + body.length);
        bytes.set([0xEF, 0xBB, 0xBF], 0);
        bytes.set(body, 3);
        const file = new File([bytes], 'u8.csv', { type: 'text/csv' });
        const parsed = await SpreadsheetConverter.parseCsv(file);
        assert.deepEqual(parsed.data, [['a', 'b'], ['1', '2']]);
    });
});

describe('SpreadsheetConverter.convertToJson', () => {
    test('maps subsequent rows to objects keyed by header row', () => {
        const blob = SpreadsheetConverter.convertToJson([['Name', 'Age'], ['Alice', '30']], { includeHeaders: true });
        return blob.text().then(text => {
            assert.deepEqual(JSON.parse(text), [{ Name: 'Alice', Age: '30' }]);
        });
    });

    test('returns an empty array for an empty dataset', async () => {
        const blob = SpreadsheetConverter.convertToJson([], {});
        assert.equal(await blob.text(), '[]');
    });
});

describe('SpreadsheetConverter.numberToColumnName', () => {
    test('converts 1-based column indices to Excel-style letters', () => {
        assert.equal(SpreadsheetConverter.numberToColumnName(1), 'A');
        assert.equal(SpreadsheetConverter.numberToColumnName(26), 'Z');
        assert.equal(SpreadsheetConverter.numberToColumnName(27), 'AA');
        assert.equal(SpreadsheetConverter.numberToColumnName(52), 'AZ');
        assert.equal(SpreadsheetConverter.numberToColumnName(703), 'AAA');
    });
});

describe('SpreadsheetConverter.getSpreadsheetStats', () => {
    test('computes row/column/cell counts and fill rate', () => {
        const stats = SpreadsheetConverter.getSpreadsheetStats([
            ['Name', 'Age'],
            ['Alice', '30'],
            ['', '']
        ]);
        assert.equal(stats.rowCount, 3);
        assert.equal(stats.columnCount, 2);
        assert.equal(stats.cellCount, 6);
        assert.equal(stats.nonEmptyCellCount, 4);
        assert.equal(stats.hasHeaders, true);
    });

    test('handles empty input without throwing', () => {
        const stats = SpreadsheetConverter.getSpreadsheetStats([]);
        assert.equal(stats.rowCount, 0);
        assert.equal(stats.hasHeaders, false);
    });
});

describe('SpreadsheetConverter.optimizeData', () => {
    test('trims whitespace and removes fully empty rows/columns', () => {
        const optimized = SpreadsheetConverter.optimizeData([
            [' a ', '', 'c '],
            ['', '', ''],
            ['d', '', 'e']
        ]);
        assert.deepEqual(optimized, [['a', 'c'], ['d', 'e']]);
    });
});

describe('SpreadsheetConverter.getFileType / isValidSpreadsheetFile', () => {
    test('extracts lowercase extension without the dot', () => {
        assert.equal(SpreadsheetConverter.getFileType({ name: 'Report.CSV' }), 'csv');
    });

    test('accepts a recognized extension even with an unusual MIME type', () => {
        const file = { name: 'data.csv', type: 'application/octet-stream' };
        assert.equal(SpreadsheetConverter.isValidSpreadsheetFile(file), true);
    });

    test('rejects an unsupported extension and generic MIME type', () => {
        const file = { name: 'archive.zip', type: 'application/zip' };
        assert.equal(SpreadsheetConverter.isValidSpreadsheetFile(file), false);
    });
});

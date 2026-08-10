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

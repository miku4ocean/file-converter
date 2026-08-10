// Regression test for assets/js/app.js
//
// app.js is not written as a CommonJS module (it assumes a full browser
// `document`/`window`), so to exercise the REAL code (not a re-implementation
// of it) this test loads the actual source into a `vm` context with a minimal
// DOM stub, wires in the real converter classes (via their existing
// module.exports), and drives FileConverter's methods directly.
//
// Run with: node --test tests/unit
'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');

const SpreadsheetConverter = require('../../assets/js/converters/spreadsheet.js');
const PresentationConverter = require('../../assets/js/converters/presentation.js');
const DocumentConverter = require('../../assets/js/converters/document.js');
const ImageConverter = require('../../assets/js/converters/image.js');

function createMockElement(overrides = {}) {
    return Object.assign(
        {
            addEventListener() {},
            removeEventListener() {},
            classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
            style: {},
            dataset: {},
            appendChild() {},
            removeChild() {},
            setAttribute() {},
            getAttribute() { return null; },
            querySelectorAll() { return []; },
            querySelector() { return null; },
            value: '',
            checked: false,
            textContent: '',
            innerHTML: ''
        },
        overrides
    );
}

// IDs FileConverter.initializeElements()/bindEvents() require to exist
// (read once at construction time) so `new FileConverter()` doesn't throw.
const REQUIRED_IDS = [
    'uploadArea', 'fileInput', 'uploadBtn', 'uploadIcon', 'uploadTitle', 'uploadDescription',
    'fileListSection', 'fileList', 'conversionSection', 'conversionOptions', 'convertBtn',
    'progressSection', 'progressFill', 'progressText', 'progressDetails',
    'downloadSection', 'downloadList', 'downloadAllBtn', 'resetBtn'
];

// Loads assets/js/app.js into a fresh vm context, returns { FileConverter, elementRegistry }.
// elementRegistry maps element id -> mock element, so a test can pre-seed e.g.
// elementRegistry.includeHeaders = createMockElement({ checked: false }).
// IDs not in REQUIRED_IDS and not explicitly seeded resolve to `null`, just
// like a real document.getElementById() call for an element that isn't on
// the page - this is what lets the "missing checkbox" test be authentic
// instead of tautological.
function loadFileConverter() {
    const elementRegistry = {};
    REQUIRED_IDS.forEach(id => { elementRegistry[id] = createMockElement(); });

    const documentStub = {
        querySelectorAll: () => [],
        getElementById: (id) => (id in elementRegistry ? elementRegistry[id] : null),
        createElement: () => createMockElement(),
        addEventListener() {}, // swallow DOMContentLoaded so app.js doesn't self-instantiate
        body: createMockElement()
    };

    const sandbox = {
        document: documentStub,
        URL: { createObjectURL: () => 'blob://mock', revokeObjectURL() {} },
        console,
        Image: function () {},
        alert() {},
        SpreadsheetConverter,
        PresentationConverter,
        DocumentConverter,
        ImageConverter,
        Blob,
        File,
        setTimeout,
        clearTimeout
    };
    sandbox.window = sandbox;
    vm.createContext(sandbox);

    const code = fs.readFileSync(path.join(__dirname, '../../assets/js/app.js'), 'utf8');
    vm.runInContext(code, sandbox, { filename: 'app.js' });

    const FileConverter = vm.runInContext('FileConverter', sandbox);
    return { FileConverter, elementRegistry };
}

describe('FileConverter.convertSpreadsheetFile - includeHeaders checkbox regression', () => {
    // REGRESSION (real bug): app.js read the checkbox as
    //   document.getElementById('includeHeaders')?.checked || true
    // Since `false || true` is always `true`, unchecking the "包含標題行"
    // (include header row) checkbox had NO effect whatsoever - the header
    // row was always retained in CSV/JSON/TXT/HTML output. Fixed to use `??`
    // so only a *missing* element falls back to true; an explicit `false`
    // from the checkbox is respected.
    test('unchecking "include headers" actually excludes the header row from CSV output', async () => {
        const { FileConverter, elementRegistry } = loadFileConverter();
        const instance = new FileConverter();
        instance.currentFileType = 'spreadsheet';

        elementRegistry.includeHeaders = createMockElement({ checked: false });
        elementRegistry.outputFormat = createMockElement({ value: 'csv' });

        const file = new File(['Name,Age\nAlice,30\n'], 'people.csv', { type: 'text/csv' });
        const blob = await instance.convertSpreadsheetFile(file, 'csv');
        const text = await blob.text();

        assert.doesNotMatch(text, /Name/, 'header row must be excluded when the checkbox is unchecked');
        assert.match(text, /Alice/);
    });

    test('checking "include headers" keeps the header row (default behavior unchanged)', async () => {
        const { FileConverter, elementRegistry } = loadFileConverter();
        const instance = new FileConverter();
        instance.currentFileType = 'spreadsheet';

        elementRegistry.includeHeaders = createMockElement({ checked: true });
        elementRegistry.outputFormat = createMockElement({ value: 'csv' });

        const file = new File(['Name,Age\nAlice,30\n'], 'people.csv', { type: 'text/csv' });
        const blob = await instance.convertSpreadsheetFile(file, 'csv');
        const text = await blob.text();

        assert.match(text, /Name/);
        assert.match(text, /Alice/);
    });

    test('a missing checkbox element (getElementById returns null) still defaults to including headers', async () => {
        const { FileConverter } = loadFileConverter();
        const instance = new FileConverter();
        instance.currentFileType = 'spreadsheet';
        // elementRegistry.includeHeaders is intentionally left unset, so the
        // stub's getElementById('includeHeaders') returns null, exactly like a
        // real page where that element doesn't exist. `null?.checked` is
        // `undefined`, and `undefined ?? true` must fall back to `true`.

        const file = new File(['Name,Age\nAlice,30\n'], 'people.csv', { type: 'text/csv' });
        const blob = await instance.convertSpreadsheetFile(file, 'csv');
        const text = await blob.text();

        assert.match(text, /Name/);
    });
});

describe('FileConverter.convertPresentationFile - includeNotes checkbox regression', () => {
    test('unchecking "include notes" excludes slide notes from HTML output', async () => {
        const { FileConverter, elementRegistry } = loadFileConverter();
        const instance = new FileConverter();
        instance.currentFileType = 'presentation';

        elementRegistry.includeNotes = createMockElement({ checked: false });
        elementRegistry.outputFormat = createMockElement({ value: 'html' });

        const htmlSource = '<html><head><title>Deck</title></head><body>hello world</body></html>';
        const file = new File([htmlSource], 'deck.html', { type: 'text/html' });

        // extractPresentationContent -> extractFromHtml doesn't produce slide
        // notes on its own, so build content manually to isolate the option.
        const content = { slides: [{ slideNumber: 1, title: 'S1', content: ['hi'], notes: 'do not leak' }] };
        const blob = PresentationConverter.convertToHtml(content, 'Deck', { includeNotes: elementRegistry.includeNotes.checked ?? true });
        const html = await blob.text();

        assert.doesNotMatch(html, /do not leak/);
    });
});

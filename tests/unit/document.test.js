// Unit tests for assets/js/converters/document.js
// Run with: node --test tests/unit
'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');

const DocumentConverter = require('../../assets/js/converters/document.js');

describe('DocumentConverter.escapeXml', () => {
    test('escapes the five XML special characters', () => {
        assert.equal(
            DocumentConverter.escapeXml(`<a href="x">Tom & Jerry's "quote"</a>`),
            '&lt;a href=&quot;x&quot;&gt;Tom &amp; Jerry&apos;s &quot;quote&quot;&lt;/a&gt;'
        );
    });

    test('returns empty string for falsy input', () => {
        assert.equal(DocumentConverter.escapeXml(''), '');
        assert.equal(DocumentConverter.escapeXml(null), '');
        assert.equal(DocumentConverter.escapeXml(undefined), '');
    });
});

describe('DocumentConverter.escapeRtf', () => {
    test('escapes backslashes and braces, converts newlines to \\par', () => {
        assert.equal(
            DocumentConverter.escapeRtf('a\\b{c}\nd'),
            'a\\\\b\\{c\\}\\par d'
        );
    });
});

describe('DocumentConverter.compressContent', () => {
    const messy = 'Line1   with   spaces\n\n\n\nLine2\t\ttabbed';

    test('standard level collapses spaces/tabs and trims', () => {
        const out = DocumentConverter.compressContent(messy, 'standard');
        assert.equal(out.includes('   '), false);
        assert.equal(out, out.trim());
    });

    test('medium level also collapses 3+ newlines to a blank line', () => {
        const out = DocumentConverter.compressContent(messy, 'medium');
        assert.equal(/\n{3,}/.test(out), false);
    });

    test('high level collapses all whitespace runs including newlines', () => {
        const out = DocumentConverter.compressContent(messy, 'high');
        assert.equal(/\n\s*\n/.test(out), false);
    });
});

describe('DocumentConverter.convertToText / convertToHtml / convertToMarkdown (live definitions)', () => {
    test('convertToText prepends an underlined title', async () => {
        const blob = DocumentConverter.convertToText('Body text', 'My Title');
        const text = await blob.text();
        assert.equal(text, 'My Title\n========\n\nBody text');
    });

    test('convertToHtml escapes title and content (XML/HTML special chars)', async () => {
        const blob = DocumentConverter.convertToHtml('<b>bold</b> & stuff', 'A & B');
        const text = await blob.text();
        assert.match(text, /<title>A &amp; B<\/title>/);
        assert.match(text, /<h1>A &amp; B<\/h1>/);
        // paragraph content must be escaped too, not injected as raw markup
        assert.match(text, /&lt;b&gt;bold&lt;\/b&gt; &amp; stuff/);
        assert.doesNotMatch(text, /<b>bold<\/b>/);
    });

    test('convertToMarkdown adds an H1 heading from the title', async () => {
        const blob = DocumentConverter.convertToMarkdown('para one\n\npara two', 'Doc Title');
        const text = await blob.text();
        assert.match(text, /^# Doc Title\n\n/);
        assert.match(text, /para one/);
        assert.match(text, /para two/);
    });
});

describe('DocumentConverter.convertToRtf', () => {
    test('produces a well-formed RTF wrapper and escapes special chars', async () => {
        const blob = DocumentConverter.convertToRtf('line one\n\nline {two}', 'Title\\Slash');
        const text = await blob.text();
        assert.match(text, /^\{\\rtf1\\ansi\\deff0/);
        assert.match(text, /Title\\\\Slash/); // backslash escaped
        assert.match(text, /line \\\{two\\\}/); // braces escaped
        assert.match(text, /}$/);
    });
});

describe('DocumentConverter.getFileType / isValidDocumentFile', () => {
    test('extracts lowercase extension without the dot', () => {
        assert.equal(DocumentConverter.getFileType({ name: 'Report.TXT' }), 'txt');
    });

    test('accepts a recognized extension', () => {
        assert.equal(DocumentConverter.isValidDocumentFile({ name: 'notes.md', type: '' }), true);
    });

    test('rejects an unsupported file', () => {
        assert.equal(DocumentConverter.isValidDocumentFile({ name: 'movie.mp4', type: 'video/mp4' }), false);
    });
});

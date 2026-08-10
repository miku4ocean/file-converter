// Unit tests for assets/js/converters/presentation.js
// Run with: node --test tests/unit
'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');

const PresentationConverter = require('../../assets/js/converters/presentation.js');

describe('PresentationConverter.convertToHtml - HTML injection regression', () => {
    // REGRESSION (real bug): slide title/content/notes are text extracted from
    // an *untrusted* uploaded file (PPTX text runs, PDF text items, or raw HTML
    // body text). convertToHtml used to interpolate them directly into the
    // generated HTML with no escaping. A slide whose visible text happens to
    // contain "<script>...</script>" (or any markup) was therefore re-parsed
    // as live HTML/script the moment a victim opened the downloaded .html file
    // in a browser - a stored HTML/script injection via a crafted presentation.
    const maliciousContent = {
        slides: [
            {
                slideNumber: 1,
                title: '<script>window.__pwned = true;</script>',
                content: ['<img src=x onerror=alert(1)>', 'safe & sound'],
                notes: '</div><script>alert(2)</script>'
            }
        ]
    };

    test('slide title is HTML-escaped, not injected as markup', async () => {
        const blob = PresentationConverter.convertToHtml(maliciousContent, 'Deck');
        const html = await blob.text();
        assert.doesNotMatch(html, /<script>window\.__pwned/);
        assert.match(html, /&lt;script&gt;window\.__pwned = true;&lt;\/script&gt;/);
    });

    test('slide content items are HTML-escaped', async () => {
        const blob = PresentationConverter.convertToHtml(maliciousContent, 'Deck');
        const html = await blob.text();
        assert.doesNotMatch(html, /<img src=x onerror=alert\(1\)>/);
        assert.match(html, /&lt;img src=x onerror=alert\(1\)&gt;/);
        assert.match(html, /safe &amp; sound/);
    });

    test('slide notes are HTML-escaped', async () => {
        const blob = PresentationConverter.convertToHtml(maliciousContent, 'Deck');
        const html = await blob.text();
        assert.doesNotMatch(html, /<\/div><script>alert\(2\)/);
        assert.match(html, /&lt;\/div&gt;&lt;script&gt;alert\(2\)&lt;\/script&gt;/);
    });

    test('presentation title itself is escaped', async () => {
        const blob = PresentationConverter.convertToHtml({ slides: [] }, '<b>Evil</b> Title');
        const html = await blob.text();
        assert.doesNotMatch(html, /<b>Evil<\/b> Title/);
        assert.match(html, /&lt;b&gt;Evil&lt;\/b&gt; Title/);
    });

    test('fallback (no .slides) branch also escapes title and content', async () => {
        const blob = PresentationConverter.convertToHtml('<script>alert(3)</script>', '<i>Title</i>');
        const html = await blob.text();
        assert.doesNotMatch(html, /<script>alert\(3\)/);
        assert.doesNotMatch(html, /<i>Title<\/i>/);
    });
});

describe('PresentationConverter includeNotes option', () => {
    const content = {
        slides: [
            { slideNumber: 1, title: 'S1', content: ['hello'], notes: 'secret note' }
        ]
    };

    test('convertToHtml includes notes by default', async () => {
        const html = await PresentationConverter.convertToHtml(content, 'T').text();
        assert.match(html, /secret note/);
    });

    test('convertToHtml omits notes when includeNotes is false', async () => {
        const html = await PresentationConverter.convertToHtml(content, 'T', { includeNotes: false }).text();
        assert.doesNotMatch(html, /secret note/);
    });

    test('convertToText omits notes when includeNotes is false', async () => {
        const text = await PresentationConverter.convertToText(content, 'T', { includeNotes: false }).text();
        assert.doesNotMatch(text, /secret note/);
    });

    test('convertToText includes notes by default', async () => {
        const text = await PresentationConverter.convertToText(content, 'T').text();
        assert.match(text, /secret note/);
    });
});

describe('PresentationConverter.escapeHtml', () => {
    test('escapes all five special characters', () => {
        assert.equal(
            PresentationConverter.escapeHtml(`<a href="x">Tom & Jerry's</a>`),
            '&lt;a href=&quot;x&quot;&gt;Tom &amp; Jerry&#39;s&lt;/a&gt;'
        );
    });

    test('returns empty string for null/undefined', () => {
        assert.equal(PresentationConverter.escapeHtml(null), '');
        assert.equal(PresentationConverter.escapeHtml(undefined), '');
    });
});

describe('PresentationConverter.getFileType / isValidPresentationFile', () => {
    test('extracts lowercase extension without the dot', () => {
        assert.equal(PresentationConverter.getFileType({ name: 'Deck.PPTX' }), 'pptx');
    });

    test('accepts a recognized extension', () => {
        assert.equal(PresentationConverter.isValidPresentationFile({ name: 'deck.pptx', type: '' }), true);
    });

    test('rejects an unsupported file', () => {
        assert.equal(PresentationConverter.isValidPresentationFile({ name: 'song.mp3', type: 'audio/mpeg' }), false);
    });
});

describe('PresentationConverter.formatFileSize', () => {
    test('formats bytes into human-readable units', () => {
        assert.equal(PresentationConverter.formatFileSize(0), '0 Bytes');
        assert.equal(PresentationConverter.formatFileSize(1024), '1 KB');
        assert.equal(PresentationConverter.formatFileSize(1536), '1.5 KB');
    });
});

// Unit tests for assets/js/converters/presentation.js
// Run with: node --test tests/unit
'use strict';

const { test, describe, before, after } = require('node:test');
const assert = require('node:assert/strict');

const PresentationConverter = require('../../assets/js/converters/presentation.js');

// parseSlideXML() is browser-only (needs the real DOMParser). Node has no
// DOMParser, so this is a minimal stand-in - just enough of the DOM surface
// (tagName, textContent, querySelectorAll with comma-separated tag-name
// selectors incl. the escaped-colon "a\\:t" form the real code uses) to
// exercise the REAL parsing/grouping logic in parseSlideXML against
// realistic PPTX slide XML fixtures. It is not a general XML/CSS engine.
function createXmlDomParserStub() {
    function decodeEntities(text) {
        return text
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
            .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
            .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
            .replace(/&amp;/g, '&');
    }

    class MiniElement {
        constructor(tagName) {
            this.tagName = tagName;
            this.childNodes = [];
        }
        get textContent() {
            let out = '';
            for (const c of this.childNodes) {
                out += typeof c === 'string' ? c : c.textContent;
            }
            return out;
        }
        querySelectorAll(selector) {
            const names = selector.split(',').map(s => s.trim().replace(/\\:/g, ':'));
            const results = [];
            const walk = (node) => {
                for (const c of node.childNodes) {
                    if (typeof c !== 'string') {
                        if (names.includes(c.tagName)) results.push(c);
                        walk(c);
                    }
                }
            };
            walk(this);
            return results;
        }
    }

    function parseXML(xmlString) {
        const root = new MiniElement('#document');
        const stack = [root];
        const tagRe = /<(\/?)([A-Za-z_][\w.\-:]*)((?:\s+[^<>]*?)?)(\/?)>|<!--[\s\S]*?-->|<\?[\s\S]*?\?>/g;
        let lastIndex = 0;
        let m;
        while ((m = tagRe.exec(xmlString)) !== null) {
            const text = xmlString.slice(lastIndex, m.index);
            if (text) {
                const decoded = decodeEntities(text);
                if (decoded) stack[stack.length - 1].childNodes.push(decoded);
            }
            lastIndex = tagRe.lastIndex;
            if (m[0].startsWith('<!--') || m[0].startsWith('<?')) continue;
            const [, closing, tagName, , selfClose] = m;
            if (closing) {
                stack.pop();
            } else {
                const node = new MiniElement(tagName);
                stack[stack.length - 1].childNodes.push(node);
                if (!selfClose) stack.push(node);
            }
        }
        return root;
    }

    return class DOMParserStub {
        parseFromString(xmlString) {
            return parseXML(xmlString);
        }
    };
}

describe('PresentationConverter.parseSlideXML - PPTX 多 run 段落重組（R2 新真 bug）', () => {
    // REGRESSION (real bug): PowerPoint splits ONE visual line into multiple
    // <a:r><a:t> runs whenever formatting changes mid-line (bold word,
    // color change, hyperlink boundary) - this happens on essentially any
    // real-world deck. parseSlideXML used to select every <a:t> directly and
    // treat each run as its own line: the title became just the first run
    // (truncated, e.g. "Quarterly " instead of "Quarterly Revenue Report")
    // and the remaining runs of that same line turned into bogus extra
    // bullet points indistinguishable from real separate paragraphs. Fixed
    // by grouping runs by their enclosing <a:p> paragraph before treating
    // each paragraph as one line.
    const originalDOMParser = global.DOMParser;

    before(() => {
        global.DOMParser = createXmlDomParserStub();
    });

    after(() => {
        global.DOMParser = originalDOMParser;
    });

    const multiRunTitleXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:txBody>
      <a:p>
        <a:r><a:rPr b="1"/><a:t>Quarterly </a:t></a:r>
        <a:r><a:rPr/><a:t>Revenue </a:t></a:r>
        <a:r><a:rPr i="1"/><a:t>Report</a:t></a:r>
      </a:p>
    </p:txBody></p:sp>
    <p:sp><p:txBody>
      <a:p><a:r><a:t>Bullet one</a:t></a:r></a:p>
      <a:p><a:r><a:t>Bullet two</a:t></a:r></a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:sld>`;

    test('runs split by mid-line formatting are reassembled into one title, not truncated', () => {
        const slide = PresentationConverter.parseSlideXML(multiRunTitleXml, 1);
        assert.equal(slide.title, 'Quarterly Revenue Report');
    });

    test('reassembled title runs are not leaked as bogus extra bullet points', () => {
        const slide = PresentationConverter.parseSlideXML(multiRunTitleXml, 1);
        assert.deepEqual(slide.content, ['Bullet one', 'Bullet two']);
    });

    test('genuinely separate paragraphs remain separate content lines', () => {
        const xml = `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree><p:sp><p:txBody>
            <a:p><a:r><a:t>Short title</a:t></a:r></a:p>
            <a:p><a:r><a:t>First point</a:t></a:r></a:p>
            <a:p><a:r><a:t>Second point</a:t></a:r></a:p>
          </p:txBody></p:sp></p:spTree></p:cSld>
        </p:sld>`;
        const slide = PresentationConverter.parseSlideXML(xml, 2);
        assert.equal(slide.title, 'Short title');
        assert.deepEqual(slide.content, ['First point', 'Second point']);
    });

    test('falls back to flat run extraction when there is no <a:p> paragraph wrapper', () => {
        const xml = `<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:t>Loose run one</a:t>
            <a:t>Loose run two</a:t>
        </root>`;
        const slide = PresentationConverter.parseSlideXML(xml, 3);
        assert.equal(slide.title, 'Loose run one');
        assert.deepEqual(slide.content, ['Loose run two']);
    });
});

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

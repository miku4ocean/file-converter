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

// Minimal File-like stub for extractors (only .text()/.arrayBuffer() needed).
function fileStub(name, content) {
    return {
        name,
        text: async () => content,
        arrayBuffer: async () => new TextEncoder().encode(content).buffer
    };
}

describe('DocumentConverter.convertToFormat - round-trip 保真（R2 回歸）', () => {
    // REGRESSION (real bug): convertToText/convertToHtml/convertToMarkdown 在
    // class 內各有兩份定義，後者靜默覆蓋前者；而 convertToFormat 以「第三個
    // 位置參數」傳入 originalHtml / originalMarkdown，被存活版本當成 options
    // 物件丟棄。結果：MD→MD 與 HTML→HTML 轉換輸出的是被剝光格式的重生成
    // 內容（標題層級/粗體/連結/原始標籤全滅），使用者不易察覺。
    test('md → md keeps the original markdown verbatim', async () => {
        const md = '# Title\n\n**bold** and [link](https://example.com)\n';
        const extracted = await DocumentConverter.extractFromMarkdown(fileStub('note.md', md));
        const blob = await DocumentConverter.convertToFormat(extracted, 'md');
        assert.equal(await blob.text(), md);
    });

    test('html → html keeps the original html verbatim', async () => {
        const extracted = {
            title: 'T',
            content: 'Text',
            originalHtml: '<!DOCTYPE html><html><body><em>Text</em></body></html>'
        };
        const blob = await DocumentConverter.convertToFormat(extracted, 'html');
        assert.equal(await blob.text(), extracted.originalHtml);
    });
});

describe('DocumentConverter.convertToHtml - 換行轉 <br>（R2 回歸）', () => {
    // REGRESSION (real bug): 先把 \n 換成 <br> 再對整串做 escapeXml，
    // <br> 被跳脫成 &lt;br&gt;，輸出頁面上出現字面 "<br>" 文字而非換行。
    test('single newlines become real <br> tags, not escaped literal text', async () => {
        const blob = DocumentConverter.convertToHtml('line1\nline2', 'T');
        const text = await blob.text();
        assert.match(text, /line1<br>line2/);
        assert.doesNotMatch(text, /&lt;br&gt;/);
    });

    test('user content containing a literal <br> string is still escaped (guard)', async () => {
        const blob = DocumentConverter.convertToHtml('a <br> b', 'T');
        const text = await blob.text();
        assert.match(text, /a &lt;br&gt; b/);
    });
});

describe('DocumentConverter.extractFromMarkdown - code fence 順序（R2 回歸）', () => {
    // REGRESSION (real bug): inline-code 的 `...` 規則在 fenced code block
    // 規則「之前」執行，先把 ``` 圍欄吃掉兩個反引號，導致 code block 永遠
    // 移除不掉 —— 輸出殘留破碎反引號、語言標記與本應移除的程式碼。
    test('fenced code blocks are removed cleanly', async () => {
        const md = 'Before\n\n```js\nconst secret = 1;\n```\n\nAfter';
        const extracted = await DocumentConverter.extractFromMarkdown(fileStub('c.md', md));
        assert.match(extracted.content, /Before/);
        assert.match(extracted.content, /After/);
        assert.doesNotMatch(extracted.content, /`/);
        assert.doesNotMatch(extracted.content, /const secret/);
    });

    test('inline code keeps its text content (guard)', async () => {
        const extracted = await DocumentConverter.extractFromMarkdown(fileStub('c.md', 'use `foo()` here'));
        assert.match(extracted.content, /use foo\(\) here/);
    });
});

describe('DocumentConverter - RTF Unicode 保真（R2 回歸）', () => {
    // REGRESSION (real bug, 兩面):
    // 1. escapeRtf 對非 ASCII 字元原樣輸出 —— 產生的 \ansi RTF 在
    //    Word/WordPad 開啟時 CJK 全成亂碼（RTF 規範要求 \uN? 跳脫）。
    // 2. extractFromRtf 的指令剝除 regex 把 \uN 連同數字整段刪掉，
    //    CJK 內容只剩 fallback '?'；\'xx 十六進位跳脫也原樣殘留。
    test('escapeRtf encodes CJK as \\uN? sequences', () => {
        assert.equal(DocumentConverter.escapeRtf('中'), '\\u20013?');
    });

    test('escapeRtf encodes an emoji as a surrogate pair of \\uN?', () => {
        assert.equal(DocumentConverter.escapeRtf('😀'), '\\u-10179?\\u-8704?');
    });

    test('extractFromRtf decodes \\uN? escapes back to CJK', async () => {
        const rtf = "{\\rtf1\\ansi{\\fonttbl{\\f0 Arial;}}\\f0\\fs24 \\u20013?\\u25991? OK}";
        const extracted = await DocumentConverter.extractFromRtf(fileStub('t.rtf', rtf));
        assert.match(extracted.content, /中文/);
        assert.match(extracted.content, /OK/);
    });

    test("extractFromRtf decodes \\'xx hex escapes (latin-1 range)", async () => {
        const rtf = "{\\rtf1\\ansi caf\\'e9}";
        const extracted = await DocumentConverter.extractFromRtf(fileStub('t.rtf', rtf));
        assert.match(extracted.content, /café/);
    });

    // REGRESSION (real bug): 字型表等 destination group 的內容（"Times New
    // Roman;" 之類）被當正文洩漏進萃取結果。
    test('extractFromRtf does not leak font-table names into the text', async () => {
        const rtf = "{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Times New Roman;}}\\f0\\fs24 body}";
        const extracted = await DocumentConverter.extractFromRtf(fileStub('t.rtf', rtf));
        assert.doesNotMatch(extracted.content, /Times New Roman/);
        assert.match(extracted.content, /body/);
    });

    test('txt→rtf→txt round-trip preserves CJK and emoji', async () => {
        const original = '中文段落一\n\n中文段落二 😀';
        const blob = DocumentConverter.convertToRtf(original, '標題');
        const rtfText = await blob.text();
        const extracted = await DocumentConverter.extractFromRtf(fileStub('r.rtf', rtfText));
        assert.match(extracted.content, /標題/);
        assert.match(extracted.content, /中文段落一/);
        assert.match(extracted.content, /中文段落二 😀/);
        assert.doesNotMatch(extracted.content, /Times New Roman/);
    });
});

describe('DocumentConverter.extractFromText - UTF-16 編碼偵測（R2 回歸）', () => {
    // REGRESSION (real bug): extractFromText 用 file.text()（固定 UTF-8），
    // Windows 記事本「Unicode」（UTF-16LE）存檔的 txt 解出來是 NUL 亂碼。
    test('decodes a UTF-16LE (BOM) txt file', async () => {
        const payload = Buffer.from('中文內容', 'utf16le');
        const bytes = new Uint8Array(2 + payload.length);
        bytes.set([0xFF, 0xFE], 0);
        bytes.set(payload, 2);
        const file = new File([bytes], 'u16.txt', { type: 'text/plain' });
        const extracted = await DocumentConverter.extractFromText(file);
        assert.equal(extracted.content, '中文內容');
    });
});

describe('DocumentConverter.extractFromHtml - UTF-16 編碼偵測（R2 第二輪新真 bug）', () => {
    // REGRESSION (real bug): extractFromHtml 用 file.text()（固定 UTF-8）,
    // 與 extractFromText/extractFromMarkdown 同款 bug 但當時未一併修——用
    // Windows 記事本「Unicode」(UTF-16LE) 存的 .html 檔會解出夾滿 NUL 的
    // 亂碼且無任何報錯，DOMParser 再對亂碼 "解析" 只會產生更破損的結果。
    // 改用既有的 decodeTextFile()（依 BOM 偵測）取代 file.text()。

    // extractFromHtml needs a real DOMParser; Node has none, so this is a
    // minimal HTML-tag-tree stand-in - just enough (querySelector(All),
    // textContent, remove(), body/documentElement) to exercise the REAL
    // extractFromHtml logic against small hand-built HTML fixtures.
    function createHtmlDomParserStub() {
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

        const voidTags = new Set(['br', 'hr', 'img', 'meta', 'link', 'input']);

        class MiniElement {
            constructor(tagName) {
                this.tagName = tagName;
                this.childNodes = [];
                this.parentNode = null;
            }
            get textContent() {
                let out = '';
                for (const c of this.childNodes) out += typeof c === 'string' ? c : c.textContent;
                return out;
            }
            get body() { return this.querySelector('body'); }
            get documentElement() {
                return this.childNodes.find(c => typeof c !== 'string') || null;
            }
            querySelectorAll(selector) {
                const names = selector.split(',').map(s => s.trim());
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
            querySelector(selector) {
                return this.querySelectorAll(selector)[0] || null;
            }
            remove() {
                if (this.parentNode) {
                    const idx = this.parentNode.childNodes.indexOf(this);
                    if (idx !== -1) this.parentNode.childNodes.splice(idx, 1);
                }
            }
        }

        function parseHTML(htmlString) {
            const root = new MiniElement('#document');
            const stack = [root];
            const tagRe = /<(\/?)([A-Za-z_][\w.\-:]*)((?:\s+[^<>]*?)?)(\/?)>|<!--[\s\S]*?-->|<\?[\s\S]*?\?>/g;
            let lastIndex = 0;
            let m;
            while ((m = tagRe.exec(htmlString)) !== null) {
                const text = htmlString.slice(lastIndex, m.index);
                if (text) {
                    const decoded = decodeEntities(text);
                    if (decoded) stack[stack.length - 1].childNodes.push(decoded);
                }
                lastIndex = tagRe.lastIndex;
                if (m[0].startsWith('<!--') || m[0].startsWith('<?')) continue;
                const closing = m[1];
                const tagName = m[2].toLowerCase();
                const selfClose = m[4];
                if (closing) {
                    for (let i = stack.length - 1; i > 0; i--) {
                        if (stack[i].tagName === tagName) {
                            stack.length = i;
                            break;
                        }
                    }
                } else {
                    const node = new MiniElement(tagName);
                    node.parentNode = stack[stack.length - 1];
                    stack[stack.length - 1].childNodes.push(node);
                    if (!selfClose && !voidTags.has(tagName)) stack.push(node);
                }
            }
            return root;
        }

        return class DOMParserStub {
            parseFromString(htmlString) {
                return parseHTML(htmlString);
            }
        };
    }

    const originalDOMParser = global.DOMParser;

    test('decodes a UTF-16LE (BOM) html file instead of producing NUL mojibake', async () => {
        global.DOMParser = createHtmlDomParserStub();
        try {
            const html = '<html><head><title>中文標題</title></head><body>內容測試</body></html>';
            const payload = Buffer.from(html, 'utf16le');
            const bytes = new Uint8Array(2 + payload.length);
            bytes.set([0xFF, 0xFE], 0);
            bytes.set(payload, 2);
            const file = new File([bytes], 'u16.html', { type: 'text/html' });

            const extracted = await DocumentConverter.extractFromHtml(file);
            assert.equal(extracted.title, '中文標題');
            assert.equal(extracted.content, '內容測試');
            assert.doesNotMatch(extracted.originalHtml, /\u0000/);
            assert.equal(extracted.originalHtml, html);
        } finally {
            global.DOMParser = originalDOMParser;
        }
    });

    test('still decodes plain UTF-8 html (guard)', async () => {
        global.DOMParser = createHtmlDomParserStub();
        try {
            const html = '<html><head><title>Plain</title></head><body>Hello</body></html>';
            const file = new File([html], 'plain.html', { type: 'text/html' });
            const extracted = await DocumentConverter.extractFromHtml(file);
            assert.equal(extracted.title, 'Plain');
            assert.equal(extracted.content, 'Hello');
        } finally {
            global.DOMParser = originalDOMParser;
        }
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

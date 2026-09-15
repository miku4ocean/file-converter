import { readFileSync, existsSync } from 'fs';
let pass = 0, fail = 0;
function test(name, fn) { try { fn(); pass++; console.log(`  ✓ ${name}`); } catch(e) { fail++; console.log(`  ✗ ${name}: ${e.message}`); } }
function assert(cond, msg) { if (!cond) throw new Error(msg); }
const candidates = ['index.html', 'src/index.html', 'public/index.html'];
let htmlFile = candidates.find(f => existsSync(f));
assert(htmlFile, 'No index.html found');
const html = readFileSync(htmlFile, 'utf8');
test('HTML file exists and is non-empty', () => { assert(html.length > 100, `Only ${html.length} chars`); });
test('Has DOCTYPE', () => { assert(html.toLowerCase().includes('<!doctype'), 'Missing DOCTYPE'); });
test('Has charset meta', () => { assert(html.includes('charset'), 'Missing charset'); });
test('Has title', () => { assert(/<title>[^<]+<\/title>/.test(html), 'Missing or empty title'); });
test('Has body content', () => { assert(/<body[\s>]/.test(html), 'Missing body tag'); });
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail > 0 ? 1 : 0);

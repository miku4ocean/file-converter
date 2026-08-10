// Unit tests for assets/js/converters/image.js
// Only the parts of ImageConverter that don't require a real DOM/Canvas
// (Image/canvas-based resizing needs a browser and is out of scope for a
// Node unit test) are covered here.
// Run with: node --test tests/unit
'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');

const ImageConverter = require('../../assets/js/converters/image.js');

describe('ImageConverter.suggestOutputFormat', () => {
    test('always prefers PNG when the source has transparency', () => {
        assert.equal(ImageConverter.suggestOutputFormat('jpeg', true), 'png');
        assert.equal(ImageConverter.suggestOutputFormat('bmp', true), 'png');
    });

    test('suggests jpeg for png/bmp without transparency', () => {
        assert.equal(ImageConverter.suggestOutputFormat('png', false), 'jpeg');
        assert.equal(ImageConverter.suggestOutputFormat('bmp', false), 'jpeg');
    });

    test('suggests png for gif to preserve quality', () => {
        assert.equal(ImageConverter.suggestOutputFormat('gif', false), 'png');
    });

    test('keeps webp as webp', () => {
        assert.equal(ImageConverter.suggestOutputFormat('webp', false), 'webp');
    });

    test('is case-insensitive on the input format', () => {
        assert.equal(ImageConverter.suggestOutputFormat('PNG', false), 'jpeg');
    });
});

describe('ImageConverter.isValidImageFile', () => {
    test('accepts a known image MIME type with nonzero size', () => {
        assert.equal(ImageConverter.isValidImageFile({ type: 'image/png', size: 100 }), true);
    });

    test('rejects a zero-byte file even with a valid MIME type', () => {
        assert.equal(ImageConverter.isValidImageFile({ type: 'image/png', size: 0 }), false);
    });

    test('rejects an unsupported MIME type', () => {
        assert.equal(ImageConverter.isValidImageFile({ type: 'application/pdf', size: 100 }), false);
    });
});

describe('ImageConverter.batchProcess', () => {
    test('collects per-file success results and reports progress', async () => {
        const files = [{ name: 'a.png' }, { name: 'b.png' }];
        const progressCalls = [];

        const results = await ImageConverter.batchProcess(
            files,
            async (file) => `processed-${file.name}`,
            (done, total, name) => progressCalls.push([done, total, name])
        );

        assert.equal(results.length, 2);
        assert.equal(results[0].success, true);
        assert.equal(results[0].result, 'processed-a.png');
        assert.deepEqual(progressCalls, [[1, 2, 'a.png'], [2, 2, 'b.png']]);
    });

    test('captures per-file failures without aborting the batch', async () => {
        const files = [{ name: 'ok.png' }, { name: 'bad.png' }, { name: 'ok2.png' }];

        const results = await ImageConverter.batchProcess(files, async (file) => {
            if (file.name === 'bad.png') throw new Error('boom');
            return 'ok';
        });

        assert.equal(results.length, 3);
        assert.equal(results[0].success, true);
        assert.equal(results[1].success, false);
        assert.equal(results[1].error.message, 'boom');
        assert.equal(results[2].success, true);
    });
});

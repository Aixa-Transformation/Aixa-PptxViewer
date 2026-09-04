// @vitest-environment happy-dom

import type { TextSegment } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import { getInlineEditorSelection } from './inline-selection-utils';

describe('getInlineEditorSelection live paragraph coordinates', () => {
	afterEach(() => {
		document.body.replaceChildren();
		window.getSelection()?.removeAllRanges();
	});

	it('identifies a newly inserted paragraph even when it reuses an old segment index', () => {
		const editor = document.createElement('div');
		editor.dataset.inlineEditor = '';
		editor.innerHTML = `
			<div data-pptx-paragraph data-pptx-list-seg-idx="0">
				<span data-seg-idx="0">&#8226; </span><span data-seg-idx="1">First</span>
			</div>
			<div data-pptx-paragraph>
				<span data-seg-idx="1">Newly typed paragraph</span>
			</div>`;
		document.body.append(editor);
		const segments: TextSegment[] = [
			{ text: '\u2022 ', style: {}, bulletInfo: { char: '\u2022' } },
			{ text: 'First', style: {} },
		];
		const insertedText = editor.querySelectorAll<HTMLElement>('[data-seg-idx="1"]')[1]
			.firstChild!;
		const range = document.createRange();
		range.setStart(insertedText, 0);
		range.setEnd(insertedText, insertedText.textContent?.length ?? 0);
		const selection = window.getSelection()!;
		selection.removeAllRanges();
		selection.addRange(range);

		expect(getInlineEditorSelection(segments)).toMatchObject({
			startSegIdx: 1,
			startOffset: 0,
			endSegIdx: 1,
			endOffset: 21,
			startParagraphIndex: 1,
			startParagraphOffset: 0,
			endParagraphIndex: 1,
			endParagraphOffset: 21,
			endAtParagraphStart: false,
		});
	});
});

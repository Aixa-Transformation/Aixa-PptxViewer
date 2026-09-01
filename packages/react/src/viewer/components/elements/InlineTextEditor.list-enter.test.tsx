// @vitest-environment happy-dom

import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { InlineTextEditor } from './InlineTextEditor';

function bulletSegments(): TextSegment[] {
	return [
		{ text: '\u2022 ', style: {}, bulletInfo: { char: '\u2022' } },
		{ text: 'First', style: {} },
		{ text: '\n', style: {}, isParagraphBreak: true },
		{ text: '\u2022 ', style: {}, bulletInfo: { char: '\u2022' } },
		{ text: 'Second', style: {} },
	];
}

function bulletElement(): PptxElement {
	const textSegments = bulletSegments();
	return {
		id: 'shape-1',
		type: 'shape',
		shapeType: 'rect',
		x: 0,
		y: 0,
		width: 400,
		height: 200,
		text: textSegments.map((segment) => segment.text).join(''),
		textStyle: { fontSize: 20 },
		textSegments,
		paragraphIndents: [
			{ marginLeft: 36, indent: -18 },
			{ marginLeft: 36, indent: -18 },
		],
	} as PptxElement;
}

function numberedElement(): PptxElement {
	const textSegments: TextSegment[] = [
		{
			text: '1. ',
			style: {},
			bulletInfo: { autoNumType: 'arabicPeriod', autoNumStartAt: 1, paragraphIndex: 0 },
		},
		{ text: 'First', style: {} },
		{ text: '', style: {} },
		{ text: '\n', style: {}, isParagraphBreak: true },
		{
			text: '2. ',
			style: {},
			bulletInfo: { autoNumType: 'arabicPeriod', autoNumStartAt: 1, paragraphIndex: 1 },
		},
		{ text: '', style: {} },
		{ text: 'Second', style: {} },
		{ text: '', style: {} },
		{ text: '\n', style: {}, isParagraphBreak: true },
		{
			text: '3. ',
			style: {},
			bulletInfo: { autoNumType: 'arabicPeriod', autoNumStartAt: 1, paragraphIndex: 2 },
		},
		{ text: '', style: {} },
		{ text: 'Third', style: {} },
	];
	return {
		...bulletElement(),
		text: textSegments.map((segment) => segment.text).join(''),
		textSegments,
		paragraphIndents: [
			{ marginLeft: 53.76, indent: -53.76 },
			{ marginLeft: 53.76, indent: -53.76 },
			{ marginLeft: 53.76, indent: -53.76 },
		],
	} as PptxElement;
}

function suppliedDeckBulletElement(): PptxElement {
	const textSegments: TextSegment[] = [
		{ text: '\u2022 ', style: {}, bulletInfo: { char: '\u2022', fontFamily: 'Arial' } },
		{ text: 'First', style: {} },
		{ text: '', style: {} },
		{ text: '\n', style: {}, isParagraphBreak: true },
		{ text: '\u2022 ', style: {}, bulletInfo: { char: '\u2022', fontFamily: 'Arial' } },
		{ text: '', style: {} },
		{ text: 'Second', style: {} },
	];
	return {
		...bulletElement(),
		text: textSegments.map((segment) => segment.text).join(''),
		textSegments,
		paragraphIndents: [
			{ marginLeft: 36.48, indent: -36.48 },
			{ marginLeft: 36.48, indent: -36.48 },
		],
	} as PptxElement;
}

describe('InlineTextEditor list Enter', () => {
	let host: HTMLDivElement;

	beforeEach(() => {
		globalThis.IS_REACT_ACT_ENVIRONMENT = true;
		host = document.createElement('div');
		document.body.append(host);
	});

	afterEach(() => {
		host.remove();
		vi.restoreAllMocks();
	});

	it('clones the authored bullet position and discards an untouched continuation on blur', () => {
		const onCommit = vi.fn();
		const onEditChange = vi.fn();
		const root = createRoot(host);
		act(() => {
			root.render(
				<InlineTextEditor
					initialText={'\u2022 First\n\u2022 Second'}
					spellCheck={false}
					textStyle={{}}
					textStyleRaw={{}}
					layoutStyle={{}}
					element={bulletElement()}
					onCommit={onCommit}
					onCancel={vi.fn()}
					onEditChange={onEditChange}
				/>,
			);
		});

		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const originalText = editor.innerText;
		const paragraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		const lastParagraph = paragraphs[1];
		const lastText = lastParagraph.querySelector<HTMLElement>('[data-seg-idx="4"]')!;
		// Reproduce imported runs whose stale marker-like box width used to be
		// cloned into the continuation and wrap typing every few characters.
		lastText.style.display = 'inline-block';
		lastText.style.width = '24px';
		lastText.style.minWidth = '24px';
		const textNode = lastText.firstChild!;
		const selection = window.getSelection()!;
		const range = document.createRange();
		range.setStart(textNode, textNode.textContent?.length ?? 0);
		range.collapse(true);
		selection.removeAllRanges();
		selection.addRange(range);

		act(() => {
			editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		});

		const continuedParagraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		expect(continuedParagraphs).toHaveLength(3);
		const continuation = continuedParagraphs[2];
		expect(continuation.style.display).toBe('flex');
		expect(continuation.style.marginLeft).toBe('0px');
		expect(continuation.style.textIndent).toBe('0px');
		const originalMarker = lastParagraph.querySelector<HTMLElement>('[data-seg-idx="3"]')!;
		const continuedMarker = continuation.querySelector<HTMLElement>('[data-seg-idx="3"]')!;
		const continuedContent = continuation.querySelector<HTMLElement>(
			'[data-pptx-list-continuation-content="true"]',
		)!;
		expect(continuedMarker.style.width).toBe(originalMarker.style.width);
		expect(continuedMarker.textContent).toBe('\u2022 ');
		expect(continuedContent.style.display).toBe('block');
		expect(continuedContent.style.flex).toBe('1 1 auto');
		expect(continuedContent.style.width).toBe('auto');
		expect(continuedContent.style.minWidth).toBe('0');
		expect(continuedContent.style.overflowWrap).toBe('break-word');
		expect(continuedContent.style.wordBreak).toBe('normal');
		expect(continuedContent.querySelector('br')).toBeNull();
		expect(continuedContent.textContent).toBe('\u200B');
		expect(onEditChange).not.toHaveBeenCalled();

		// The first real input removes the inline caret placeholder without
		// introducing a line break or saving the zero-width character.
		continuedContent.firstChild!.textContent = '\u200BContinued bullet text';
		act(() => {
			editor.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
		});
		expect(continuedContent.textContent).toBe('Continued bullet text');
		expect(continuedContent.querySelector('br')).toBeNull();

		// Restore an untouched placeholder to exercise the empty-continuation
		// cleanup path independently from the input assertion above.
		continuedContent.textContent = '\u200B';
		onEditChange.mockClear();

		act(() => {
			editor.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
		});

		expect(editor.querySelectorAll('[data-pptx-paragraph]')).toHaveLength(2);
		expect(onCommit).toHaveBeenCalledWith(undefined, originalText, 1);
		expect(onEditChange).toHaveBeenLastCalledWith(originalText);

		act(() => root.unmount());
	});

	it('keeps the shape fixed and shrinks text gradually while typing', () => {
		const onCommit = vi.fn();
		const root = createRoot(host);
		const element = {
			...bulletElement(),
				height: 100,
			textStyle: { fontSize: 20, autoFit: false, autoFitMode: 'none' as const },
		} as PptxElement;

		act(() => {
			root.render(
				<div data-pptx-element='true' style={{ height: element.height }}>
					<InlineTextEditor
						initialText={element.text ?? ''}
						spellCheck={false}
						textStyle={{ fontSize: 20 }}
						textStyleRaw={element.textStyle}
						layoutStyle={{}}
						element={element}
						onCommit={onCommit}
						onCancel={vi.fn()}
						onEditChange={vi.fn()}
					/>
				</div>,
			);
		});

		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const shape = editor.closest<HTMLElement>('[data-pptx-element="true"]')!;
		expect(editor.style.overflowWrap).toBe('break-word');
		expect(editor.style.wordBreak).toBe('normal');
		let requiredHeight = 240;
		Object.defineProperty(editor, 'scrollHeight', {
			configurable: true,
			get: () => {
				const currentFontSize = Number.parseFloat(editor.style.fontSize) || 20;
				return Math.ceil(requiredHeight * (currentFontSize / 20));
			},
		});
		Object.defineProperty(editor, 'clientHeight', { configurable: true, value: 100 });
		Object.defineProperty(editor, 'scrollWidth', { configurable: true, value: 400 });
		Object.defineProperty(editor, 'clientWidth', { configurable: true, value: 400 });

		act(() => {
			editor.innerText = 'Enough typed text to require substantially more vertical space';
			editor.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
		});

		const liveScale = Number.parseFloat(editor.style.fontSize) / 20;
		expect(shape.style.height).toBe('100px');
		expect(liveScale).toBeCloseTo(18 / 20, 4);

		// Each overflowing insertion decreases the effective font by one 2-unit
		// PowerPoint-style step rather than jumping directly to a miniature size.
		act(() => {
			editor.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
		});
		expect(Number.parseFloat(editor.style.fontSize) / 20).toBeCloseTo(16 / 20, 4);

		// A later insertion must not make the font jump larger if a subsequent
		// layout measurement temporarily reports more available room.
		requiredHeight = 100;
		act(() => {
			editor.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
		});
		expect(Number.parseFloat(editor.style.fontSize) / 20).toBeCloseTo(16 / 20, 4);

		// Deleting content grows back by the same single step.
		act(() => {
			editor.dispatchEvent(
				new InputEvent('input', { bubbles: true, inputType: 'deleteContentBackward' }),
			);
		});
		expect(Number.parseFloat(editor.style.fontSize) / 20).toBeCloseTo(18 / 20, 4);

		act(() => {
			editor.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
		});
		expect(onCommit).toHaveBeenCalledWith(undefined, editor.innerText, expect.any(Number));
		expect(onCommit.mock.calls[0][2]).toBeCloseTo(18 / 20, 4);

		act(() => root.unmount());
	});

	it('splits a bullet in the middle into a full-width continuation paragraph', () => {
		const root = createRoot(host);
		act(() => {
			root.render(
				<InlineTextEditor
					initialText={'\u2022 First\n\u2022 Second'}
					spellCheck={false}
					textStyle={{}}
					textStyleRaw={{}}
					layoutStyle={{ display: 'flex', flexDirection: 'column' }}
					element={bulletElement()}
					onCommit={vi.fn()}
					onCancel={vi.fn()}
					onEditChange={vi.fn()}
				/>,
			);
		});

		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const paragraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		const secondText = paragraphs[1].querySelector<HTMLElement>('[data-seg-idx="4"]')!;
		const textNode = secondText.firstChild!;
		const selection = window.getSelection()!;
		const range = document.createRange();
		range.setStart(textNode, 3);
		range.collapse(true);
		selection.removeAllRanges();
		selection.addRange(range);

		act(() => {
			editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		});

		const splitParagraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		expect(splitParagraphs).toHaveLength(3);
		expect(secondText.textContent).toBe('Sec');
		const continuation = splitParagraphs[2];
		expect(continuation.style.alignSelf).toBe('stretch');
		expect(continuation.style.minWidth).toBe('0');
		expect(continuation.style.maxWidth).toBe('100%');
		expect(continuation.textContent).toBe('\u2022 \u200Bond');
		const continuationContent = continuation.querySelector<HTMLElement>(
			'[data-pptx-list-continuation-content="true"]',
		)!;
		expect(continuation.style.display).toBe('flex');
		expect(continuation.style.textIndent).toBe('0px');
		expect(continuationContent.style.display).toBe('block');
		expect(continuationContent.style.flex).toBe('1 1 auto');
		expect(continuationContent.style.width).toBe('auto');
		expect(continuationContent.style.whiteSpace).toBe('pre-wrap');

		act(() => root.unmount());
	});

	it('keeps the caret in a full-width empty bullet when splitting at paragraph start', () => {
		const root = createRoot(host);
		act(() => {
			root.render(
				<InlineTextEditor
					initialText={'\u2022 First\n\u2022 Second'}
					spellCheck={false}
					textStyle={{}}
					textStyleRaw={{}}
					layoutStyle={{ display: 'flex', flexDirection: 'column' }}
					element={bulletElement()}
					onCommit={vi.fn()}
					onCancel={vi.fn()}
					onEditChange={vi.fn()}
				/>,
			);
		});

		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const paragraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		const secondText = paragraphs[1].querySelector<HTMLElement>('[data-seg-idx="4"]')!;
		const selection = window.getSelection()!;
		const range = document.createRange();
		range.setStart(secondText.firstChild!, 0);
		range.collapse(true);
		selection.removeAllRanges();
		selection.addRange(range);

		act(() => {
			editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		});

		const splitParagraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		expect(splitParagraphs).toHaveLength(3);
		const emptyBullet = splitParagraphs[1];
		const movedBullet = splitParagraphs[2];
		expect(emptyBullet.textContent).toBe('\u2022 \u200B');
		expect(movedBullet.textContent).toBe('\u2022 Second');
		const editableContent = emptyBullet.querySelector<HTMLElement>(
			'[data-pptx-list-continuation-content="true"]',
		)!;
		expect(emptyBullet.style.display).toBe('flex');
		expect(emptyBullet.style.marginLeft).toBe('0px');
		expect(emptyBullet.style.textIndent).toBe('0px');
		expect(editableContent.style.display).toBe('block');
		expect(editableContent.style.flex).toBe('1 1 auto');
		expect(editableContent.style.width).toBe('auto');
		expect(editableContent.style.maxWidth).toBe('100%');

		act(() => root.unmount());
	});

	it('continues numbered paragraphs without creating a narrow native block', () => {
		const onCommit = vi.fn();
		const root = createRoot(host);
		act(() => {
			root.render(
				<InlineTextEditor
					initialText={'1. First\n2. Second\n3. Third'}
					spellCheck={false}
					textStyle={{}}
					textStyleRaw={{}}
					layoutStyle={{ display: 'flex', flexDirection: 'column' }}
					element={numberedElement()}
					onCommit={onCommit}
					onCancel={vi.fn()}
					onEditChange={vi.fn()}
				/>,
			);
		});

		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const firstParagraph = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]')[0];
		const firstText = firstParagraph.querySelector<HTMLElement>('[data-seg-idx="1"]')!;
		const range = document.createRange();
		range.setStart(firstText.firstChild!, firstText.textContent?.length ?? 0);
		range.collapse(true);
		const selection = window.getSelection()!;
		selection.removeAllRanges();
		selection.addRange(range);

		act(() => {
			editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		});

		const paragraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		expect(paragraphs).toHaveLength(4);
		const continuation = paragraphs[1];
		expect(continuation.textContent).toBe('2. \u200B');
		expect(paragraphs[2].textContent).toBe('3. Second');
		expect(paragraphs[3].textContent).toBe('4. Third');
		expect(continuation.style.alignSelf).toBe('stretch');
		expect(continuation.style.minWidth).toBe('0');
		expect(continuation.style.maxWidth).toBe('100%');
		expect(continuation.style.display).toBe('flex');
		expect(continuation.style.marginLeft).toBe('0px');
		expect(continuation.style.textIndent).toBe('0px');
		const content = continuation.querySelector<HTMLElement>(
			'[data-pptx-list-continuation-content="true"]',
		)!;
		expect(content.style.display).toBe('block');
		expect(content.style.flex).toBe('1 1 auto');
		expect(content.style.width).toBe('auto');

		// The untouched numbered continuation is removed cleanly even though its
		// marker segment is index zero in the source paragraph.
		act(() => {
			editor.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
		});
		expect(editor.querySelectorAll('[data-pptx-paragraph]')).toHaveLength(3);
		expect(onCommit).toHaveBeenCalled();

		act(() => root.unmount());
	});

	it('commits an inserted second numbered item without a phantom paragraph', () => {
		const onCommit = vi.fn();
		const root = createRoot(host);
		act(() => {
			root.render(
				<InlineTextEditor
					initialText={'1. First\n2. Second\n3. Third'}
					spellCheck={false}
					textStyle={{}}
					textStyleRaw={{}}
					layoutStyle={{ display: 'flex', flexDirection: 'column' }}
					element={numberedElement()}
					onCommit={onCommit}
					onCancel={vi.fn()}
					onEditChange={vi.fn()}
				/>,
			);
		});

		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const firstText = editor.querySelector<HTMLElement>('[data-seg-idx="1"]')!;
		const range = document.createRange();
		range.setStart(firstText.firstChild!, firstText.textContent?.length ?? 0);
		range.collapse(true);
		const selection = window.getSelection()!;
		selection.removeAllRanges();
		selection.addRange(range);

		act(() => {
			editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		});

		const insertedContent = editor.querySelector<HTMLElement>(
			'[data-pptx-list-continuation-content="true"]',
		)!;
		insertedContent.firstChild!.textContent = '\u200BInserted item';
		act(() => {
			editor.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
			editor.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
		});

		expect(onCommit).toHaveBeenCalledWith(
			undefined,
			'1. First\n2. Inserted item\n3. Second\n4. Third',
			expect.any(Number),
		);
		expect(onCommit.mock.calls[0][1]).not.toContain('\n2. \n');

		act(() => root.unmount());
	});

	it('handles the supplied deck topology when Enter is pressed before a numbered item', () => {
		const root = createRoot(host);
		act(() => {
			root.render(
				<InlineTextEditor
					initialText={'1. First\n2. Second\n3. Third'}
					spellCheck={false}
					textStyle={{}}
					textStyleRaw={{}}
					layoutStyle={{ display: 'flex', flexDirection: 'column' }}
					element={numberedElement()}
					onCommit={vi.fn()}
					onCancel={vi.fn()}
					onEditChange={vi.fn()}
				/>,
			);
		});

		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const originalParagraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		const thirdParagraph = originalParagraphs[2];
		// The supplied PPTX has an empty formatting run between each number marker
		// and visible text. A click at the text start can therefore resolve to the
		// paragraph boundary immediately after that empty run.
		const range = document.createRange();
		range.setStart(thirdParagraph, 2);
		range.collapse(true);
		const selection = window.getSelection()!;
		selection.removeAllRanges();
		selection.addRange(range);

		act(() => {
			editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		});

		const paragraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		expect(paragraphs).toHaveLength(4);
		const emptyThird = paragraphs[2];
		const movedFourth = paragraphs[3];
		expect(emptyThird.textContent).toBe('3. \u200B');
		expect(movedFourth.textContent).toBe('4. Third');
		expect(emptyThird.style.alignSelf).toBe('stretch');
		expect(emptyThird.style.display).toBe('flex');
		expect(emptyThird.style.width).toBe('100%');
		expect(emptyThird.style.marginLeft).toBe('0px');
		expect(emptyThird.style.textIndent).toBe('0px');
		const content = emptyThird.querySelector<HTMLElement>(
			'[data-pptx-list-continuation-content="true"]',
		)!;
		expect(content.style.display).toBe('block');
		expect(content.style.flex).toBe('1 1 auto');
		expect(content.style.width).toBe('auto');
		expect(content.querySelector('div')).toBeNull();

		content.firstChild!.textContent = '\u200Btyped text remains on a normal line';
		act(() => {
			editor.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
		});
		expect(content.textContent).toBe('typed text remains on a normal line');
		expect(content.style.width).toBe('auto');

		act(() => root.unmount());
	});

	it('handles the supplied deck empty formatting run before bullet text', () => {
		const root = createRoot(host);
		act(() => {
			root.render(
				<InlineTextEditor
					initialText={'\u2022 First\n\u2022 Second'}
					spellCheck={false}
					textStyle={{}}
					textStyleRaw={{}}
					layoutStyle={{ display: 'flex', flexDirection: 'column' }}
					element={suppliedDeckBulletElement()}
					onCommit={vi.fn()}
					onCancel={vi.fn()}
					onEditChange={vi.fn()}
				/>,
			);
		});

		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const secondParagraph = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]')[1];
		const range = document.createRange();
		range.setStart(secondParagraph, 2);
		range.collapse(true);
		const selection = window.getSelection()!;
		selection.removeAllRanges();
		selection.addRange(range);

		act(() => {
			editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		});

		const paragraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		expect(paragraphs).toHaveLength(3);
		expect(paragraphs[1].textContent).toBe('\u2022 \u200B');
		expect(paragraphs[2].textContent).toBe('\u2022 Second');
		const content = paragraphs[1].querySelector<HTMLElement>(
			'[data-pptx-list-continuation-content="true"]',
		)!;
		expect(paragraphs[1].style.display).toBe('flex');
		expect(paragraphs[1].style.marginLeft).toBe('0px');
		expect(paragraphs[1].style.textIndent).toBe('0px');
		expect(content.style.flex).toBe('1 1 auto');
		expect(content.style.width).toBe('auto');
		expect(content.style.wordBreak).toBe('normal');
		expect(content.style.overflowWrap).toBe('break-word');

		act(() => root.unmount());
	});

	it('keeps Shift+Enter as a soft line break instead of creating a list item', () => {
		const root = createRoot(host);
		act(() => {
			root.render(
				<InlineTextEditor
					initialText={'1. First\n2. Second\n3. Third'}
					spellCheck={false}
					textStyle={{}}
					textStyleRaw={{}}
					layoutStyle={{ display: 'flex', flexDirection: 'column' }}
					element={numberedElement()}
					onCommit={vi.fn()}
					onCancel={vi.fn()}
					onEditChange={vi.fn()}
				/>,
			);
		});

		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const thirdText = editor.querySelector<HTMLElement>('[data-seg-idx="11"]')!;
		const range = document.createRange();
		range.setStart(thirdText.firstChild!, 2);
		range.collapse(true);
		const selection = window.getSelection()!;
		selection.removeAllRanges();
		selection.addRange(range);
		const event = new KeyboardEvent('keydown', { key: 'Enter', shiftKey: true, bubbles: true });

		act(() => {
			editor.dispatchEvent(event);
		});

		expect(event.defaultPrevented).toBe(false);
		expect(editor.querySelectorAll('[data-pptx-paragraph]')).toHaveLength(3);

		act(() => root.unmount());
	});

	it('exits a list after Enter is pressed twice on an empty item', () => {
		const root = createRoot(host);
		act(() => {
			root.render(
				<InlineTextEditor
					initialText={'1. First\n2. Second\n3. Third'}
					spellCheck={false}
					textStyle={{}}
					textStyleRaw={{}}
					layoutStyle={{ display: 'flex', flexDirection: 'column' }}
					element={numberedElement()}
					onCommit={vi.fn()}
					onCancel={vi.fn()}
					onEditChange={vi.fn()}
				/>,
			);
		});

		const editor = host.querySelector<HTMLElement>('[data-inline-editor]')!;
		const secondText = editor.querySelector<HTMLElement>('[data-seg-idx="6"]')!;
		const selection = window.getSelection()!;
		const range = document.createRange();
		range.setStart(secondText.firstChild!, secondText.textContent?.length ?? 0);
		range.collapse(true);
		selection.removeAllRanges();
		selection.addRange(range);

		act(() => {
			editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
			editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		});

		const paragraphs = editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]');
		expect(paragraphs).toHaveLength(4);
		const plainParagraph = paragraphs[2];
		expect(plainParagraph.hasAttribute('data-pptx-list-seg-idx')).toBe(false);
		expect(plainParagraph.textContent).toBe('\u200B');
		expect(plainParagraph.style.marginLeft).toBe('');
		expect(plainParagraph.style.textIndent).toBe('');
		expect(plainParagraph.style.display).toBe('block');
		expect(paragraphs[3].textContent).toBe('4. Third');

		act(() => root.unmount());
	});
});

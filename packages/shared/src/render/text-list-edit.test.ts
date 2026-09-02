import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { applyParagraphListType, getElementListType } from './text-list-edit';

function textElement(textSegments: TextSegment[], listType?: 'bullet' | 'numbered' | 'none') {
	return {
		id: 'shape-1',
		type: 'shape',
		x: 0,
		y: 0,
		width: 400,
		height: 200,
		text: textSegments.map((segment) => segment.text).join(''),
		textStyle: listType ? { listType } : {},
		textSegments,
	} as PptxElement;
}

describe('applyParagraphListType', () => {
	it('creates structural bullet markers for every paragraph', () => {
		const result = applyParagraphListType({ text: 'First\nSecond', listType: 'bullet' });

		expect(result.text).toBe('\u2022 First\n\u2022 Second');
		expect(result.textSegments[0]).toMatchObject({
			text: '\u2022 ',
			bulletInfo: { char: '\u2022' },
			style: { listType: 'bullet' },
		});
		expect(result.textSegments[3]).toMatchObject({
			text: '\u2022 ',
			bulletInfo: { char: '\u2022' },
		});
	});

	it('creates a contiguous PowerPoint auto-numbered sequence', () => {
		const result = applyParagraphListType({ text: 'First\nSecond', listType: 'numbered' });

		expect(result.text).toBe('1. First\n2. Second');
		expect(result.textSegments[0].bulletInfo).toStrictEqual({
			autoNumType: 'arabicPeriod',
			autoNumStartAt: 1,
			paragraphIndex: 0,
		});
		expect(result.textSegments[3].bulletInfo?.paragraphIndex).toBe(1);
	});

	it('turns an imported bullet list off without leaving marker text behind', () => {
		const result = applyParagraphListType({
			textSegments: [
				{ text: '\u2022 ', style: {}, bulletInfo: { char: '\u2022' } },
				{ text: 'First', style: { bold: true } },
				{ text: '\n', style: {}, isParagraphBreak: true },
				{ text: '\u2022 ', style: {}, bulletInfo: { char: '\u2022' } },
				{ text: 'Second', style: {} },
			],
			listType: 'none',
		});

		expect(result.text).toBe('First\nSecond');
		expect(result.textSegments[0]).toMatchObject({ text: '', bulletInfo: { none: true } });
		expect(result.textSegments[1]).toMatchObject({ text: 'First', style: { bold: true } });
		expect(result.textSegments[3]).toMatchObject({ text: '', bulletInfo: { none: true } });
	});

	it('applies the list only to paragraphs intersecting the inline selection', () => {
		const segments: TextSegment[] = [
			{ text: 'First', style: {} },
			{ text: '\n', style: {}, isParagraphBreak: true },
			{ text: 'Second', style: { italic: true } },
		];
		const result = applyParagraphListType({
			textSegments: segments,
			listType: 'bullet',
			selection: { startSegIdx: 2, startOffset: 0, endSegIdx: 2, endOffset: 6 },
		});

		expect(result.text).toBe('First\n\u2022 Second');
		expect(result.textSegments[0].bulletInfo).toBeUndefined();
		expect(result.textSegments[2].bulletInfo).toStrictEqual({ char: '\u2022' });
		expect(result.selection).toStrictEqual({
			startSegIdx: 3,
			startOffset: 0,
			endSegIdx: 3,
			endOffset: 6,
		});
	});

	it('does not include the next paragraph when the selection ends at its text start', () => {
		const segments: TextSegment[] = [
			{ text: 'First', style: {} },
			{ text: '\n', style: {}, isParagraphBreak: true },
			{ text: '', style: {} },
			{ text: 'Second', style: {} },
			{ text: '\n', style: {}, isParagraphBreak: true },
			{ text: 'Third', style: {} },
		];
		const result = applyParagraphListType({
			textSegments: segments,
			listType: 'numbered',
			selection: { startSegIdx: 0, startOffset: 1, endSegIdx: 3, endOffset: 0 },
		});

		expect(result.text).toBe('1. First\nSecond\nThird');
		expect(result.textSegments.filter((segment) => segment.bulletInfo?.autoNumType)).toHaveLength(1);
	});

	it('formats every partially selected paragraph as one contiguous numbered list', () => {
		const segments: TextSegment[] = [
			{ text: 'Alpha ', style: { bold: true } },
			{ text: 'one', style: {} },
			{ text: '\n', style: {}, isParagraphBreak: true },
			{ text: 'Beta ', style: {} },
			{ text: 'two', style: { italic: true } },
			{ text: '\n', style: {}, isParagraphBreak: true },
			{ text: 'Gamma', style: {} },
		];
		const result = applyParagraphListType({
			textSegments: segments,
			listType: 'numbered',
			selection: { startSegIdx: 1, startOffset: 1, endSegIdx: 4, endOffset: 2 },
		});

		expect(result.text).toBe('1. Alpha one\n2. Beta two\nGamma');
		expect(
			result.textSegments
				.filter((segment) => segment.bulletInfo?.autoNumType)
				.map((segment) => segment.bulletInfo?.paragraphIndex),
		).toStrictEqual([0, 1]);
		expect(result.textSegments.find((segment) => segment.text === 'Alpha ')?.style.bold).toBe(true);
		expect(result.textSegments.find((segment) => segment.text === 'two')?.style.italic).toBe(true);
	});

	it('keeps paragraph geometry on the new structural marker', () => {
		const result = applyParagraphListType({
			textSegments: [
				{
					text: 'Indented',
					style: { fontSize: 20 },
					paragraphLevel: 2,
					paragraphProperties: { paragraphMarginLeft: 48, paragraphIndent: -18 },
				},
			],
			listType: 'bullet',
		});

		expect(result.textSegments[0]).toMatchObject({
			paragraphLevel: 2,
			paragraphProperties: { paragraphMarginLeft: 48, paragraphIndent: -18 },
		});
		expect(result.textSegments[1].paragraphProperties).toBeUndefined();
	});
});

describe('getElementListType', () => {
	it('recognizes imported bullet and numbered list structures', () => {
		expect(
			getElementListType(
				textElement([
					{ text: '\u2022 ', style: {}, bulletInfo: { char: '\u2022' } },
					{ text: 'Item', style: {} },
				]),
			),
		).toBe('bullet');
		expect(
			getElementListType(
				textElement([
					{
						text: '1. ',
						style: {},
						bulletInfo: { autoNumType: 'arabicPeriod', paragraphIndex: 0 },
					},
					{ text: 'Item', style: {} },
				]),
			),
		).toBe('numbered');
	});

	it('reports mixed paragraph list states', () => {
		expect(
			getElementListType(
				textElement([
					{ text: '\u2022 ', style: {}, bulletInfo: { char: '\u2022' } },
					{ text: 'First', style: {} },
					{ text: '\n', style: {}, isParagraphBreak: true },
					{ text: 'Second', style: {} },
				]),
			),
		).toBe('mixed');
	});
});

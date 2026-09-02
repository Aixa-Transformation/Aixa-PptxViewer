import type { PptxElement, TextStyle, BulletInfo } from 'pptx-viewer-core';
import { describe, it, expect } from 'vitest';

import type { ParagraphEntry } from './text-animation';
import {
	resolveParagraphRtl,
	resolveParagraphAlign,
	resolveCssTextAlign,
	renderTextSegments,
} from './text-paragraph-render';
import { renderPictureBullet } from './text-segment-helpers';
import { renderSingleSegment } from './text-segment-render';

// ── Helpers ─────────────────────────────────────────────────────────────

/** Create a minimal ParagraphEntry for unit-testing helpers. */
function makeParagraphEntry(style: TextStyle, text = 'test'): ParagraphEntry {
	return { segment: { text, style }, globalIndex: 0 };
}

/** Create a minimal text element with given textStyle overrides. */
function makeTextElement(
	textStyleOverrides: Record<string, unknown> = {},
	extras: Record<string, unknown> = {},
): PptxElement {
	return {
		id: 'el-bidi',
		type: 'text',
		x: 0,
		y: 0,
		width: 400,
		height: 200,
		text: 'Hello',
		textStyle: { ...textStyleOverrides },
		...extras,
	} as unknown as PptxElement;
}

// =======================================================================
// 1. resolveParagraphRtl
// =======================================================================

describe('resolveParagraphRtl', () => {
	it('returns segment-level RTL when a segment explicitly sets rtl=true', () => {
		const entries = [makeParagraphEntry({ rtl: true })];
		expect(resolveParagraphRtl(entries, undefined)).toBeTruthy();
	});

	it('returns segment-level LTR when a segment explicitly sets rtl=false', () => {
		const entries = [makeParagraphEntry({ rtl: false })];
		expect(resolveParagraphRtl(entries, undefined)).toBeFalsy();
	});

	it('falls back to element-level RTL when no segment sets direction', () => {
		const entries = [makeParagraphEntry({})];
		expect(resolveParagraphRtl(entries, true)).toBeTruthy();
	});

	it('returns undefined when neither segments nor element set direction', () => {
		const entries = [makeParagraphEntry({})];
		expect(resolveParagraphRtl(entries, undefined)).toBeUndefined();
	});

	it("uses the first segment's direction when multiple segments differ", () => {
		const entries = [makeParagraphEntry({ rtl: true }), makeParagraphEntry({ rtl: false })];
		// Second entry should get globalIndex = 1 for realistic data
		entries[1].globalIndex = 1;
		expect(resolveParagraphRtl(entries, undefined)).toBeTruthy();
	});
});

// =======================================================================
// 2. resolveParagraphAlign
// =======================================================================

describe('resolveParagraphAlign', () => {
	it('returns segment-level alignment when present', () => {
		const entries = [makeParagraphEntry({ align: 'center' })];
		expect(resolveParagraphAlign(entries, undefined)).toBe('center');
	});

	it('falls back to element-level alignment', () => {
		const entries = [makeParagraphEntry({})];
		expect(resolveParagraphAlign(entries, 'right')).toBe('right');
	});

	it('returns undefined when no alignment is set anywhere', () => {
		const entries = [makeParagraphEntry({})];
		expect(resolveParagraphAlign(entries, undefined)).toBeUndefined();
	});
});

// =======================================================================
// 3. resolveCssTextAlign
// =======================================================================

describe('resolveCssTextAlign', () => {
	it("returns 'right' for RTL paragraph with no explicit alignment", () => {
		expect(resolveCssTextAlign(undefined, true)).toBe('right');
	});

	it('returns undefined for LTR paragraph with no explicit alignment', () => {
		expect(resolveCssTextAlign(undefined, false)).toBeUndefined();
	});

	it("respects explicit 'center' alignment in RTL paragraph", () => {
		expect(resolveCssTextAlign('center', true)).toBe('center');
	});

	it("respects explicit 'left' alignment in RTL paragraph (override)", () => {
		expect(resolveCssTextAlign('left', true)).toBe('left');
	});

	it("maps 'justLow' to 'justify'", () => {
		expect(resolveCssTextAlign('justLow', false)).toBe('justify');
	});

	it("maps 'dist' to 'justify'", () => {
		expect(resolveCssTextAlign('dist', true)).toBe('justify');
	});

	it("maps 'thaiDist' to 'justify'", () => {
		expect(resolveCssTextAlign('thaiDist', false)).toBe('justify');
	});

	it("returns 'justify' for explicit justify alignment", () => {
		expect(resolveCssTextAlign('justify', false)).toBe('justify');
	});
});

// =======================================================================
// 4. renderTextSegments: paragraph-level BiDi
// =======================================================================

describe('renderTextSegments: paragraph-level BiDi', () => {
	it('applies normAutofit scaling and line-spacing reduction to paragraph line boxes', () => {
		const el = makeTextElement(
			{ autoFitFontScale: 0.8, autoFitLineSpacingReduction: 0.2 },
			{
				textSegments: [
					{
						text: 'Scaled text',
						style: {
							fontSize: 40,
							align: 'center',
							lineSpacing: 1.5,
							paragraphSpacingBefore: 10,
						},
					},
					{
						text: '\n',
						style: {
							fontSize: 40,
							align: 'center',
							lineSpacing: 1.5,
							paragraphSpacingBefore: 10,
						},
					},
					{
						text: 'Second line',
						style: {
							fontSize: 40,
							align: 'center',
							lineSpacing: 1.5,
							paragraphSpacingBefore: 10,
						},
					},
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		expect(result[0].props.style.fontSize).toBe(32);
		expect(result[0].props.style.lineHeight).toBeCloseTo(1.6);
		expect(result[1].props.style.marginTop).toBe(8);
	});

	it('preserves an authored empty paragraph as a full line box', () => {
		const el = makeTextElement(
			{},
			{
				textSegments: [
					{ text: 'Subheading', style: { align: 'center', fontSize: 32 } },
					{ text: '\n', style: { align: 'center', fontSize: 32 } },
					{ text: '\n', style: { align: 'left', fontSize: 24 } },
					{ text: 'Body', style: { align: 'left', fontSize: 24 } },
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		expect(result).toHaveLength(3);
		expect(result[1].type).toBe('div');
		const emptyLineChildren = result[1].props.children as React.ReactElement[];
		expect(emptyLineChildren).toHaveLength(1);
		expect(emptyLineChildren[0].type).toBe('br');
	});

	it('collapses leading empty paragraphs like PowerPoint', () => {
		const el = makeTextElement(
			{},
			{
				textSegments: [
					{ text: '\n', style: { align: 'center', fontSize: 32 } },
					{ text: 'Title', style: { align: 'center', fontSize: 32 } },
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		expect(result).toHaveLength(1);
	});

	it('collapses a trailing empty bullet paragraph like PowerPoint', () => {
		const el = makeTextElement(
			{},
			{
				textSegments: [
					{ text: 'Last visible bullet', style: { fontSize: 22 } },
					{ text: '\n', style: { fontSize: 22 } },
					{
						text: '• ',
						style: { fontSize: 22 },
						bulletInfo: { char: '•' },
					},
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		expect(result).toHaveLength(1);
	});

	it('collapses an empty leading run followed by a soft break', () => {
		const el = makeTextElement(
			{},
			{
				textSegments: [
					{ text: '', style: { align: 'center', fontSize: 32 } },
					{ text: '\n', style: { align: 'center', fontSize: 32 } },
					{ text: 'Title', style: { align: 'center', fontSize: 32 } },
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		expect(result).toHaveLength(1);
	});

	it('sets direction:rtl and textAlign:right on a pure RTL paragraph', () => {
		const el = makeTextElement(
			{ rtl: true },
			{
				textSegments: [{ text: '\u0645\u0631\u062D\u0628\u0627', style: { rtl: true } }],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		expect(result).toBeDefined();
		expect(result.length).toBeGreaterThan(0);
		// Should render as a <div> wrapper because of RTL
		const para = result[0] as React.ReactElement;
		expect(para.props.style.direction).toBe('rtl');
		expect(para.props.style.textAlign).toBe('right');
		expect(para.props.style.unicodeBidi).toBe('embed');
	});

	it('sets direction:ltr on a pure LTR paragraph with explicit rtl=false', () => {
		const el = makeTextElement(
			{ rtl: false },
			{
				textSegments: [{ text: 'Hello', style: { rtl: false } }],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;
		expect(para.props.style.direction).toBe('ltr');
		expect(para.props.style.unicodeBidi).toBe('embed');
	});

	it('does not set direction when no RTL flag is present', () => {
		const el = makeTextElement(
			{},
			{
				textSegments: [{ text: 'Hello', style: {} }],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		// Without RTL or other triggers, should use Fragment (no wrapper div)
		const para = result[0] as React.ReactElement;
		// React.Fragment doesn't have style props
		expect(para.props.style).toBeUndefined();
	});

	it('swaps marginLeft to marginRight for RTL paragraph with indent', () => {
		const el = makeTextElement(
			{ rtl: true },
			{
				textSegments: [{ text: '\u0645\u0631\u062D\u0628\u0627', style: { rtl: true } }],
				paragraphIndents: [{ level: 0, marginLeft: 30, indent: -10 }],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;
		expect(para.props.style.marginRight).toBe(30);
		expect(para.props.style.marginLeft).toBeUndefined();
	});

	it('keeps marginLeft for LTR paragraph with indent', () => {
		const el = makeTextElement(
			{ rtl: false },
			{
				textSegments: [{ text: 'Hello', style: { rtl: false } }],
				paragraphIndents: [{ level: 0, marginLeft: 30, indent: -10 }],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;
		expect(para.props.style.marginLeft).toBe(30);
		expect(para.props.style.marginRight).toBeUndefined();
	});

	it('suppresses first-paragraph edge spacing when spcFirstLastPara is omitted', () => {
		const el = makeTextElement(
			{ paragraphSpacingBefore: 13.333, align: 'left' },
			{ textSegments: [{ text: 'Subtitle', style: { align: 'left' } }] },
		);
		const rendered = renderTextSegments(el, '#000000') as React.ReactElement[];
		const para = rendered[0];

		expect(para.props.style.marginTop).toBeUndefined();
	});

	it('applies explicit center alignment to RTL paragraph', () => {
		const el = makeTextElement(
			{ rtl: true, align: 'center' },
			{
				textSegments: [
					{ text: '\u0645\u0631\u062D\u0628\u0627', style: { rtl: true, align: 'center' } },
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;
		expect(para.props.style.direction).toBe('rtl');
		expect(para.props.style.textAlign).toBe('center');
	});
});

// =======================================================================
// 5. renderSingleSegment: run-level BiDi
// =======================================================================

describe('renderSingleSegment: run-level BiDi override', () => {
	const baseElement = makeTextElement({});
	const fallbackColor = '#000000';

	it('applies bidi-override when run direction differs from paragraph (LTR run in RTL para)', () => {
		const segment = {
			text: 'Hello',
			style: { rtl: false } as TextStyle,
		};
		const result = renderSingleSegment(
			baseElement,
			segment,
			0,
			fallbackColor,
			undefined,
			undefined,
			undefined,
			undefined,
			/* paragraphRtl */ true,
		) as React.ReactElement;
		expect(result.props.style.direction).toBe('ltr');
		expect(result.props.style.unicodeBidi).toBe('bidi-override');
	});

	it('applies bidi-override when run direction differs from paragraph (RTL run in LTR para)', () => {
		const segment = {
			text: '\u0645\u0631\u062D\u0628\u0627',
			style: { rtl: true } as TextStyle,
		};
		const result = renderSingleSegment(
			baseElement,
			segment,
			0,
			fallbackColor,
			undefined,
			undefined,
			undefined,
			undefined,
			/* paragraphRtl */ false,
		) as React.ReactElement;
		expect(result.props.style.direction).toBe('rtl');
		expect(result.props.style.unicodeBidi).toBe('bidi-override');
	});

	it('applies embed when run direction matches paragraph direction', () => {
		const segment = {
			text: '\u0645\u0631\u062D\u0628\u0627',
			style: { rtl: true } as TextStyle,
		};
		const result = renderSingleSegment(
			baseElement,
			segment,
			0,
			fallbackColor,
			undefined,
			undefined,
			undefined,
			undefined,
			/* paragraphRtl */ true,
		) as React.ReactElement;
		expect(result.props.style.direction).toBe('rtl');
		expect(result.props.style.unicodeBidi).toBe('embed');
	});

	it('does not set BiDi properties when run has no explicit direction', () => {
		const segment = {
			text: 'neutral',
			style: {} as TextStyle,
		};
		const result = renderSingleSegment(
			baseElement,
			segment,
			0,
			fallbackColor,
			undefined,
			undefined,
			undefined,
			undefined,
			/* paragraphRtl */ true,
		) as React.ReactElement;
		expect(result.props.style.direction).toBeUndefined();
		expect(result.props.style.unicodeBidi).toBeUndefined();
	});
});

// =======================================================================
// 6. Mixed RTL/LTR runs in a paragraph
// =======================================================================

describe('renderTextSegments: mixed RTL/LTR runs', () => {
	it('renders mixed RTL and LTR runs in an RTL paragraph', () => {
		const el = makeTextElement(
			{ rtl: true },
			{
				textSegments: [
					{ text: '\u0645\u0631\u062D\u0628\u0627 ', style: { rtl: true } },
					{ text: 'OpenAI', style: { rtl: false } },
					{ text: ' \u0639\u0627\u0644\u0645', style: { rtl: true } },
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		// Should be a single paragraph div
		expect(result).toHaveLength(1);
		const para = result[0] as React.ReactElement;
		expect(para.props.style.direction).toBe('rtl');
		expect(para.props.style.textAlign).toBe('right');
	});

	it('renders mixed LTR and RTL runs in an LTR paragraph', () => {
		const el = makeTextElement(
			{ rtl: false },
			{
				textSegments: [
					{ text: 'Hello ', style: { rtl: false } },
					{ text: '\u0645\u0631\u062D\u0628\u0627', style: { rtl: true } },
					{ text: ' world', style: { rtl: false } },
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;
		expect(para.props.style.direction).toBe('ltr');
	});
});

// =======================================================================
// 7. Embedded numbers in RTL text
// =======================================================================

describe('renderTextSegments: embedded numbers in RTL text', () => {
	it('uses embed unicode-bidi for RTL paragraph (numbers render LTR naturally)', () => {
		const el = makeTextElement(
			{ rtl: true },
			{
				textSegments: [
					{ text: '\u0627\u0644\u0631\u0642\u0645 123 \u0647\u0646\u0627', style: { rtl: true } },
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;
		// `embed` allows the Unicode Bidi Algorithm to handle numbers correctly:
		// numbers are inherently LTR and will render left-to-right within the RTL context
		expect(para.props.style.unicodeBidi).toBe('embed');
		expect(para.props.style.direction).toBe('rtl');
	});

	it('run-level embed reinforces number rendering in explicitly RTL runs', () => {
		const segment = {
			text: '\u0633\u0639\u0631 42.5$',
			style: { rtl: true } as TextStyle,
		};
		const result = renderSingleSegment(
			makeTextElement({}),
			segment,
			0,
			'#000',
			undefined,
			undefined,
			undefined,
			undefined,
			/* paragraphRtl */ true,
		) as React.ReactElement;
		// Run matches paragraph direction -> embed
		expect(result.props.style.unicodeBidi).toBe('embed');
	});
});

// =======================================================================
// 8. RTL bullet alignment
// =======================================================================

describe('bullet alignment in RTL paragraphs', () => {
	it('keeps an unbroken list word on the marker line after inline editing ends', () => {
		const el = makeTextElement(
			{},
			{
				textSegments: [
					{ text: '1. ', style: { fontSize: 20 }, bulletInfo: { autoNumType: 'arabicPeriod' } },
					{ text: 'averylongunbrokenwordthatmustwrap', style: { fontSize: 20 } },
				],
				paragraphIndents: [{ marginLeft: 36, indent: -18 }],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;

		expect(para.props.style.overflowWrap).toBe('anywhere');
		expect(para.props.style.wordBreak).toBe('normal');
		expect(para.props.style.minWidth).toBe(0);
		expect(para.props.style.maxWidth).toBe('100%');
	});

	it('uses marginInlineEnd on fallback character bullets (works for both LTR and RTL)', () => {
		const bulletInfo: BulletInfo = { imageRelId: 'rId5' };
		const result = renderPictureBullet('el-1', 0, bulletInfo, 16) as React.ReactElement;
		// Should use marginInlineEnd instead of marginRight
		expect(result.props.style.marginInlineEnd).toBe(4);
		expect(result.props.style.marginRight).toBeUndefined();
	});

	it('uses marginInlineEnd on image bullets', () => {
		const bulletInfo: BulletInfo = {
			imageDataUrl: 'data:image/png;base64,iVBOR',
			imageRelId: 'rId5',
		};
		const result = renderPictureBullet('el-1', 0, bulletInfo, 16) as React.ReactElement;
		expect(result.props.style.marginInlineEnd).toBe(4);
		expect(result.props.style.marginRight).toBeUndefined();
	});

	it('renders an RTL paragraph with a bullet and correct margin direction', () => {
		const el = makeTextElement(
			{ rtl: true },
			{
				textSegments: [
					{
						text: '\u2022 \u0645\u0631\u062D\u0628\u0627',
						style: { rtl: true },
						bulletInfo: { char: '\u2022' },
					},
				],
				paragraphIndents: [{ level: 0, marginLeft: 20, indent: -10 }],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;
		// RTL paragraph should swap margin to right
		expect(para.props.style.marginRight).toBe(20);
		expect(para.props.style.marginLeft).toBeUndefined();
		expect(para.props.style.direction).toBe('rtl');
	});
});

// =======================================================================
// 9. Nested BiDi: multiple paragraphs with different directions
// =======================================================================

describe('renderTextSegments: nested/multi-paragraph BiDi', () => {
	it('renders paragraphs with alternating directions', () => {
		const el = makeTextElement(
			{},
			{
				textSegments: [
					{ text: 'English paragraph', style: { rtl: false } },
					{ text: '\n', style: {} },
					{ text: '\u0641\u0642\u0631\u0629 \u0639\u0631\u0628\u064A\u0629', style: { rtl: true } },
					{ text: '\n', style: {} },
					{ text: 'Back to English', style: { rtl: false } },
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		expect(result).toHaveLength(3);

		// First paragraph: LTR
		const para1 = result[0] as React.ReactElement;
		expect(para1.props.style.direction).toBe('ltr');

		// Second paragraph: RTL
		const para2 = result[1] as React.ReactElement;
		expect(para2.props.style.direction).toBe('rtl');
		expect(para2.props.style.textAlign).toBe('right');

		// Third paragraph: LTR
		const para3 = result[2] as React.ReactElement;
		expect(para3.props.style.direction).toBe('ltr');
	});

	it('inherits element-level RTL when paragraph segments have no explicit direction', () => {
		const el = makeTextElement(
			{ rtl: true },
			{
				textSegments: [{ text: '\u0645\u0631\u062D\u0628\u0627', style: {} }],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;
		expect(para.props.style.direction).toBe('rtl');
		expect(para.props.style.textAlign).toBe('right');
	});
});

// =======================================================================
// 10. RTL paragraph with explicit left alignment (override)
// =======================================================================

describe('rTL paragraph alignment overrides', () => {
	it('allows explicit left alignment in an RTL paragraph', () => {
		const el = makeTextElement(
			{ rtl: true, align: 'left' },
			{
				textSegments: [
					{ text: '\u0645\u0631\u062D\u0628\u0627', style: { rtl: true, align: 'left' } },
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;
		expect(para.props.style.direction).toBe('rtl');
		expect(para.props.style.textAlign).toBe('left');
		expect(para.props['data-pptx-paragraph']).toBe(true);
		expect(para.props.style['--pptx-paragraph-align']).toBe('left');
	});

	it('applies justify alignment to RTL paragraph', () => {
		const el = makeTextElement(
			{ rtl: true, align: 'justify' },
			{
				textSegments: [
					{ text: '\u0645\u0631\u062D\u0628\u0627', style: { rtl: true, align: 'justify' } },
				],
			},
		);
		const result = renderTextSegments(el, '#000') as React.ReactElement[];
		const para = result[0] as React.ReactElement;
		expect(para.props.style.direction).toBe('rtl');
		expect(para.props.style.textAlign).toBe('justify');
	});
});

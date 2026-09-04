import { describe, expect, it, vi } from 'vitest';

import {
	buildNormalAutoFitTextStyle,
	findLargestFittingInlineTextScale,
	inlineTextFitMetricsFit,
	normalizeInlineTextAutoFitScale,
	resolveInlineTextCommitAutoFitScale,
} from './inline-text-autofit';

describe('inlineTextFitMetricsFit', () => {
	it('detects text that overflows above a middle-aligned flex box', () => {
		expect(
			inlineTextFitMetricsFit({
				scrollHeight: 100,
				clientHeight: 100,
				boxTop: 100,
				boxBottom: 200,
				contentTop: 82,
				contentBottom: 182,
			}),
		).toBe(false);
	});

	it('accepts content whose visual bounds remain inside the shape', () => {
		expect(
			inlineTextFitMetricsFit({
				// Browser flex bookkeeping may inflate this even though the measured
				// glyphs are fully inside the fixed PowerPoint text box.
				scrollHeight: 240,
				clientHeight: 100,
				boxTop: 100,
				boxBottom: 200,
				contentTop: 102,
				contentBottom: 198,
			}),
		).toBe(true);
	});

	it('uses scroll height when glyph bounds are unavailable', () => {
		expect(inlineTextFitMetricsFit({ scrollHeight: 140, clientHeight: 100 })).toBe(false);
		expect(inlineTextFitMetricsFit({ scrollHeight: 100, clientHeight: 100 })).toBe(true);
	});
});

describe('findLargestFittingInlineTextScale', () => {
	it('keeps the original font size when the content already fits', () => {
		const fits = vi.fn(() => true);
		expect(findLargestFittingInlineTextScale({ fits })).toBe(1);
		expect(fits).toHaveBeenCalledWith(1);
	});

	it('finds the largest scale that fits the fixed text box', () => {
		const scale = findLargestFittingInlineTextScale({ fits: (candidate) => candidate <= 0.625 });
		expect(scale).toBeCloseTo(0.625, 3);
	});

	it('uses the minimum scale when even the smallest text still overflows', () => {
		expect(findLargestFittingInlineTextScale({ fits: () => false, minScale: 0.2 })).toBe(0.2);
	});

	it('never grows above the current scale while text is being inserted', () => {
		const fits = vi.fn((candidate: number) => candidate <= 0.8);
		expect(findLargestFittingInlineTextScale({ fits, maxScale: 0.6 })).toBe(0.6);
		expect(Math.max(...fits.mock.calls.map(([candidate]) => candidate))).toBeLessThanOrEqual(0.6);
	});
});

describe('buildNormalAutoFitTextStyle', () => {
	it('converts an imported no-autofit shape to PowerPoint normAutofit', () => {
		expect(
			buildNormalAutoFitTextStyle(
				{ fontFamily: 'Aptos', autoFit: false, autoFitMode: 'none' },
				0.625,
			),
		).toStrictEqual({
			fontFamily: 'Aptos',
			autoFit: true,
			autoFitMode: 'normal',
			autoFitFontScale: 0.625,
		});
	});

	it('removes a stale scale when the content fits at full size', () => {
		expect(buildNormalAutoFitTextStyle({ autoFitFontScale: 0.5 }, 1)).toStrictEqual({
			autoFit: true,
			autoFitMode: 'normal',
		});
	});
});

describe('normalizeInlineTextAutoFitScale', () => {
	it('clamps and rounds scales for OOXML normAutofit serialization', () => {
		expect(normalizeInlineTextAutoFitScale(0.756789)).toBe(0.75679);
		expect(normalizeInlineTextAutoFitScale(2)).toBe(1);
		expect(normalizeInlineTextAutoFitScale(0)).toBe(0.05);
		expect(normalizeInlineTextAutoFitScale(Number.NaN)).toBe(1);
	});
});

describe('resolveInlineTextCommitAutoFitScale', () => {
	it('uses the final blur measurement when it is supplied', () => {
		expect(resolveInlineTextCommitAutoFitScale(0.62, 0.7)).toBe(0.62);
	});

	it('keeps the latest live scale when pointer-down commits before blur', () => {
		expect(resolveInlineTextCommitAutoFitScale(undefined, 0.7)).toBe(0.7);
	});

	it('does not invent a scale when neither commit path measured one', () => {
		expect(resolveInlineTextCommitAutoFitScale(undefined, null)).toBeUndefined();
	});
});

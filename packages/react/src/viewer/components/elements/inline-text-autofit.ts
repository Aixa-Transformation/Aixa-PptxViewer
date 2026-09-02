import type { TextStyle } from 'pptx-viewer-core';

export interface InlineTextScaleSearchInput {
	fits: (scale: number) => boolean;
	minScale?: number;
	maxScale?: number;
	iterations?: number;
}

export interface InlineTextFitMetrics {
	scrollHeight: number;
	clientHeight: number;
	boxTop?: number;
	boxBottom?: number;
	contentTop?: number;
	contentBottom?: number;
	tolerance?: number;
}

/**
 * Decide whether editable text is fully inside its fixed shape. `scrollHeight`
 * alone is insufficient for middle/bottom-aligned flex bodies: overflowing
 * content can extend above the box while the browser still reports identical
 * scroll/client heights. The optional visual bounds cover that case.
 */
export function inlineTextFitMetricsFit({
	scrollHeight,
	clientHeight,
	boxTop,
	boxBottom,
	contentTop,
	contentBottom,
	tolerance = 1,
}: InlineTextFitMetrics): boolean {
	if (scrollHeight > clientHeight + tolerance) {
		return false;
	}
	const hasVisualBounds =
		[boxTop, boxBottom, contentTop, contentBottom].every(
			(value) => typeof value === 'number' && Number.isFinite(value),
		) &&
		(boxBottom as number) > (boxTop as number) &&
		(contentBottom as number) > (contentTop as number);
	if (!hasVisualBounds) {
		return true;
	}
	return (
		(contentTop as number) >= (boxTop as number) - tolerance &&
		(contentBottom as number) <= (boxBottom as number) + tolerance
	);
}

/**
 * Find the largest font scale that fits inside the existing text-box geometry.
 * This mirrors PowerPoint's `a:normAutofit` behaviour: the shape stays fixed
 * and the text becomes smaller only when its content overflows.
 */
export function findLargestFittingInlineTextScale({
	fits,
	minScale = 0.1,
	maxScale = 1,
	iterations = 12,
}: InlineTextScaleSearchInput): number {
	const safeMin = Math.max(0.05, Math.min(1, minScale));
	const safeMax = Math.max(safeMin, Math.min(1, maxScale));
	if (fits(safeMax)) {
		return safeMax;
	}
	if (!fits(safeMin)) {
		return safeMin;
	}

	let low = safeMin;
	let high = safeMax;
	for (let index = 0; index < Math.max(1, iterations); index += 1) {
		const candidate = (low + high) / 2;
		if (fits(candidate)) {
			low = candidate;
		} else {
			high = candidate;
		}
	}
	return Math.round(low * 100000) / 100000;
}

/** Clamp a computed normAutofit scale to a serializer-safe percentage. */
export function normalizeInlineTextAutoFitScale(scale: number): number {
	if (!Number.isFinite(scale)) {
		return 1;
	}
	return Math.round(Math.max(0.05, Math.min(1, scale)) * 100000) / 100000;
}

/** Build the text-style state that serializes as DrawingML `a:normAutofit`. */
export function buildNormalAutoFitTextStyle(
	textStyle: TextStyle | undefined,
	scale: number,
): TextStyle {
	const normalizedScale = normalizeInlineTextAutoFitScale(scale);
	const result: TextStyle = {
		...(textStyle ?? {}),
		autoFit: true,
		autoFitMode: 'normal',
	};
	if (normalizedScale < 0.99999) {
		result.autoFitFontScale = normalizedScale;
	} else {
		delete result.autoFitFontScale;
	}
	return result;
}

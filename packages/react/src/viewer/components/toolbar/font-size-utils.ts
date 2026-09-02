import { hasTextProperties } from 'pptx-viewer-core';
import type { PptxElement } from 'pptx-viewer-core';

export function getToolbarAutoFitScale(
	element: PptxElement | null | undefined,
	liveAutoFitFontScale?: number | null,
): number {
	if (
		typeof liveAutoFitFontScale === 'number' &&
		Number.isFinite(liveAutoFitFontScale) &&
		liveAutoFitFontScale > 0 &&
		liveAutoFitFontScale <= 1
	) {
		return liveAutoFitFontScale;
	}
	if (!element || !hasTextProperties(element)) {
		return 1;
	}
	const scale = element.textStyle?.autoFitFontScale;
	return typeof scale === 'number' && Number.isFinite(scale) && scale > 0 && scale < 1
		? scale
		: 1;
}

export function getEffectiveToolbarFontSize(
	element: PptxElement | null | undefined,
	authoredFontSize: number,
	liveAutoFitFontScale?: number | null,
): number {
	// PowerPoint's ribbon presents whole-point sizes even when the authored OOXML
	// size divided by normAutofit produces a fractional internal value. Keep the
	// precise authored value for serialization, but never expose decimals in the
	// toolbar UI.
	return Math.max(
		1,
		Math.round(authoredFontSize * getToolbarAutoFitScale(element, liveAutoFitFontScale)),
	);
}

/** Convert the size chosen in the toolbar back to DrawingML's authored size. */
export function getAuthoredFontSizeForToolbar(
	element: PptxElement | null | undefined,
	effectiveFontSize: number,
	liveAutoFitFontScale?: number | null,
): number {
	return (
		Math.round(
			(effectiveFontSize / getToolbarAutoFitScale(element, liveAutoFitFontScale)) * 100,
		) / 100
	);
}

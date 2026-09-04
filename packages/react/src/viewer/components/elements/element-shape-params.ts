import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';

import { DEFAULT_FILL_COLOR, DEFAULT_STROKE_COLOR } from '../../constants';
import { buildCssGradientFromShapeStyle, normalizeHexColor } from '../../utils';
import { getReactSlideBackgroundStyle } from '../../utils/slide-background-style';

export function shapeParams(element: PptxElement, activeSlide?: PptxSlide) {
	const style = hasShapeProperties(element) ? element.shapeStyle : undefined;
	const strokeWidth = Math.max(0, style?.strokeWidth || 0);
	const strokeColor = normalizeHexColor(style?.strokeColor, DEFAULT_STROKE_COLOR);
	const backgroundFillStyle = style?.useBackgroundFill
		? getReactSlideBackgroundStyle(activeSlide)
		: undefined;
	const resolvedBackgroundColor = backgroundFillStyle?.backgroundColor;
	const fillColor = normalizeHexColor(
		typeof resolvedBackgroundColor === 'string' ? resolvedBackgroundColor : style?.fillColor,
		DEFAULT_FILL_COLOR,
	);
	const hasFill = style?.useBackgroundFill
		? true
		: (style?.fillColor !== undefined && style.fillColor !== 'transparent') ||
			Boolean(buildCssGradientFromShapeStyle(style) || style?.fillGradient) ||
			(style?.fillMode === 'pattern' && Boolean(style.fillPatternPreset));
	return {
		hf: hasFill,
		fc: fillColor,
		sw: strokeWidth,
		sc: strokeColor,
		backgroundFillStyle,
	} as const;
}

import type { TextStyle } from 'pptx-viewer-core';

import {
	findLargestFittingInlineTextScale,
	inlineTextFitMetricsFit,
	normalizeInlineTextAutoFitScale,
} from './inline-text-autofit';

/** Smallest font size (px) that shrink-to-fit is allowed to produce. */
const MIN_AUTO_FIT_FONT_SIZE = 8;

/**
 * PowerPoint shrinks text only for `a:normAutofit` / `a:spAutoFit` bodies.
 * `a:noAutofit` keeps the authored size and lets the text overflow, so the
 * viewer and the inline editor must both leave those shapes alone.
 */
export function isTextAutoFitEnabled(textStyle: TextStyle | undefined): boolean {
	if (!textStyle || textStyle.autoFitMode === 'none') {
		return false;
	}
	return (
		textStyle.autoFit === true ||
		textStyle.autoFitMode === 'normal' ||
		textStyle.autoFitMode === 'shrink'
	);
}

/** The `normAutofit` scale already baked into the React-rendered text metrics. */
export function getRenderedAutoFitScale(textStyle: TextStyle | undefined): number {
	const scale = textStyle?.autoFitFontScale;
	return typeof scale === 'number' && scale > 0 && scale < 1 ? scale : 1;
}

const METRIC_KEYS = {
	'font-size': { base: 'pptxAfBaseFontSize', applied: 'pptxAfAppliedFontSize' },
	'line-height': { base: 'pptxAfBaseLineHeight', applied: 'pptxAfAppliedLineHeight' },
	'margin-top': { base: 'pptxAfBaseMarginTop', applied: 'pptxAfAppliedMarginTop' },
	'margin-bottom': { base: 'pptxAfBaseMarginBottom', applied: 'pptxAfAppliedMarginBottom' },
} as const;

type MetricProperty = keyof typeof METRIC_KEYS;

interface ScaledMetric {
	property: MetricProperty;
	base: number;
	unit: string;
}

interface AutoFitNode {
	element: HTMLElement;
	metrics: ScaledMetric[];
}

function parseLength(value: string): { amount: number; unit: string } | undefined {
	const match = /^(-?\d*\.?\d+)(px|pt)$/.exec(value.trim());
	if (!match) {
		return undefined;
	}
	const amount = Number.parseFloat(match[1]);
	return Number.isFinite(amount) ? { amount, unit: match[2] } : undefined;
}

/**
 * Resolve the unscaled value of one inline metric. The value React rendered is
 * the source of truth; a cached base is reused only while the inline value is
 * still the one this module wrote, so any re-render (edited text, new
 * formatting, a newly committed autofit scale) invalidates it automatically.
 */
function resolveMetric(
	element: HTMLElement,
	property: MetricProperty,
	renderedScale: number,
): ScaledMetric | undefined {
	const keys = METRIC_KEYS[property];
	const inlineValue = element.style.getPropertyValue(property);
	const parsed = parseLength(inlineValue);
	if (!parsed) {
		return undefined;
	}
	if (element.dataset[keys.applied] === inlineValue) {
		const cached = Number(element.dataset[keys.base]);
		if (Number.isFinite(cached)) {
			return { property, base: cached, unit: parsed.unit };
		}
	}
	const base = parsed.amount / renderedScale;
	element.dataset[keys.base] = String(base);
	return { property, base, unit: parsed.unit };
}

function applyMetric(element: HTMLElement, metric: ScaledMetric, scale: number): void {
	const keys = METRIC_KEYS[metric.property];
	element.style.setProperty(metric.property, `${metric.base * scale}${metric.unit}`);
	element.dataset[keys.applied] = element.style.getPropertyValue(metric.property);
}

function collectNodes(
	container: HTMLElement,
	renderedScale: number,
): { nodes: AutoFitNode[]; minScale: number } {
	const elements: HTMLElement[] = [
		container,
		...container.querySelectorAll<HTMLElement>('[style]'),
	];
	const nodes: AutoFitNode[] = [];
	let smallestFontSize = Number.POSITIVE_INFINITY;
	for (const element of elements) {
		const metrics: ScaledMetric[] = [];
		// Only elements that carry an inline font size are rescaled; everything
		// else inherits, which preserves the authored size relationships.
		const fontSize = resolveMetric(element, 'font-size', renderedScale);
		if (fontSize && fontSize.base > 0) {
			metrics.push(fontSize);
			smallestFontSize = Math.min(smallestFontSize, fontSize.base);
		}
		// Exact DrawingML line spacing resolves to a fixed length, so it has to
		// shrink alongside the glyphs. Unitless line heights already follow.
		const lineHeight = resolveMetric(element, 'line-height', renderedScale);
		if (lineHeight && lineHeight.base > 0) {
			metrics.push(lineHeight);
		}
		if (element.hasAttribute('data-pptx-paragraph')) {
			for (const property of ['margin-top', 'margin-bottom'] as const) {
				const margin = resolveMetric(element, property, renderedScale);
				if (margin) {
					metrics.push(margin);
				}
			}
		}
		if (metrics.length > 0) {
			nodes.push({ element, metrics });
		}
	}
	const minScale = Number.isFinite(smallestFontSize)
		? Math.max(0.05, Math.min(1, MIN_AUTO_FIT_FONT_SIZE / smallestFontSize))
		: 0.1;
	return { nodes, minScale };
}

function measureFits(container: HTMLElement, bottomInset: number): boolean {
	let visualBounds: {
		boxTop?: number;
		boxBottom?: number;
		contentTop?: number;
		contentBottom?: number;
	} = {};
	try {
		const box = container.getBoundingClientRect();
		const contentRange = document.createRange();
		contentRange.selectNodeContents(container);
		const content = contentRange.getBoundingClientRect();
		// Slides render inside a CSS transform, so client rects are in scaled
		// pixels while paddings and insets are in layout pixels.
		const renderScale = container.offsetHeight > 0 ? box.height / container.offsetHeight : 1;
		const computed = window.getComputedStyle(container);
		const paddingTop = (Number.parseFloat(computed.paddingTop) || 0) * renderScale;
		const paddingBottom = (Number.parseFloat(computed.paddingBottom) || 0) * renderScale;
		visualBounds = {
			boxTop: box.top + paddingTop,
			boxBottom: box.bottom - paddingBottom - bottomInset * renderScale,
			contentTop: content.top,
			contentBottom: content.bottom,
		};
	} catch {
		// Older DOM shims do not expose Range#getBoundingClientRect; the
		// scroll-height comparison below remains a safe fallback there.
	}
	return inlineTextFitMetricsFit({
		scrollHeight: container.scrollHeight,
		clientHeight: Math.max(0, container.clientHeight - bottomInset),
		...visualBounds,
	});
}

export interface TextAutoFitInput {
	/** The rendered text-body element (view mode) or contentEditable (editing). */
	container: HTMLElement;
	/** Scale already present in the rendered markup (`autoFitFontScale`). */
	renderedScale: number;
	/** Extra bottom space (layout px) to keep free while fitting. */
	bottomInset?: number;
}

/**
 * Shrink (or restore) the text inside `container` so it fits the fixed shape,
 * mirroring PowerPoint's `a:normAutofit`. The geometry never changes; only the
 * font sizes, exact line heights and paragraph spacing are scaled.
 *
 * View mode and the inline editor both call this on the same markup, which is
 * what keeps the rendered size identical when editing starts and ends.
 *
 * @returns the applied scale, suitable for `autoFitFontScale`.
 */
export function fitTextToContainer({
	container,
	renderedScale,
	bottomInset = 0,
}: TextAutoFitInput): number {
	// A shape can briefly report a zero-size box while it is mounted, scrolled
	// into view or moved between responsive panes. Treating that transient state
	// as overflow would collapse the text before the real geometry exists.
	if (container.clientHeight <= 1 || container.clientWidth <= 1) {
		return renderedScale;
	}
	const divisor = renderedScale > 0 ? renderedScale : 1;
	const { nodes, minScale } = collectNodes(container, divisor);
	if (nodes.length === 0) {
		return 1;
	}
	const applyScale = (scale: number): void => {
		for (const node of nodes) {
			for (const metric of node.metrics) {
				applyMetric(node.element, metric, scale);
			}
		}
	};
	const fits = (scale: number): boolean => {
		applyScale(scale);
		return measureFits(container, bottomInset);
	};
	const scale = normalizeInlineTextAutoFitScale(
		findLargestFittingInlineTextScale({ fits, minScale, maxScale: 1 }),
	);
	applyScale(scale);
	return scale;
}

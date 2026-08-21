import type { PptxSmartArtDrawingShape, PptxSmartArtNode, SmartArtStyle } from 'pptx-viewer-core';
import { resolveDrawingShapeNodeId } from 'pptx-viewer-shared';
import React from 'react';

import {
	colour,
	contrastingTextColor,
	styleShadow,
	styleStroke,
} from '../../utils/smartart-helpers';
import { renderConnectorMarker } from '../../utils/connector-path';
import {
	fitFontSize,
	chevronPoints,
	smartArtNodeGroupProps,
	SmartArtNodeText,
} from './smartart-renderer-utils';

// ── Props ───────────────────────────────────────────────────────────────────

/** Props for the pre-computed drawing shape renderer. */
interface DrawingShapeRendererProps {
	/** Unique element ID for generating stable React keys. */
	elementId: string;
	/** Pre-computed drawing shapes from PowerPoint's layout engine. */
	shapes: PptxSmartArtDrawingShape[];
	/** Resolved SmartArt style (controls shadow, stroke). */
	style: SmartArtStyle;
	/** Resolved colour palette. */
	palette: string[];
	/**
	 * Model nodes, used to map a clicked drawing shape back to a node id for
	 * inline editing. When omitted, shapes are not tagged as editable.
	 */
	nodes?: readonly PptxSmartArtNode[];
	/**
	 * Per-node accessibility labels keyed by node id (from the shared
	 * `buildSmartArtA11y` view-model). When present, each shape with a resolvable
	 * node id gains `role="img"` + `aria-label` and an SVG `<title>`.
	 */
	nodeLabels?: Map<string, string>;
}

// ── Fill helpers ──────────────────────────────────────────────────────────────

/**
 * Build an SVG gradient `<def>` for a cached drawing shape that carries a
 * gradient fill, plus the `fill` reference (`url(#id)`) to paint it with.
 * Returns `null` when the shape has no gradient stops.
 */
function drawingShapeGradientDef(
	id: string,
	shape: PptxSmartArtDrawingShape,
): { def: React.ReactElement; ref: string } | null {
	const stops = shape.fillGradientStops;
	if (!stops || stops.length === 0) {
		return null;
	}
	const stopEls = stops.map((s, i) => (
		<stop
			key={`${id}-s${i}`}
			offset={`${Math.max(0, Math.min(100, s.position))}%`}
			stopColor={s.color}
			{...(s.opacity !== undefined ? { stopOpacity: s.opacity } : {})}
		/>
	));
	if (shape.fillGradientType === 'radial') {
		return {
			ref: `url(#${id})`,
			def: (
				<radialGradient id={id} key={id} cx='50%' cy='50%' r='50%'>
					{stopEls}
				</radialGradient>
			),
		};
	}
	// OOXML angle is clockwise from the +x axis with y pointing down, which
	// matches the SVG coordinate system, so sin/cos map directly.
	const rad = ((shape.fillGradientAngle ?? 0) * Math.PI) / 180;
	const dx = Math.cos(rad) / 2;
	const dy = Math.sin(rad) / 2;
	return {
		ref: `url(#${id})`,
		def: (
			<linearGradient
				id={id}
				key={id}
				x1={`${(0.5 - dx) * 100}%`}
				y1={`${(0.5 - dy) * 100}%`}
				x2={`${(0.5 + dx) * 100}%`}
				y2={`${(0.5 + dy) * 100}%`}
			>
				{stopEls}
			</linearGradient>
		),
	};
}

function drawingLineDashArray(token: string | undefined, strokeWidth: number): string | undefined {
	if (!token || token === 'solid') return undefined;
	if (token === 'dash' || token === 'sysDash') return `${4 * strokeWidth} ${3 * strokeWidth}`;
	if (token === 'dot' || token === 'sysDot') return `${strokeWidth} ${2 * strokeWidth}`;
	if (token.includes('DashDot')) return `${4 * strokeWidth} ${2 * strokeWidth} ${strokeWidth} ${2 * strokeWidth}`;
	return `${4 * strokeWidth} ${3 * strokeWidth}`;
}

/** PowerPoint's `upArrowCallout` cache is a banner with a centred arrow below it. */
function upArrowCalloutPoints(
	x: number,
	y: number,
	width: number,
	height: number,
	bannerRatio: number,
): string {
	const bannerBottom = y + height * bannerRatio;
	const stemLeft = x + width * 0.47;
	const stemRight = x + width * 0.53;
	const shoulderY = y + height * 0.78;
	return [
		`${x},${y}`,
		`${x + width},${y}`,
		`${x + width},${bannerBottom}`,
		`${stemRight},${bannerBottom}`,
		`${stemRight},${shoulderY}`,
		`${x + width / 2},${y + height}`,
		`${stemLeft},${shoulderY}`,
		`${stemLeft},${bannerBottom}`,
		`${x},${bannerBottom}`,
	].join(' ');
}

// ── Component ───────────────────────────────────────────────────────────────

/**
 * Renders pre-computed drawing shapes that come directly from PowerPoint's
 * layout engine output.
 *
 * Each shape is positioned within an SVG viewBox derived from the bounding
 * box of all shapes. Supports ellipses, chevrons/homePlates, and rounded
 * rectangles, with optional rotation, stroke, shadow, and text labels.
 */
export function DrawingShapeRenderer({
	elementId,
	shapes,
	style,
	palette,
	nodes,
	nodeLabels,
}: DrawingShapeRendererProps): React.ReactElement {
	let minX = Infinity;
	let minY = Infinity;
	let maxX = -Infinity;
	let maxY = -Infinity;
	for (const s of shapes) {
		if (s.x < minX) {
			minX = s.x;
		}
		if (s.y < minY) {
			minY = s.y;
		}
		if (s.x + s.width > maxX) {
			maxX = s.x + s.width;
		}
		if (s.y + s.height > maxY) {
			maxY = s.y + s.height;
		}
	}

	const drawingW = maxX - minX || 1;
	const drawingH = maxY - minY || 1;
	const shadow = styleShadow(style);
	const sw = styleStroke(style);

	return (
		<svg
			viewBox={`0 0 ${drawingW} ${drawingH}`}
			className='w-full h-full pointer-events-none'
			preserveAspectRatio='xMidYMid meet'
			data-testid='smartart-drawing-shapes'
		>
			{shapes.map((shape, i) => {
				const gradient = drawingShapeGradientDef(`${elementId}-dspgrad-${shape.id}-${i}`, shape);
				// Precedence: gradient -> pattern foreground -> solid/palette.
				const fill = shape.fillNone
					? 'none'
					: gradient?.ref ??
						shape.fillPatternForegroundColor ??
						shape.fillColor ??
						colour(i, palette);
				const relX = shape.x - minX;
				const relY = shape.y - minY;
				const rx = shape.shapeType === 'roundRect' ? Math.min(shape.width, shape.height) * 0.1 : 0;
				const isEllipse = shape.shapeType === 'ellipse';
				const isLine = shape.shapeType === 'line';
				const isChevron = shape.shapeType === 'chevron' || shape.shapeType === 'homePlate';
				const isUpArrowCallout = shape.shapeType === 'upArrowCallout';
				// The cached drawing may use shape adjustments that are not repeated in
				// the normalized shape type. Detail rectangles begin exactly at the
				// banner's bottom, so recover the authored banner ratio from that geometry.
				const overlappingDetailTop = isUpArrowCallout
					? shapes
							.filter(
								(candidate) =>
									candidate !== shape &&
									candidate.shapeType === 'rect' &&
									candidate.y > shape.y &&
									candidate.y < shape.y + shape.height &&
									candidate.x < shape.x + shape.width &&
									candidate.x + candidate.width > shape.x,
							)
							.map((candidate) => candidate.y)
							.sort((a, b) => a - b)[0]
					: undefined;
				const calloutBannerRatio =
					overlappingDetailTop !== undefined
						? Math.max(0.2, Math.min(0.8, (overlappingDetailTop - shape.y) / shape.height))
						: 0.66;
				// Cached `upArrowCallout` shapes are normalized below into a banner with
				// a downward arrow (the visual result after PowerPoint's 180deg transform).
				// Applying the source rotation again flips the arrow back above the banner
				// and lets the following detail cells cover its heading text.
				const rotation = shape.rotation && !isUpArrowCallout
					? `rotate(${shape.rotation} ${relX + shape.width / 2} ${relY + shape.height / 2})`
					: undefined;
				const strokeCol = shape.strokeColor ?? (sw > 0 ? 'rgba(255,255,255,0.3)' : 'none');
				const strokeW = shape.strokeWidth ?? sw;
				const lineColor = strokeCol === 'none' ? shape.fillColor ?? colour(i, palette) : strokeCol;
				const lineWidth = Math.max(strokeW, 0.5);
				const markerSeed = `${elementId}-${shape.id}-${i}`.replace(/[^a-zA-Z0-9_-]/gu, '_');
				const startMarkerId = `${markerSeed}-start`;
				const endMarkerId = `${markerSeed}-end`;
				const underlayFill = shape.fillNone
					? shapes
							.slice(0, i)
							.reverse()
							.find((candidate) => {
								const centerX = shape.x + shape.width / 2;
								const centerY = shape.y + shape.height / 2;
								return (
									!candidate.fillNone &&
									centerX >= candidate.x &&
									centerX <= candidate.x + candidate.width &&
									centerY >= candidate.y &&
									centerY <= candidate.y + candidate.height
								);
							})?.fillColor
					: undefined;
				const fontSize =
					shape.fontSize ?? fitFontSize(shape.text ?? '', shape.width * 0.85, shape.height, 14);
				const contrastFill =
					underlayFill ??
					shape.fillColor ??
					shape.fillPatternForegroundColor ??
					shape.fillGradientStops?.[Math.floor(shape.fillGradientStops.length / 2)]?.color ??
					fill;

				const nodeId = nodes ? resolveDrawingShapeNodeId(shape, i, shapes, nodes) : undefined;
				const nodeLabel = nodeId ? nodeLabels?.get(nodeId) : undefined;
				const groupProps = nodeId
					? smartArtNodeGroupProps(nodeId, shadow, nodeLabel)
					: { style: { filter: shadow } };

				return (
					<g key={`${elementId}-dsp-${shape.id}-${i}`} {...groupProps}>
						{nodeLabel ? <title>{nodeLabel}</title> : null}
						{gradient || isLine ? (
							<defs>
								{gradient?.def}
								{isLine
									? renderConnectorMarker(startMarkerId, shape.startArrow, lineColor, shape.startArrowWidth, shape.startArrowLength)
									: null}
								{isLine
									? renderConnectorMarker(endMarkerId, shape.endArrow, lineColor, shape.endArrowWidth, shape.endArrowLength)
									: null}
							</defs>
						) : null}
						{isLine ? (
							<line
								x1={relX}
								y1={relY}
								x2={relX + shape.width}
								y2={relY + shape.height}
								stroke={lineColor}
								strokeWidth={lineWidth}
								strokeDasharray={drawingLineDashArray(shape.strokeDash, lineWidth)}
								markerStart={shape.startArrow && shape.startArrow !== 'none' ? `url(#${startMarkerId})` : undefined}
								markerEnd={shape.endArrow && shape.endArrow !== 'none' ? `url(#${endMarkerId})` : undefined}
								vectorEffect='non-scaling-stroke'
							/>
						) : shape.fillImageUrl ? (
							<image
								x={relX}
								y={relY}
								width={shape.width}
								height={shape.height}
								href={shape.fillImageUrl}
								preserveAspectRatio='xMidYMid meet'
								transform={rotation}
							/>
						) : isEllipse ? (
							<ellipse
								cx={relX + shape.width / 2}
								cy={relY + shape.height / 2}
								rx={shape.width / 2}
								ry={shape.height / 2}
								fill={fill}
								stroke={strokeCol}
								strokeWidth={strokeW}
								transform={rotation}
							/>
						) : isChevron || isUpArrowCallout ? (
							<polygon
								points={
									isUpArrowCallout
										? upArrowCalloutPoints(
												relX,
												relY,
												shape.width,
												shape.height,
												calloutBannerRatio,
											)
										: chevronPoints(relX, relY, shape.width, shape.height)
								}
								fill={fill}
								stroke={strokeCol}
								strokeWidth={strokeW}
								transform={rotation}
							/>
						) : (
							<rect
								x={relX}
								y={relY}
								width={shape.width}
								height={shape.height}
								rx={rx}
								fill={fill}
								stroke={strokeCol}
								strokeWidth={strokeW}
								transform={rotation}
							/>
						)}
						{shape.text ? (
							<SmartArtNodeText
								x={relX + shape.width / 2}
								y={
									isUpArrowCallout
										? relY + shape.height * (calloutBannerRatio / 2)
										: relY + shape.height / 2
								}
								text={shape.text}
								fill={shape.fontColor ?? contrastingTextColor(contrastFill)}
								fontSize={fontSize}
								maxWidth={shape.width * 0.82}
								className='pointer-events-none'
							/>
						) : null}
					</g>
				);
			})}
		</svg>
	);
}
// Modified by Aixa Ltd from the original ChristopherVR/pptx-viewer project.

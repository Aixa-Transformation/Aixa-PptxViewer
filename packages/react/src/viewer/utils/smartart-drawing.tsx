import type { PptxElement, PptxSmartArtDrawingShape, SmartArtStyle } from 'pptx-viewer-core';
import React from 'react';

import { colour, contrastingTextColor, styleShadow, styleStroke } from './smartart-helpers';
import { renderConnectorMarker } from './connector-path';

function smartArtDashArray(token: string | undefined, strokeWidth: number): string | undefined {
	if (!token || token === 'solid') return undefined;
	if (token === 'dash' || token === 'sysDash') return `${4 * strokeWidth} ${3 * strokeWidth}`;
	if (token === 'dot' || token === 'sysDot') return `${strokeWidth} ${2 * strokeWidth}`;
	if (token.includes('DashDot')) return `${4 * strokeWidth} ${2 * strokeWidth} ${strokeWidth} ${2 * strokeWidth}`;
	return `${4 * strokeWidth} ${3 * strokeWidth}`;
}

/**
 * Render pre-computed drawing shapes from `ppt/diagrams/drawing*.xml`.
 * These are the shapes as computed by PowerPoint's layout engine.
 */
export function renderDrawingShapes(
	element: PptxElement,
	shapes: PptxSmartArtDrawingShape[],
	style: SmartArtStyle,
	palette: string[],
): React.ReactNode {
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
		>
			{shapes.map((shape, i) => {
				const fill = shape.fillNone ? 'none' : shape.fillColor ?? colour(i, palette);
				const relX = shape.x - minX;
				const relY = shape.y - minY;
				const rx = shape.shapeType === 'roundRect' ? Math.min(shape.width, shape.height) * 0.1 : 0;
				const isEllipse = shape.shapeType === 'ellipse';
				const isLine = shape.shapeType === 'line';
				const lineColor = shape.strokeColor ?? shape.fillColor ?? colour(i, palette);
				const lineWidth = shape.strokeWidth ?? Math.max(sw, 1);
				const markerSeed = `${element.id}-${shape.id}-${i}`.replace(/[^a-zA-Z0-9_-]/gu, '_');
				const startMarkerId = `${markerSeed}-start`;
				const endMarkerId = `${markerSeed}-end`;

				return (
					<g key={`${element.id}-dsp-${shape.id}-${i}`} style={{ filter: shadow }}>
						{isLine ? (
							<>
								<defs>
									{renderConnectorMarker(startMarkerId, shape.startArrow, lineColor, shape.startArrowWidth, shape.startArrowLength)}
									{renderConnectorMarker(endMarkerId, shape.endArrow, lineColor, shape.endArrowWidth, shape.endArrowLength)}
								</defs>
								<line
									x1={relX}
									y1={relY}
									x2={relX + shape.width}
									y2={relY + shape.height}
									stroke={lineColor}
									strokeWidth={lineWidth}
									strokeDasharray={smartArtDashArray(shape.strokeDash, lineWidth)}
									markerStart={shape.startArrow && shape.startArrow !== 'none' ? `url(#${startMarkerId})` : undefined}
									markerEnd={shape.endArrow && shape.endArrow !== 'none' ? `url(#${endMarkerId})` : undefined}
									vectorEffect='non-scaling-stroke'
								/>
							</>
						) : shape.fillImageUrl ? (
							<image
								x={relX}
								y={relY}
								width={shape.width}
								height={shape.height}
								href={shape.fillImageUrl}
								preserveAspectRatio='xMidYMid meet'
							/>
						) : isEllipse ? (
							<ellipse
								cx={relX + shape.width / 2}
								cy={relY + shape.height / 2}
								rx={shape.width / 2}
								ry={shape.height / 2}
								fill={fill}
								stroke={shape.strokeColor ?? (sw > 0 ? 'rgba(255,255,255,0.3)' : 'none')}
								strokeWidth={shape.strokeWidth ?? sw}
								transform={
									shape.rotation
										? `rotate(${shape.rotation} ${relX + shape.width / 2} ${relY + shape.height / 2})`
										: undefined
								}
							/>
						) : (
							<rect
								x={relX}
								y={relY}
								width={shape.width}
								height={shape.height}
								rx={rx}
								fill={fill}
								stroke={shape.strokeColor ?? (sw > 0 ? 'rgba(255,255,255,0.3)' : 'none')}
								strokeWidth={shape.strokeWidth ?? sw}
								transform={
									shape.rotation
										? `rotate(${shape.rotation} ${relX + shape.width / 2} ${relY + shape.height / 2})`
										: undefined
								}
							/>
						)}
						{shape.text ? (
							<text
								x={relX + shape.width / 2}
								y={relY + shape.height / 2}
								textAnchor='middle'
								dominantBaseline='central'
								fill={shape.fontColor ?? contrastingTextColor(fill)}
								fontSize={shape.fontSize ?? Math.max(8, Math.min(14, shape.height * 0.2))}
								className='pointer-events-none'
							>
								{shape.text}
							</text>
						) : null}
					</g>
				);
			})}
		</svg>
	);
}
// Modified by Aixa Ltd from the original ChristopherVR/pptx-viewer project.

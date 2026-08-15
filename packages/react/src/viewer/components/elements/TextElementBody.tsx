import { getLinkedTextBoxSegments, hasTextProperties } from 'pptx-viewer-core';
import React from 'react';

import { DEFAULT_TEXT_COLOR } from '../../constants';
import {
	cn,
	getTextCompensationTransform,
	getTextLayoutStyle,
	getTextWarpStyle,
	renderTextSegments,
} from '../../utils';
import { buildTextBody3DSceneStyle } from '../../utils/text-effects';
import { shouldUseSvgWarp, WarpedText } from '../../utils/text-warp';
import { ActionButtonGlyphOverlay, isActionButtonShape } from './ActionButtonGlyphOverlay';
import type { RenderBodyOptions } from './element-body-types';

export function shouldRenderTextBody(
	isTextElement: boolean,
	hasActualText: boolean,
	promptText: string | undefined,
	isPresentationPassive: boolean | undefined,
): boolean {
	if (!isTextElement) return Boolean(promptText) && !isPresentationPassive;
	return hasActualText || !promptText || !isPresentationPassive;
}

export function renderTextElementBody(options: RenderBodyOptions): React.ReactNode {
	const {
		el,
		vecShape,
		isTxtEl,
		txtS,
		txtSE,
		findHl,
		onHyperlinkClick,
		fieldContext,
		presentationElementStates,
		isPresentationPassive,
		slideElements,
	} = options;
	const isLinkedTextBox = hasTextProperties(el) && el.linkedTxbxId !== undefined;
	const linkedSegments =
		isLinkedTextBox && slideElements ? getLinkedTextBoxSegments(el, slideElements) : undefined;
	const useSvgWarp = shouldUseSvgWarp(
		hasTextProperties(el) ? el.textStyle?.textWarpPreset : undefined,
	);
	const scene3dStyle = hasTextProperties(el) ? buildTextBody3DSceneStyle(el.textStyle) : undefined;
	// `a:bodyPr/@rot` rotates the whole text body (degrees, clockwise positive).
	const bodyRotation = hasTextProperties(el) ? el.textStyle?.textBodyRotation : undefined;
	const rotationTransform =
		typeof bodyRotation === 'number' && Number.isFinite(bodyRotation) && bodyRotation !== 0
			? `rotate(${bodyRotation}deg)`
			: undefined;
	const composedTransform =
		[rotationTransform, getTextCompensationTransform(el), scene3dStyle?.transform]
			.filter(Boolean)
			.join(' ') || undefined;
	const transformStyle: React.CSSProperties = {
		transform: composedTransform,
		transformOrigin: 'center',
		...(scene3dStyle?.perspective ? { perspective: scene3dStyle.perspective } : {}),
		...(scene3dStyle?.transformStyle ? { transformStyle: scene3dStyle.transformStyle } : {}),
		...(isLinkedTextBox ? { overflow: 'hidden' } : {}),
	};
	// `a:noAutofit` with top anchoring keeps the authored font size and lets the
	// text body extend below its placeholder. A fixed `height: 100%` turns the
	// body into a constrained flex container and can suppress its final paragraph
	// even when vertical overflow is allowed. Keep the placeholder as the minimum
	// height while allowing the content box to grow naturally, matching Office.
	const allowsNaturalVerticalOverflow =
		hasTextProperties(el) &&
		el.textStyle?.autoFitMode === 'none' &&
		el.textStyle?.vertOverflow !== 'clip' &&
		(el.textStyle?.vAlign === undefined || el.textStyle.vAlign === 'top') &&
		!isLinkedTextBox;
	const textBodySizeStyle: React.CSSProperties | undefined = allowsNaturalVerticalOverflow
		? { height: 'auto', minHeight: '100%' }
		: undefined;
	const shapeType = 'shapeType' in el ? (el as { shapeType?: string }).shapeType : undefined;
	// Placeholder prompts ("Click to add title", master sample text, etc.) are
	// editor affordances, not slide content.  PowerPoint omits them from slide
	// show/export, so passive/read-only rendering must do the same.
	const textProperties = hasTextProperties(el) ? el : undefined;
	const hasActualText = Boolean(textProperties?.text) || Boolean(textProperties?.textSegments?.length);
	const shouldRenderText = shouldRenderTextBody(
		isTxtEl,
		hasActualText,
		textProperties?.promptText,
		isPresentationPassive,
	);

	// PowerPoint slide show/export omits an empty placeholder completely,
	// including any placeholder fill or outline inherited from the layout.
	if (!shouldRenderText) return null;

	return (
		<>
			{vecShape}
			{isActionButtonShape(shapeType) && <ActionButtonGlyphOverlay element={el} />}
			{useSvgWarp ? (
					<div
						className={cn(
							'relative z-10 w-full h-full',
							onHyperlinkClick ? '' : 'pointer-events-none',
						)}
						style={{ ...getTextLayoutStyle(el), ...transformStyle }}
					>
						<WarpedText
							element={el}
							width={el.width}
							height={el.height}
							fallbackColor={DEFAULT_TEXT_COLOR}
							findHighlights={findHl}
							fieldContext={fieldContext}
						/>
					</div>
				) : (
					<div
						className={cn(
							'relative z-10 w-full h-full whitespace-pre-wrap break-words leading-[1.3]',
							onHyperlinkClick ? '' : 'pointer-events-none',
						)}
						style={{
							...getTextLayoutStyle(el),
							...txtS,
							...getTextWarpStyle(txtSE),
							...transformStyle,
							...textBodySizeStyle,
						}}
					>
						{renderTextSegments(
							el,
							DEFAULT_TEXT_COLOR,
							undefined,
							findHl,
							onHyperlinkClick,
							fieldContext,
							presentationElementStates,
							linkedSegments ?? undefined,
							!isPresentationPassive,
						)}
					</div>
				)}
		</>
	);
}

import {
	evaluatePresetShape,
	getLinkedTextBoxSegments,
	hasTextProperties,
	type PptxElement,
} from 'pptx-viewer-core';
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

/**
 * PowerPoint keeps chevron text inside the rectangular portion between the
 * rear notch and the arrow tip.  Rendering text over the full shape bounds
 * clips the first words against the notch and changes the authored wrapping.
 *
 * Keep this scoped to chevrons for now: other preset shapes historically use
 * the full element box in the editor and applying every preset's text rect at
 * once would be a much broader layout change.
 */
export function getChevronTextFrameStyle(el: PptxElement): React.CSSProperties | undefined {
	if (!('shapeType' in el) || el.shapeType !== 'chevron') return undefined;

	const textRect = evaluatePresetShape(
		el.shapeType,
		el.width,
		el.height,
		el.shapeAdjustments,
	)?.textRect;
	if (!textRect) return undefined;

	const width = Math.max(0, textRect.r - textRect.l);
	const height = Math.max(0, textRect.b - textRect.t);
	if (width <= 0 || height <= 0) return undefined;

	return {
		position: 'absolute',
		left: textRect.l,
		top: textRect.t,
		width,
		height,
	};
}

/**
 * Normal PowerPoint autofit owns overflow inside a fixed text shape. Keep the
 * final renderer clipped to the same viewport as InlineTextEditor so a font
 * floor or a one-frame layout difference cannot paint text above/below the
 * selection box. `a:noAutofit` deliberately retains its authored overflow,
 * and nowrap text keeps its horizontal overflow semantics.
 */
export function shouldClipFixedTextBody(el: PptxElement, isLinkedTextBox = false): boolean {
	return Boolean(
		isLinkedTextBox ||
			(hasTextProperties(el) &&
				el.textStyle?.autoFitMode === 'normal' &&
				el.textStyle?.textWrap !== 'none'),
	);
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
		...(shouldClipFixedTextBody(el, isLinkedTextBox) ? { overflow: 'hidden' } : {}),
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
	const presetTextFrameStyle = getChevronTextFrameStyle(el);
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

	// An empty placeholder has nothing but faint prompt text, so mark out its
	// clickable area. The double outline reads against both dark and light
	// backgrounds without depending on the theme colours.
	const emptyPlaceholderStyle: React.CSSProperties | undefined =
		!hasActualText && textProperties?.promptText && !isPresentationPassive
			? {
					outline: '1px dashed rgba(127, 127, 127, 0.9)',
					outlineOffset: '-1px',
					boxShadow: 'inset 0 0 0 2px rgba(255, 255, 255, 0.25)',
				}
			: undefined;

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
						style={{
							...getTextLayoutStyle(el),
							...transformStyle,
							...presetTextFrameStyle,
						}}
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
							...presetTextFrameStyle,
							...emptyPlaceholderStyle,
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

import { hasTextProperties } from 'pptx-viewer-core';
import type { PptxElement, TextStyle, BulletInfo } from 'pptx-viewer-core';
import {
	resolveCssTextAlign,
	resolveParagraphAlign,
	resolveParagraphRtl,
} from 'pptx-viewer-shared';
import React from 'react';

import type { ElementAnimationState } from './animation-timeline';
import { getKinsokuLineBreakStyles } from './kinsoku-styles';
import { wrapWithTextBuildAnimation } from './text-animation';
import type { ParagraphEntry } from './text-animation';
import type { FieldSubstitutionContext } from './text-field-substitution';
import { resolveParagraphSpacing } from './text-paragraph-spacing';
import type { ElementFindHighlights } from './text-segment-helpers';
import { renderSingleSegment } from './text-segment-render';

// Per-paragraph BiDi direction + text-alignment resolution now lives in
// pptx-viewer-shared (render/text-paragraph-style). Re-exported here so existing
// React import paths keep working.
export { resolveCssTextAlign, resolveParagraphAlign, resolveParagraphRtl };

function groupSegmentsIntoParagraphs(
	segments: ReadonlyArray<{
		text: string;
		style: TextStyle;
		bulletInfo?: BulletInfo;
		fieldType?: string;
		equationXml?: Record<string, unknown>;
		equationNumber?: string;
		rubyText?: string;
		rubyAlignment?: string;
		rubyFontSize?: number;
		rubyStyle?: TextStyle;
	}>,
): Array<Array<ParagraphEntry>> {
	const paragraphs: Array<Array<ParagraphEntry>> = [];
	let current: Array<ParagraphEntry> = [];

	for (let i = 0; i < segments.length; i++) {
		const seg = segments[i];
		if (seg.text === '\n') {
			const currentHasVisibleContent = current.some(({ segment }) =>
				Boolean(segment.text && segment.text.trim().length > 0),
			);
			// Preserve an authored empty DrawingML paragraph. PowerPoint gives it
			// a full line box (often deliberately used as spacing between a
			// centered subheading and body bullets); dropping all of its segments
			// collapses the following content upward. Leading empty paragraphs are
			// different: PowerPoint collapses those instead of shifting the whole
			// text body down.
			if (!currentHasVisibleContent && paragraphs.length === 0) {
				// A leading a:r with an empty a:t followed by a:br is another
				// common PowerPoint representation of a collapsed leading line.
				// Discard both instead of letting the empty run create a line box.
				current = [];
				continue;
			}
			if (current.length === 0 && paragraphs.length > 0) {
				current.push({
					segment: { ...seg, text: '' },
					globalIndex: i,
				});
			}
			if (current.length > 0) {
				paragraphs.push(current);
			}
			current = [];
		} else {
			current.push({ segment: seg, globalIndex: i });
		}
	}
	if (current.length > 0 || paragraphs.length === 0) {
		paragraphs.push(current);
	}

	// PowerPoint frequently leaves a final empty a:p in placeholders. It is a
	// paragraph terminator, not an authored spacer, and does not consume a line
	// in slide-show rendering. An inherited bullet may have already produced a
	// marker-only segment for that paragraph, so test visible non-marker content
	// rather than merely checking whether the segment array is empty. Keep empty
	// paragraphs between visible paragraphs: those are intentional spacing.
	while (
		paragraphs.length > 1 &&
		!paragraphs[paragraphs.length - 1].some(
			({ segment }) =>
				!segment.bulletInfo && Boolean(segment.text && segment.text.trim().length > 0),
		)
	) {
		paragraphs.pop();
	}

	return paragraphs;
}

export function renderTextSegments(
	element: PptxElement,
	fallbackColor: string,
	emptyFallback?: string,
	findHighlights?: ElementFindHighlights,
	onHyperlinkClick?: (url: string) => void,
	fieldContext?: FieldSubstitutionContext,
	/** Per-sub-element animation states for text build animations. */
	subElementAnimStates?: ReadonlyMap<string, ElementAnimationState>,
	/** When provided, these segments replace element.textSegments for rendering (used by linked text box overflow). */
	segmentOverrides?: ReadonlyArray<{
		text: string;
		style: TextStyle;
		bulletInfo?: BulletInfo;
		fieldType?: string;
		equationXml?: Record<string, unknown>;
		equationNumber?: string;
		isParagraphBreak?: boolean;
		rubyText?: string;
		rubyAlignment?: string;
		rubyFontSize?: number;
		rubyStyle?: TextStyle;
		/** Per-paragraph geometry authored on the first segment of a paragraph. */
		paragraphProperties?: TextStyle;
	}>,
	/** When true, hyperlinks require Ctrl+Click (editing mode). */
	requireCtrlClick?: boolean,
): React.ReactNode {
	if (!hasTextProperties(element)) {
		return emptyFallback || null;
	}

	const effectiveSegments = segmentOverrides ?? element.textSegments;

	if (!effectiveSegments || effectiveSegments.length === 0) {
		if (!element.text && element.promptText) {
			return (
				<span
					style={{
						opacity: 0.5,
						color: '#888888',
						pointerEvents: 'none',
					}}
				>
					{element.promptText}
				</span>
			);
		}
		return element.text || emptyFallback || '';
	}

	const paragraphs = groupSegmentsIntoParagraphs(effectiveSegments);
	const paragraphIndents = hasTextProperties(element) ? element.paragraphIndents : undefined;
	const elementRtl = hasTextProperties(element) ? element.textStyle?.rtl : undefined;

	const elementAlign = hasTextProperties(element) ? element.textStyle?.align : undefined;
	const bodyStyle = hasTextProperties(element) ? element.textStyle : undefined;
	// `spcFirstLastPara` is false by default in DrawingML. Edge paragraph
	// spacing applies only when the body explicitly opts in.
	const spaceFirstLast = bodyStyle?.spaceFirstLastParagraph === true;

	return paragraphs.map((paraSegments, paraIndex) => {
		const paraIndent = paragraphIndents?.[paraIndex];
		const rawMarginLeft =
			typeof paraIndent?.marginLeft === 'number' && paraIndent.marginLeft !== 0
				? paraIndent.marginLeft
				: undefined;
		const rawTextIndent =
			typeof paraIndent?.indent === 'number' && paraIndent.indent !== 0
				? paraIndent.indent
				: undefined;

		const firstSeg = paraSegments[0];
		const bulletInfo = firstSeg?.segment.bulletInfo;
		const authoredText = paraSegments
			.filter(({ segment }) => !segment.bulletInfo)
			.map(({ segment }) => segment.text)
			.join('')
			.trimStart();
		const generatedMarker = firstSeg?.segment.text?.trim();
		// Some producers preserve the visible auto-number in a text run while
		// also authoring a:buAutoNum. PowerPoint displays a single marker; avoid
		// rendering the generated marker twice in that case (for example
		// `3. 3. Item`). Keep ordinary numbered lists untouched.
		const hasDuplicateAuthoredAutoNumber = Boolean(
			bulletInfo?.autoNumType &&
			generatedMarker &&
			(authoredText === generatedMarker ||
				authoredText.startsWith(`${generatedMarker} `) ||
				authoredText.startsWith(`${generatedMarker}\t`)),
		);
		// Suppress bullets for paragraphs with no visible text content.
		// In PowerPoint, empty bullet paragraphs (e.g. residual first paragraphs
		// or line breaks with no text) don't render a bullet character.
		const hasVisibleTextContent = paraSegments.some(({ segment }) => {
			// Skip the bullet segment itself: it only contains the marker text
			if (segment.bulletInfo) {
				return false;
			}
			return Boolean(segment.text) && segment.text.trim().length > 0;
		});
		const hasBullet =
			bulletInfo &&
			!bulletInfo.none &&
			hasVisibleTextContent &&
			!hasDuplicateAuthoredAutoNumber;
		const paraRtl = resolveParagraphRtl(paraSegments, elementRtl);
		const isRtlParagraph = paraRtl === true;

		// Resolve explicit paragraph alignment from segment styles
		const paraAlign = resolveParagraphAlign(paraSegments, elementAlign);
		const cssTextAlign = resolveCssTextAlign(paraAlign, isRtlParagraph);

		// For RTL paragraphs, swap marginLeft/textIndent to marginRight
		// so bullets and indentation appear on the correct (right) side.
		const paraMarginLeft = isRtlParagraph ? undefined : rawMarginLeft;
		const paraMarginRight = isRtlParagraph ? rawMarginLeft : undefined;
		const paraTextIndent = rawTextIndent;

		// Per-paragraph kinsoku line-breaking styles from the first segment's style.
		// Paragraph-level properties (eaLineBreak, hangingPunctuation, latinLineBreak)
		// are stored on the TextStyle of paragraph segments.
		const paraKinsokuStyle = getKinsokuLineBreakStyles(firstSeg?.segment.style);
		const hasParaKinsoku = Object.keys(paraKinsokuStyle).length > 0;

		// Per-paragraph line spacing (a:lnSpc) and space before/after
		// (a:spcBef / a:spcAft), sourced from this paragraph's own geometry with
		// a body-level fallback for inherited/single-level text.
		const paraProps = effectiveSegments[firstSeg?.globalIndex ?? -1]?.paragraphProperties;
		// The first concrete run carries the fully resolved list-level paragraph
		// spacing even when `paragraphProperties` only contains locally authored
		// indent fields. Merge both so level-1/level-2 spacing is not replaced by
		// the body-level spacing from level 0.
		const effectiveParaProps = {
			...firstSeg?.segment.style,
			...paraProps,
		};
		const paragraphAutoFitScale =
			typeof element.textStyle?.autoFitFontScale === 'number' &&
			element.textStyle.autoFitFontScale > 0 &&
			element.textStyle.autoFitFontScale < 1
				? element.textStyle.autoFitFontScale
				: 1;
		const spacing = resolveParagraphSpacing({
			paraProps: effectiveParaProps,
			bodyStyle,
			isFirst: paraIndex === 0,
			isLast: paraIndex === paragraphs.length - 1,
			spaceFirstLast,
		});
		// PowerPoint's normAutofit scales point-based paragraph spacing together
		// with the text. Keeping authored spcBef/spcAft at their original size
		// accumulates excessive gaps in long bullet lists even though the glyphs
		// themselves have been reduced.
		const hasAuthoredParagraphLineSpacing =
			paraProps?.lineSpacing !== undefined || paraProps?.lineSpacingExactPt !== undefined;
		if (!hasAuthoredParagraphLineSpacing) {
			if (typeof spacing.marginTop === 'number') {
				spacing.marginTop *= paragraphAutoFitScale;
			}
			if (typeof spacing.marginBottom === 'number') {
				spacing.marginBottom *= paragraphAutoFitScale;
			}
		}
		// `normAutofit/@lnSpcReduction` reduces the resolved paragraph line
		// multiplier in addition to scaling the run font. Paragraph wrappers own
		// their line-height, so the body-level auto-fit style cannot apply this for
		// them. Merge it here to avoid progressively over-spacing dense bullet
		// lists and clipping their final paragraphs.
		const autoFitLineSpacingReduction = element.textStyle?.autoFitLineSpacingReduction;
		if (
			typeof spacing.lineHeight === 'number' &&
			typeof autoFitLineSpacingReduction === 'number' &&
			autoFitLineSpacingReduction > 0
		) {
			spacing.lineHeight = Math.max(1, spacing.lineHeight - autoFitLineSpacingReduction);
		}
		const hasParaSpacing =
			spacing.marginTop !== undefined ||
			spacing.marginBottom !== undefined ||
			spacing.lineHeight !== undefined;

		const paraStyle: React.CSSProperties & { '--pptx-paragraph-align'?: string } = {
			...paraKinsokuStyle,
		};
		// A CSS line box contains an invisible "strut" derived from the wrapper's
		// inherited font size. When a PowerPoint paragraph overrides the body font
		// (for example 22pt runs inside a 24pt body placeholder), leaving the
		// wrapper at the body size makes every line several pixels too tall and
		// eventually clips the last paragraph. Size the paragraph strut from its
		// resolved first run, matching DrawingML's per-paragraph line metrics.
		const paragraphFontSize =
			typeof effectiveParaProps.fontSize === 'number'
				? effectiveParaProps.fontSize * paragraphAutoFitScale
				: undefined;
		if (
			typeof paragraphFontSize === 'number' &&
			Number.isFinite(paragraphFontSize) &&
			paragraphFontSize > 0
		) {
			paraStyle.fontSize = paragraphFontSize;
		}
		if (spacing.marginTop !== undefined) {
			paraStyle.marginTop = spacing.marginTop;
		}
		if (spacing.marginBottom !== undefined) {
			paraStyle.marginBottom = spacing.marginBottom;
		}
		if (spacing.lineHeight !== undefined) {
			paraStyle.lineHeight = spacing.lineHeight;
		}
		if (paraMarginLeft !== undefined) {
			paraStyle.marginLeft = paraMarginLeft;
		}
		if (paraMarginRight !== undefined) {
			paraStyle.marginRight = paraMarginRight;
		}
		if (paraTextIndent !== undefined) {
			paraStyle.textIndent = paraTextIndent;
		}
		if (paraRtl !== undefined) {
			paraStyle.direction = paraRtl ? 'rtl' : 'ltr';
			// Use 'embed' so the paragraph establishes a BiDi embedding level.
			// This ensures numbers within RTL text render LTR naturally per the
			// Unicode Bidi Algorithm, while 'plaintext' is used as a fallback
			// only at the element/body level.
			paraStyle.unicodeBidi = 'embed';
		}
		if (cssTextAlign !== undefined) {
			paraStyle.textAlign = cssTextAlign;
			// Also expose the authored alignment as a custom property. The package
			// stylesheet consumes it with `!important` so host applications cannot
			// accidentally replace PowerPoint paragraph alignment with a global
			// `text-align` rule (common in presentation/player shells).
			paraStyle['--pptx-paragraph-align'] = cssTextAlign;
		}

		const needsWrapper =
			paraMarginLeft !== undefined ||
			paraMarginRight !== undefined ||
			paraTextIndent !== undefined ||
			hasBullet ||
			paraRtl !== undefined ||
			cssTextAlign !== undefined ||
			hasParaKinsoku ||
			hasParaSpacing;

		const renderedSegments = paraSegments
			.filter(({ segment }) => {
				// Skip bullet segments when the bullet should be suppressed
				if (!hasBullet && segment.bulletInfo) {
					return false;
				}
				return true;
			})
			.map(({ segment, globalIndex }) =>
				renderSingleSegment(
					element,
					segment,
					globalIndex,
					fallbackColor,
					findHighlights,
					hasBullet && globalIndex === firstSeg.globalIndex ? bulletInfo : undefined,
					onHyperlinkClick,
					fieldContext,
					paraRtl,
					requireCtrlClick,
					hasBullet && globalIndex === firstSeg.globalIndex
						? Math.max(0, -(rawTextIndent ?? 0))
						: undefined,
				),
			);
		const isAuthoredEmptyParagraph =
			paraSegments.length > 0 &&
			paraSegments.every(({ segment }) => !segment.text || segment.text.trim().length === 0);

		const wrappedContent = wrapWithTextBuildAnimation(
			element.id,
			paraIndex,
			isAuthoredEmptyParagraph ? [<br key={`${element.id}-empty-${paraIndex}`} />] : renderedSegments,
			paraSegments,
			subElementAnimStates,
		);

		if (!needsWrapper) {
			return (
				<React.Fragment key={`${element.id}-para-${paraIndex}`}>
					{wrappedContent}
					{paraIndex < paragraphs.length - 1 ? <br /> : null}
				</React.Fragment>
			);
		}

		return (
			<div key={`${element.id}-para-${paraIndex}`} data-pptx-paragraph style={paraStyle}>
				{wrappedContent}
			</div>
		);
	});
}
// Modified by Aixa Ltd from the original ChristopherVR/pptx-viewer project.

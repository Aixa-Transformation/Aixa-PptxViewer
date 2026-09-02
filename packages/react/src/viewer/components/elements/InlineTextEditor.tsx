import type { PptxElement, TextStyle } from 'pptx-viewer-core';
import { hasTextProperties } from 'pptx-viewer-core';
import React, { useRef, useEffect, useLayoutEffect, useCallback } from 'react';

import { DEFAULT_TEXT_COLOR } from '../../constants';
import { getTextCompensationTransform, getTextWarpStyle, renderTextSegments } from '../../utils';
import {
	getPendingSelectionRestore,
	restoreSegmentSelection,
} from '../../utils/inline-selection-utils';
import {
	findLargestFittingInlineTextScale,
	inlineTextFitMetricsFit,
	normalizeInlineTextAutoFitScale,
} from './inline-text-autofit';
import { getListContinuationMarker } from './inline-list-enter';

const INLINE_TEXT_AUTOFIT_STEP = 2;
const MIN_INLINE_TEXT_FONT_SIZE = 8;
const VIEW_MODE_FIT_SAFETY_PX = 4;
const LIST_CONTINUATION_CARET_PLACEHOLDER = '\u200B';

/**
 * Rich inline text editor: uses a `contentEditable` div that renders the same
 * rich text segments as view mode so formatting (per-run fonts, sizes, colors,
 * bullets, paragraph indentation, text effects) is preserved while editing.
 *
 * The editor extracts plain text on commit via `innerText` and passes it to the
 * parent's `onEditChange` callback, which feeds into `remapTextToSegments` to
 * redistribute the edited text across the original rich segments.
 *
 * The outer wrapper matches the view-mode text container exactly:
 * - `getTextLayoutStyle` for flex vertical alignment, body-inset padding, columns
 * - `getTextStyleForElement` (textStyle) for element-level font defaults
 * - `getTextWarpStyle` for text warp 3D transforms
 * - `getTextCompensationTransform` for rotation compensation
 */
export function InlineTextEditor({
	initialText,
	spellCheck,
	rtl,
	textDirection: _textDirection,
	textStyle,
	textStyleRaw,
	layoutStyle,
	element,
	onCommit,
	onCancel,
	onEditChange,
	onFormatText,
	slideHeight: _slideHeight,
}: {
	initialText: string;
	spellCheck: boolean;
	rtl?: boolean;
	textDirection?: TextStyle['textDirection'];
	textStyle: React.CSSProperties;
	/** Raw TextStyle object for computing warp transforms. */
	textStyleRaw?: TextStyle;
	/** Layout style from getTextLayoutStyle; provides flex vertical alignment. */
	layoutStyle: React.CSSProperties;
	element: PptxElement;
	onCommit: (
		autoFitHeight?: number,
		committedTextOverride?: string,
		autoFitFontScale?: number,
	) => void;
	onCancel: () => void;
	onEditChange: (t: string, autoFitFontScale?: number) => void;
	/** Called when the user applies formatting via keyboard shortcut (Ctrl+B/I/U). */
	onFormatText?: (updates: Partial<TextStyle>) => void;
	/** Retained for API compatibility; font autofit keeps the shape geometry fixed. */
	slideHeight?: number;
}) {
	const editorRef = useRef<HTMLDivElement>(null);
	const pendingListParagraphsRef = useRef<Set<HTMLElement>>(new Set());
	const importedAutoFitScale =
		typeof textStyleRaw?.autoFitFontScale === 'number' &&
		textStyleRaw.autoFitFontScale > 0 &&
		textStyleRaw.autoFitFontScale < 1
			? textStyleRaw.autoFitFontScale
			: 1;
	const autoFitFontScaleRef = useRef(importedAutoFitScale);

	// The editor is UNCONTROLLED: its content is seeded exactly once (below) and
	// the DOM owns the text from then on. `initialText` is updated by the parent
	// on every keystroke (via onEditChange), so if we rendered it as children
	// React would rewrite the text node on each change and the caret would jump
	// back to the start / typing would reverse. We therefore capture the seed
	// content on first render and never re-render it; live edits flow out through
	// handleInput, and the latest value is read from the DOM on commit/blur.
	const seedRef = useRef<{ initialText: string; hasRichSegments: boolean } | null>(null);
	if (seedRef.current === null) {
		seedRef.current = {
			initialText,
			hasRichSegments: Boolean(
				hasTextProperties(element) && element.textSegments && element.textSegments.length > 0,
			),
		};
	}
	const seed = seedRef.current;

	// Extract plain text from the contentEditable div
	const extractText = useCallback((): string => {
		const el = editorRef.current;
		if (!el) {
			return seed.initialText;
		}
		// Serialize the editor's top-level flow instead of relying on innerText.
		// A text box can mix wrapped DrawingML paragraphs with ordinary inline runs;
		// reading only [data-pptx-paragraph] nodes drops every unwrapped paragraph
		// below them. Conversely, innerText inserts a phantom newline between a
		// flex list marker and its editable content column. Walking direct children
		// preserves both representations and inserts separators only at real
		// paragraph boundaries.
		{
			const readNodeText = (node: Node): string => {
				if (node.nodeType === Node.TEXT_NODE) {
					return (node.textContent ?? '').replaceAll(
						LIST_CONTINUATION_CARET_PLACEHOLDER,
						'',
					);
				}
				if (node instanceof HTMLBRElement) {
					return '\n';
				}
				return Array.from(node.childNodes).map(readNodeText).join('');
			};
			const paragraphs: string[] = [];
			let inlineParagraph = '';
			const pushInlineParagraph = (): void => {
				paragraphs.push(inlineParagraph);
				inlineParagraph = '';
			};
			for (const child of Array.from(el.childNodes)) {
				if (child instanceof HTMLBRElement) {
					pushInlineParagraph();
					continue;
				}
				if (
					child instanceof HTMLElement &&
					(child.hasAttribute('data-pptx-paragraph') ||
						child.tagName === 'DIV' ||
						child.tagName === 'P')
				) {
					if (inlineParagraph.length > 0) {
						pushInlineParagraph();
					}
					const blockText = readNodeText(child);
					// A lone BR is the browser's caret holder for an authored empty
					// paragraph; the wrapper boundary already supplies its newline.
					paragraphs.push(blockText === '\n' ? '' : blockText);
					continue;
				}
				inlineParagraph += readNodeText(child);
			}
			if (inlineParagraph.length > 0 || paragraphs.length === 0) {
				paragraphs.push(inlineParagraph);
			}
			return paragraphs.join('\n');
		}
	}, [seed]);

	const measureAndScaleText = useCallback((
		allowGrowth = false,
		limitToOneStep = false,
		bottomSafetyInset = 0,
	): number => {
		const editor = editorRef.current;
		if (!editor) {
			return autoFitFontScaleRef.current;
		}

		// Preserve the authored font sizes per node. Applying one scale to those
		// baselines reflows every rich-text run inside the fixed shape,
		// including mixed-size text and list markers. Existing normAutofit values
		// are divided out once so deleting text can grow back toward 100%.
		const currentScale = Math.max(0.05, autoFitFontScaleRef.current || 1);
		// A slide can briefly report a zero-size editing box while it is being
		// selected, scrolled into view, or moved between responsive panes. Treating
		// that transient state as overflow collapses the font to the minimum before
		// the real geometry is available.
		if (editor.clientHeight <= 1 || editor.clientWidth <= 1) {
			return currentScale;
		}
		const targets = [editor, ...editor.querySelectorAll<HTMLElement>('*')];
		let smallestBaseFontSize = Number.POSITIVE_INFINITY;
		const metricChangedOutsideAutoFit = (
			computedValue: number,
			baseValue: number,
		): boolean =>
			Number.isFinite(computedValue) &&
			Number.isFinite(baseValue) &&
			Math.abs(computedValue - baseValue * currentScale) > 0.05;
		for (const target of targets) {
			const computedStyle = window.getComputedStyle(target);
			const computedFontSize = Number.parseFloat(computedStyle.fontSize);
			let baseFontSize = Number(target.dataset.pptxAutoFitBaseFontSize);
			if (
				!Number.isFinite(baseFontSize) ||
				baseFontSize <= 0 ||
				metricChangedOutsideAutoFit(computedFontSize, baseFontSize)
			) {
				if (!Number.isFinite(computedFontSize) || computedFontSize <= 0) {
					continue;
				}
				// React can reuse this DOM node when the toolbar changes its authored
				// font size. The old dataset is only an autofit baseline; if the live
				// computed value no longer equals base * scale, rebase it instead of
				// snapping the toolbar change back on the next input/blur measurement.
				baseFontSize = computedFontSize / currentScale;
				target.dataset.pptxAutoFitBaseFontSize = String(baseFontSize);
			}
			smallestBaseFontSize = Math.min(smallestBaseFontSize, baseFontSize);

			// Exact DrawingML line spacing resolves to a fixed CSS length. Capture
			// that metric as well as the paragraph spacing so it shrinks in lockstep
			// with the glyphs instead of forcing the fit search toward its minimum.
			if (target === editor || target.hasAttribute('data-pptx-paragraph')) {
				const computedLineHeight = Number.parseFloat(computedStyle.lineHeight);
				let baseLineHeight = Number(target.dataset.pptxAutoFitBaseLineHeight);
				if (
					(!Number.isFinite(baseLineHeight) ||
						baseLineHeight <= 0 ||
						metricChangedOutsideAutoFit(computedLineHeight, baseLineHeight)) &&
					Number.isFinite(computedLineHeight) &&
					computedLineHeight > 0
				) {
					baseLineHeight = computedLineHeight / currentScale;
					target.dataset.pptxAutoFitBaseLineHeight = String(baseLineHeight);
				}
			}
			if (target.hasAttribute('data-pptx-paragraph')) {
				for (const [datasetKey, cssValue] of [
					['pptxAutoFitBaseMarginTop', computedStyle.marginTop],
					['pptxAutoFitBaseMarginBottom', computedStyle.marginBottom],
				] as const) {
					const computedMargin = Number.parseFloat(cssValue);
					let baseMargin = Number(target.dataset[datasetKey]);
					if (
						(!Number.isFinite(baseMargin) ||
							metricChangedOutsideAutoFit(computedMargin, baseMargin)) &&
						Number.isFinite(computedMargin)
					) {
						baseMargin = computedMargin / currentScale;
						target.dataset[datasetKey] = String(baseMargin);
					}
				}
			}
		}

		const computedMinimumScale = Number.isFinite(smallestBaseFontSize)
			? Math.max(0.05, Math.min(1, MIN_INLINE_TEXT_FONT_SIZE / smallestBaseFontSize))
			: 0.1;
		// Never enlarge already-authored/imported miniature text merely because
		// our preferred live-edit floor is higher than its current scale.
		const minimumScale = Math.min(currentScale, computedMinimumScale);
		const applyScale = (scale: number): void => {
			for (const target of targets) {
				const baseFontSize = Number(target.dataset.pptxAutoFitBaseFontSize);
				if (Number.isFinite(baseFontSize) && baseFontSize > 0) {
					target.style.fontSize = `${baseFontSize * scale}px`;
				}
				const baseLineHeight = Number(target.dataset.pptxAutoFitBaseLineHeight);
				if (Number.isFinite(baseLineHeight) && baseLineHeight > 0) {
					target.style.lineHeight = `${baseLineHeight * scale}px`;
				}
				if (target.hasAttribute('data-pptx-paragraph')) {
					const baseMarginTop = Number(target.dataset.pptxAutoFitBaseMarginTop);
					if (Number.isFinite(baseMarginTop)) {
						target.style.marginTop = `${baseMarginTop * scale}px`;
					}
					const baseMarginBottom = Number(target.dataset.pptxAutoFitBaseMarginBottom);
					if (Number.isFinite(baseMarginBottom)) {
						target.style.marginBottom = `${baseMarginBottom * scale}px`;
					}
				}
			}
		};
		const fits = (scale: number): boolean => {
			applyScale(scale);
			// Measure the complete laid-out text, including glyphs above a
			// middle-aligned flex box. The editor is normally clipped so typing
			// never paints outside the shape, but measuring that clipped rendering
			// can hide the very overflow normAutofit needs to detect. This mutation
			// is synchronous and restored to hidden before the browser paints.
			editor.style.overflow = 'visible';
			// The editor always wraps long words (`overflow-wrap: break-word`).
			// During an inserted list item Chromium can nevertheless report the
			// fixed marker column in scrollWidth in addition to the flexible text
			// column. Treating that transient hanging-indent width as real overflow
			// shrinks the font by another step on every keystroke even though every
			// rendered line fits. PowerPoint's Shrink text on overflow is driven by
			// whether the wrapped text body exceeds the fixed box height here.
			let visualBounds: {
				boxTop?: number;
				boxBottom?: number;
				contentTop?: number;
				contentBottom?: number;
			} = {};
			try {
				const box = editor.getBoundingClientRect();
				let contentTop = Number.POSITIVE_INFINITY;
				let contentBottom = Number.NEGATIVE_INFINITY;
				const walker = document.createTreeWalker(editor, 4);
				let textNode = walker.nextNode();
				while (textNode) {
					const visibleText = (textNode.textContent ?? '').replaceAll(
						LIST_CONTINUATION_CARET_PLACEHOLDER,
						'',
					);
					if (visibleText.length > 0) {
						const contentRange = document.createRange();
						contentRange.selectNodeContents(textNode);
						const content = contentRange.getBoundingClientRect();
						if (
							Number.isFinite(content.top) &&
							Number.isFinite(content.bottom) &&
							content.bottom > content.top
						) {
							contentTop = Math.min(contentTop, content.top);
							contentBottom = Math.max(contentBottom, content.bottom);
						}
					}
					textNode = walker.nextNode();
				}
				if (Number.isFinite(contentTop) && Number.isFinite(contentBottom)) {
					visualBounds = {
						boxTop: box.top,
						boxBottom: box.bottom,
						contentTop,
						contentBottom,
					};
				}
			} catch {
				// Older DOM shims do not expose Range#getBoundingClientRect. The
				// scroll-height check remains a safe fallback in those environments.
			}
			return inlineTextFitMetricsFit({
				scrollHeight: editor.scrollHeight,
				// A fixed-height contentEditable commonly reports scrollHeight equal
				// to clientHeight even for just one short line. Reducing clientHeight
				// by the view safety inset therefore makes every fitting box overflow
				// and shrinks it again on every edit/blur cycle. Keep the real scroll
				// viewport here; reserve the inset only against actual glyph bounds.
				clientHeight: editor.clientHeight,
				...visualBounds,
				...(typeof visualBounds.boxBottom === 'number'
					? { boxBottom: visualBounds.boxBottom - bottomSafetyInset }
					: {}),
			});
		};
		const fittedScale = normalizeInlineTextAutoFitScale(
			findLargestFittingInlineTextScale({
				fits,
				minScale: minimumScale,
				maxScale: allowGrowth ? 1 : currentScale,
			}),
		);
		const editorBaseFontSize = Number(editor.dataset.pptxAutoFitBaseFontSize);
		const scaleStep =
			Number.isFinite(editorBaseFontSize) && editorBaseFontSize > 0
				? INLINE_TEXT_AUTOFIT_STEP / editorBaseFontSize
				: 0.1;
		const nextScale = normalizeInlineTextAutoFitScale(
			!limitToOneStep
				// Entering edit mode is allowed to recover a stale miniature scale,
				// but must never make the text smaller before the user types.
				? allowGrowth
					? Math.max(currentScale, fittedScale)
					: fittedScale
				: allowGrowth
					// Deletion may restore at most two font-size units per edit.
					? Math.max(currentScale, Math.min(fittedScale, currentScale + scaleStep))
					// Insertion may shrink at most two units per edit. This prevents a
					// single wrapping fluctuation from turning 24pt text into miniature text.
					: Math.min(currentScale, Math.max(fittedScale, currentScale - scaleStep)),
		);
		applyScale(nextScale);
		editor.style.overflow = 'hidden';
		autoFitFontScaleRef.current = nextScale;
		return nextScale;
	}, []);

	const removeListContinuationCaretPlaceholder = useCallback(() => {
		const editor = editorRef.current;
		if (!editor) {
			return;
		}
		const selection = window.getSelection();
		const activeRange = selection?.rangeCount ? selection.getRangeAt(0) : undefined;
		for (const content of editor.querySelectorAll<HTMLElement>(
			'[data-pptx-list-continuation-content="true"]',
		)) {
			for (const node of Array.from(content.childNodes)) {
				if (node.nodeType !== Node.TEXT_NODE || !node.textContent?.includes(LIST_CONTINUATION_CARET_PLACEHOLDER)) {
					continue;
				}
				const previousText = node.textContent;
				const previousOffset =
					activeRange?.startContainer === node ? activeRange.startOffset : undefined;
				node.textContent = previousText.replaceAll(LIST_CONTINUATION_CARET_PLACEHOLDER, '');
				if (selection && activeRange && previousOffset !== undefined) {
					const placeholdersBeforeCaret = previousText
						.slice(0, previousOffset)
						.split(LIST_CONTINUATION_CARET_PLACEHOLDER).length - 1;
					activeRange.setStart(
						node,
						Math.max(0, Math.min(node.textContent?.length ?? 0, previousOffset - placeholdersBeforeCaret)),
					);
					activeRange.collapse(true);
					selection.removeAllRanges();
					selection.addRange(activeRange);
				}
			}
		}
	}, []);

	// Sync text to the parent and recompute PowerPoint-style font autofit on
	// every input. The shape geometry remains unchanged.
	const handleInput = useCallback((event: React.FormEvent<HTMLDivElement>) => {
		removeListContinuationCaretPlaceholder();
		const inputType = (event.nativeEvent as InputEvent).inputType ?? '';
		const liveAutoFitScale = measureAndScaleText(inputType.startsWith('delete'), true);
		onEditChange(extractText(), liveAutoFitScale);
	}, [extractText, measureAndScaleText, onEditChange, removeListContinuationCaretPlaceholder]);

	// When the caret sits at a soft word-wrap boundary (no explicit line break,
	// just CSS wrapping), the space that separates the two words is still part
	// of the text and lands right before the caret. Pressing Enter there splits
	// the DOM at that exact position, leaving the new paragraph break preceded
	// by a stray space: e.g. "fox jumps" wrapped as "fox " / "jumps" becomes
	// paragraphs "fox " and "jumps" instead of "fox" and "jumps". That extra,
	// invisible trailing character then counts toward the paragraph's measured
	// width, occasionally forcing an unwanted extra wrapped line. Since a space
	// immediately before a paragraph break is never visually meaningful, drop it
	// before the browser performs its native Enter/paragraph-split.
	const trimTrailingSpaceBeforeCaret = useCallback(() => {
		const selection = window.getSelection();
		if (!selection || !selection.isCollapsed || selection.rangeCount === 0) {
			return;
		}
		const range = selection.getRangeAt(0);
		const { startContainer, startOffset } = range;
		if (startContainer.nodeType !== Node.TEXT_NODE || startOffset === 0) {
			return;
		}
		const text = startContainer.textContent ?? '';
		if (text.charAt(startOffset - 1) !== ' ') {
			return;
		}
		const trimRange = document.createRange();
		trimRange.setStart(startContainer, startOffset - 1);
		trimRange.setEnd(startContainer, startOffset);
		trimRange.deleteContents();
	}, []);

	// Chromium creates a bare nested <div> when Enter is pressed in an ordinary
	// contentEditable paragraph. That temporary block does not carry the
	// DrawingML paragraph's line-height or spacing, so the live editor reports
	// that the content fits; after blur, remapTextToSegments restores those
	// paragraph metrics and the final line can be clipped below the fixed box.
	// Split the authored wrapper ourselves and clone its metrics so live autofit
	// measures the same paragraph topology that view mode will render.
	const insertPlainParagraphBreak = useCallback((): boolean => {
		const editor = editorRef.current;
		const selection = window.getSelection();
		if (!editor || !selection || selection.rangeCount === 0) {
			return false;
		}
		const range = selection.getRangeAt(0);
		const startElement =
			range.startContainer instanceof Element
				? range.startContainer
				: range.startContainer.parentElement;
		const paragraph = startElement?.closest<HTMLElement>('[data-pptx-paragraph]');
		if (
			!paragraph ||
			!editor.contains(paragraph) ||
			paragraph.hasAttribute('data-pptx-list-seg-idx') ||
			(!range.collapsed && !paragraph.contains(range.endContainer))
		) {
			return false;
		}

		if (!range.collapsed) {
			range.deleteContents();
		}
		const trailingRange = document.createRange();
		trailingRange.selectNodeContents(paragraph);
		trailingRange.setStart(range.startContainer, range.startOffset);
		const trailingContent = trailingRange.extractContents();
		const makeCaretPlaceholder = (): HTMLSpanElement => {
			const placeholder = document.createElement('span');
			placeholder.dataset.pptxPlainCaretPlaceholder = 'true';
			placeholder.textContent = LIST_CONTINUATION_CARET_PLACEHOLDER;
			return placeholder;
		};
		if (!(paragraph.textContent ?? '').replaceAll(LIST_CONTINUATION_CARET_PLACEHOLDER, '')) {
			paragraph.append(makeCaretPlaceholder());
		}

		const nextParagraph = paragraph.cloneNode(false) as HTMLElement;
		nextParagraph.removeAttribute('data-pptx-list-seg-idx');
		nextParagraph.removeAttribute('data-pptx-list-number');
		const caretPlaceholder = makeCaretPlaceholder();
		nextParagraph.append(caretPlaceholder, trailingContent);
		paragraph.insertAdjacentElement('afterend', nextParagraph);

		const caretText = caretPlaceholder.firstChild;
		if (!caretText) {
			return false;
		}
		const caretRange = document.createRange();
		caretRange.setStart(caretText, LIST_CONTINUATION_CARET_PLACEHOLDER.length);
		caretRange.collapse(true);
		selection.removeAllRanges();
		selection.addRange(caretRange);
		measureAndScaleText(false, true);
		return true;
	}, [measureAndScaleText]);

	// Native contentEditable Enter creates browser-specific block markup. In a
	// bullet paragraph Chromium commonly exposes that as two line breaks, which
	// remaps to an unintended empty PowerPoint paragraph and makes the list jump.
	// Clone the actual paragraph + marker DOM instead, so the live continuation
	// keeps the authored hanging indent and marker width. The commit path then
	// rebuilds it as structural OOXML (`a:buChar` / `a:buAutoNum`).
	const insertListParagraphBreak = useCallback((): boolean => {
		const editor = editorRef.current;
		const selection = window.getSelection();
		if (!editor || !selection || !selection.isCollapsed || selection.rangeCount === 0) {
			return false;
		}

		const range = selection.getRangeAt(0);
		if (!editor.contains(range.startContainer)) {
			return false;
		}
		const startElement =
			range.startContainer.nodeType === Node.ELEMENT_NODE
				? (range.startContainer as Element)
				: range.startContainer.parentElement;
		const paragraph = startElement?.closest<HTMLElement>('[data-pptx-paragraph]');
		if (!paragraph || !editor.contains(paragraph) || !hasTextProperties(element)) {
			return false;
		}
		const trailingRange = document.createRange();
		trailingRange.selectNodeContents(paragraph);
		trailingRange.setStart(range.startContainer, range.startOffset);

		const markerIndex = Number(paragraph.dataset.pptxListSegIdx);
		if (!Number.isInteger(markerIndex) || markerIndex < 0) {
			return false;
		}
		const currentListNumberValue = paragraph.dataset.pptxListNumber;
		const currentListNumber = Number(currentListNumberValue);
		const hasCurrentListNumber =
			currentListNumberValue !== undefined && Number.isFinite(currentListNumber);
		const marker = getListContinuationMarker(
			element.textSegments?.[markerIndex],
			hasCurrentListNumber ? currentListNumber : undefined,
		);
		const markerElement = paragraph.querySelector<HTMLElement>(
			`[data-seg-idx="${markerIndex}"]`,
		);
		if (!marker || !markerElement) {
			return false;
		}
		const sourceMarkerSegment = element.textSegments?.[markerIndex];
		const markerWidth = markerElement.style.width.trim();
		const isRtlList = window.getComputedStyle(paragraph).direction === 'rtl';
		const prepareParagraphBox = (target: HTMLElement): void => {
			// The editor is a column flex container. Explicit sizing is required for
			// an emptied source paragraph as well as its clone: after an Enter at the
			// beginning of an imported list item, Chromium otherwise shrinks the
			// original flex item to the number-marker width.
			target.style.alignSelf = 'stretch';
			target.style.width = 'auto';
			target.style.minWidth = '0';
			target.style.maxWidth = '100%';
			target.style.boxSizing = 'border-box';
		};
		prepareParagraphBox(paragraph);

		const leadingRange = document.createRange();
		leadingRange.selectNodeContents(paragraph);
		leadingRange.setEnd(range.startContainer, range.startOffset);
		const markerText = markerElement.textContent ?? marker;
		const leadingParagraphText = leadingRange
			.toString()
			.replaceAll(LIST_CONTINUATION_CARET_PLACEHOLDER, '');
		const leadingContentText = leadingParagraphText.startsWith(markerText)
			? leadingParagraphText.slice(markerText.length)
			: leadingParagraphText;
		const splitAtParagraphStart = leadingContentText.length === 0;
		const paragraphContentText = (paragraph.textContent ?? '')
			.replaceAll(LIST_CONTINUATION_CARET_PLACEHOLDER, '')
			.slice(markerText.length);
		if (pendingListParagraphsRef.current.has(paragraph) && !paragraphContentText.trim()) {
			// PowerPoint exits a list when Enter is pressed on an already-empty list
			// item. Keep one plain paragraph and its caret instead of cloning another
			// empty marker that can disturb the surrounding list on blur.
			markerElement.remove();
			paragraph.removeAttribute('data-pptx-list-seg-idx');
			paragraph.style.removeProperty('margin-left');
			paragraph.style.removeProperty('text-indent');
			const plainContent = paragraph.querySelector<HTMLElement>(
				'[data-pptx-list-continuation-content="true"]',
			);
			if (plainContent) {
				plainContent.style.display = 'block';
				plainContent.style.width = '100%';
				plainContent.style.minWidth = '0';
				plainContent.style.maxWidth = '100%';
				plainContent.style.textIndent = '0';
				plainContent.style.removeProperty('flex');
				plainContent.style.removeProperty('align-self');
			}
			paragraph.style.display = 'block';
			paragraph.style.width = 'auto';
			paragraph.style.removeProperty('flex-direction');
			paragraph.style.removeProperty('align-items');
			paragraph.style.removeProperty('gap');
			pendingListParagraphsRef.current.delete(paragraph);
			measureAndScaleText(false, true);
			return true;
		}

		const markerClone = markerElement.cloneNode(true) as HTMLElement;
		markerClone.textContent = marker;
		const activeSegment = startElement?.closest<HTMLElement>('[data-seg-idx]');
		const contentTemplate =
			activeSegment && activeSegment !== markerElement
				? activeSegment
				: Array.from(
						paragraph.querySelectorAll<HTMLElement>('[data-seg-idx]'),
					).find((node) => node !== markerElement);
		const prepareContentRun = (content: HTMLElement): HTMLElement => {
			content.dataset.pptxListContinuationContent = 'true';
			// Imported numbered paragraphs commonly include an empty formatting run
			// between the marker and visible text. Make the editable run the flexible
			// text column instead of relying on inline width arithmetic. Chromium can
			// otherwise reapply the paragraph's negative hanging indent while typing,
			// shifting the run under the marker and wrapping it every few characters.
			content.style.display = 'block';
			content.style.flex = '1 1 auto';
			content.style.alignSelf = 'stretch';
			content.style.width = 'auto';
			content.style.minWidth = '0';
			content.style.maxWidth = '100%';
			content.style.boxSizing = 'border-box';
			content.style.textIndent = '0';
			content.style.overflowWrap = 'break-word';
			content.style.wordBreak = 'normal';
			content.style.whiteSpace = 'pre-wrap';
			return content;
		};
		const prepareProvisionalListLayout = (
			target: HTMLElement,
			targetMarker: HTMLElement,
			content: HTMLElement,
		): void => {
			// Recreate the DrawingML hanging indent as two explicit flex columns.
			// The marker owns the indent width and the editable content owns all
			// remaining width. Resetting the paragraph's negative text-indent is the
			// important part: it prevents each newly typed line from jumping back to
			// the left edge of the shape.
			target.style.display = 'flex';
			target.style.flexDirection = isRtlList ? 'row-reverse' : 'row';
			target.style.alignItems = 'flex-start';
			target.style.gap = '0';
			target.style.width = '100%';
			target.style.marginLeft = '0';
			target.style.marginRight = '0';
			target.style.textIndent = '0';
			target.style.boxSizing = 'border-box';

			targetMarker.style.flex = markerWidth ? `0 0 ${markerWidth}` : '0 0 auto';
			if (markerWidth) {
				targetMarker.style.width = markerWidth;
				targetMarker.style.minWidth = markerWidth;
				targetMarker.style.maxWidth = markerWidth;
			}
			prepareContentRun(content);
		};
		// Keep the caret on the bullet's first line. A <br> placeholder makes
		// Chromium insert the first typed text after a forced line break, where
		// the hanging indent can leave only a marker-width column.
		const caretPlaceholder = document.createTextNode(LIST_CONTINUATION_CARET_PLACEHOLDER);
		// Extract everything after the caret before inserting the continuation.
		// This handles Enter in the middle of a bullet paragraph. Letting Chromium
		// split the contentEditable natively creates a nested block inside the run
		// span; that block inherits the marker-sized box and wraps every few letters.
		const trailingContent = trailingRange.extractContents();

		const nextParagraph = paragraph.cloneNode(false) as HTMLElement;
		prepareParagraphBox(nextParagraph);
		if (hasCurrentListNumber) {
			nextParagraph.dataset.pptxListNumber = String(currentListNumber + 1);
		}

		if (splitAtParagraphStart) {
			// PowerPoint keeps the caret in the newly-created empty bullet before
			// the previous text. Reuse the original run for that caret and move the
			// complete old content into the following paragraph.
			const currentContent = prepareContentRun(
				contentTemplate && paragraph.contains(contentTemplate)
					? contentTemplate
					: document.createElement('span'),
			);
			currentContent.replaceChildren(caretPlaceholder);
			if (!paragraph.contains(currentContent)) {
				paragraph.append(currentContent);
			}
			prepareProvisionalListLayout(paragraph, markerElement, currentContent);
			nextParagraph.append(markerClone, trailingContent);
			pendingListParagraphsRef.current.add(paragraph);
		} else {
			// At the middle/end, the current paragraph retains its leading text and
			// the caret starts the continuation immediately before the moved tail.
			const continuationContent = prepareContentRun(
				contentTemplate
					? (contentTemplate.cloneNode(false) as HTMLElement)
					: document.createElement('span'),
			);
			continuationContent.replaceChildren(caretPlaceholder, trailingContent);
			nextParagraph.append(markerClone, continuationContent);
			prepareProvisionalListLayout(nextParagraph, markerClone, continuationContent);
			pendingListParagraphsRef.current.add(nextParagraph);
		}
		paragraph.insertAdjacentElement('afterend', nextParagraph);

		// Keep the live numbering correct immediately. The structural remap repeats
		// this during commit/save, but without this pass the following authored item
		// temporarily keeps the same visible number as the inserted continuation.
		if (sourceMarkerSegment?.bulletInfo?.autoNumType) {
			let following = nextParagraph.nextElementSibling as HTMLElement | null;
			while (following?.hasAttribute('data-pptx-paragraph')) {
				const followingMarkerIndex = Number(following.dataset.pptxListSegIdx);
				const followingSource = Number.isInteger(followingMarkerIndex)
					? element.textSegments?.[followingMarkerIndex]
					: undefined;
				if (
					followingSource?.bulletInfo?.autoNumType !==
					sourceMarkerSegment.bulletInfo.autoNumType
				) {
					break;
				}
				const followingMarker = following.querySelector<HTMLElement>(
					`[data-seg-idx="${followingMarkerIndex}"]`,
				);
				const followingListNumberValue = following.dataset.pptxListNumber;
				const followingListNumber = Number(followingListNumberValue);
				const hasFollowingListNumber =
					followingListNumberValue !== undefined && Number.isFinite(followingListNumber);
				const incrementedMarker = getListContinuationMarker(
					followingSource,
					hasFollowingListNumber ? followingListNumber : undefined,
				);
				if (followingMarker && incrementedMarker) {
					followingMarker.textContent = incrementedMarker;
					if (hasFollowingListNumber) {
						following.dataset.pptxListNumber = String(followingListNumber + 1);
					}
				}
				following = following.nextElementSibling as HTMLElement | null;
			}
		}

		const caretRange = document.createRange();
		caretRange.setStart(caretPlaceholder, LIST_CONTINUATION_CARET_PLACEHOLDER.length);
		caretRange.collapse(true);
		selection.removeAllRanges();
		selection.addRange(caretRange);
		measureAndScaleText(false, true);
		return true;
	}, [element, measureAndScaleText]);

	const removeEmptyPendingListParagraphs = useCallback(() => {
		for (const paragraph of pendingListParagraphsRef.current) {
			const markerIndex = paragraph.dataset.pptxListSegIdx;
			const marker = markerIndex !== undefined
				? paragraph.querySelector<HTMLElement>(`[data-seg-idx="${markerIndex}"]`)
				: null;
			const paragraphText = (paragraph.textContent ?? '').replaceAll(
				LIST_CONTINUATION_CARET_PLACEHOLDER,
				'',
			);
			const markerText = marker?.textContent ?? '';
			const contentText = paragraphText.startsWith(markerText)
				? paragraphText.slice(markerText.length)
				: paragraphText;
			if (!contentText.trim()) {
				paragraph.remove();
			}
		}
		pendingListParagraphsRef.current.clear();
	}, []);

	// Auto-focus on mount and place cursor at end
	useEffect(() => {
		const el = editorRef.current;
		if (!el) {
			return;
		}
		el.focus();
		// Place cursor at end of content
		const selection = window.getSelection();
		if (selection) {
			const range = document.createRange();
			range.selectNodeContents(el);
			range.collapse(false);
			selection.removeAllRanges();
			selection.addRange(range);
		}
	}, []);

	// Initial fit only. Preserve the saved on-slide size when it already fits and
	// shrink immediately when it does not. Automatically growing to 100% on entry
	// makes the editor look zoomed-in and hides lower paragraphs until another
	// input event; deletion still restores size gradually through `handleInput`.
	useLayoutEffect(() => {
		measureAndScaleText(false);
	}, [measureAndScaleText]);

	// After a formatting update, React re-renders the contentEditable children
	// which destroys the DOM selection. Restore it from the pending info.
	const mountedRef = useRef(false);
	useLayoutEffect(() => {
		// Skip the initial mount; cursor is already placed by the effect above.
		if (!mountedRef.current) {
			mountedRef.current = true;
			return;
		}
		const pending = getPendingSelectionRestore();
		if (!pending || !editorRef.current) {
			return;
		}
		restoreSegmentSelection(
			editorRef.current,
			pending.startSegIdx,
			pending.startOffset,
			pending.endSegIdx,
			pending.endOffset,
		);
	});

	// Build wrapper style matching view-mode exactly:
	// layoutStyle (flex alignment, vertical padding, columns) + textStyle (font defaults,
	// horizontal padding/insets) + warp transforms + compensation transform.
	//
	// View mode applies: getTextLayoutStyle + txtS + getTextWarpStyle + compensationTransform
	// We replicate that same order here.
	const warpStyle = getTextWarpStyle(textStyleRaw);

	// Merge the compensation transform with warp transform if both exist
	const compensationTransform = getTextCompensationTransform(element);
	const warpTransform = warpStyle?.transform;
	const mergedTransform =
		[compensationTransform, warpTransform].filter(Boolean).join(' ') || undefined;

	const wrapperStyle: React.CSSProperties = {
		...layoutStyle,
		...textStyle,
		...warpStyle,
		transform: mergedTransform,
		transformOrigin: warpStyle?.transformOrigin || 'center',
	};

	return (
		<div
			ref={editorRef}
			contentEditable
			suppressContentEditableWarning
			data-inline-editor
			spellCheck={spellCheck}
			dir={rtl ? 'rtl' : 'ltr'}
			className='relative z-10 w-full h-full whitespace-pre-wrap break-words leading-[1.3] outline-none'
			style={{
				...wrapperStyle,
				cursor: 'text',
				minHeight: '1em',
				whiteSpace: 'pre-wrap',
				overflowWrap: 'break-word',
				wordBreak: 'normal',
				overflow: 'hidden',
			}}
			// Touch surfaces drive canvas drag/marquee through onPointerDown (see
			// useCanvasEventHandlers.handleStagePointerDown). Without stopping it
			// here, tapping inside the editor to reposition the caret would bubble
			// to the stage and start dragging the element instead of editing.
			onPointerDown={(e) => e.stopPropagation()}
			onMouseDown={(e) => e.stopPropagation()}
			onClick={(e) => e.stopPropagation()}
			onInput={handleInput}
			onBlur={() => {
				removeEmptyPendingListParagraphs();
				removeListContinuationCaretPlaceholder();
				// The final view rebuilds the contentEditable DOM as DrawingML
				// paragraphs. Its line-box metrics can be a few pixels taller than the
				// browser's temporary typing DOM. Reserve a small bottom inset and finish
				// the fit search before switching renderers. Keystrokes still shrink one
				// normal two-unit step at a time; blur must resolve any remaining real
				// overflow from paste, replacement, or a large single edit so view mode
				// cannot place text above or below the fixed PowerPoint shape.
				const finalAutoFitScale = measureAndScaleText(
					false,
					false,
					VIEW_MODE_FIT_SAFETY_PX,
				);
				const committedText = extractText();
				onEditChange(committedText, finalAutoFitScale);
				onCommit(undefined, committedText, finalAutoFitScale);
			}}
			onKeyDown={(e) => {
				// Inline formatting shortcuts (Ctrl/Cmd + B/I/U)
				if ((e.ctrlKey || e.metaKey) && !e.shiftKey && onFormatText) {
					const key = e.key.toLowerCase();
					if (key === 'b' || key === 'i' || key === 'u') {
						e.preventDefault();
						e.stopPropagation();
						const seg = hasTextProperties(element) ? element.textSegments?.[0] : undefined;
						const ts = seg?.style ?? (hasTextProperties(element) ? element.textStyle : undefined);
						switch (key) {
							case 'b':
								onFormatText({ bold: !ts?.bold });
								break;
							case 'i':
								onFormatText({ italic: !ts?.italic });
								break;
							case 'u':
								onFormatText({ underline: !ts?.underline });
								break;
						}
						return;
					}
				}
				if (e.key === 'Escape') {
					e.preventDefault();
					onCancel();
					return;
				}
				if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
					e.preventDefault();
					removeEmptyPendingListParagraphs();
					removeListContinuationCaretPlaceholder();
					const finalAutoFitScale = measureAndScaleText(
						false,
						false,
						VIEW_MODE_FIT_SAFETY_PX,
					);
					const committedText = extractText();
					onEditChange(committedText, finalAutoFitScale);
					onCommit(undefined, committedText, finalAutoFitScale);
					return;
				}
				if (e.key === 'Enter') {
					trimTrailingSpaceBeforeCaret();
					if (!e.shiftKey) {
						const insertedParagraph =
							insertListParagraphBreak() || insertPlainParagraphBreak();
						if (insertedParagraph) {
							e.preventDefault();
							e.stopPropagation();
							const liveAutoFitScale = measureAndScaleText(false, true);
							onEditChange(extractText(), liveAutoFitScale);
						}
					}
				}
			}}
			// Prevent paste from inserting HTML: paste as plain text only
			onPaste={(e) => {
				e.preventDefault();
				const text = e.clipboardData.getData('text/plain');
				document.execCommand('insertText', false, text);
			}}
		>
			{seed.hasRichSegments ? renderTextSegments(element, DEFAULT_TEXT_COLOR) : seed.initialText}
		</div>
	);
}

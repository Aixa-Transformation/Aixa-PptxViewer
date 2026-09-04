/**
 * Utilities for reading and restoring text selections within the inline text
 * editor contentEditable, and for applying style updates to a subset of text
 * segments based on the current selection. Framework-agnostic (uses DOM globals
 * only; no framework imports).
 */
import type { TextSegment, TextStyle } from 'pptx-viewer-core';

/** Describes which segments (and offsets within them) are selected. */
export interface InlineTextSelection {
	startSegIdx: number;
	startOffset: number;
	endSegIdx: number;
	endOffset: number;
	/**
	 * Live paragraph coordinates from the contentEditable DOM. Segment indexes
	 * alone can be stale while the user is typing because the editor deliberately
	 * keeps its DOM uncontrolled until commit. List commands use these coordinates
	 * to target newly inserted paragraphs before that commit happens.
	 */
	startParagraphIndex?: number;
	startParagraphOffset?: number;
	endParagraphIndex?: number;
	endParagraphOffset?: number;
	/** True when the normalized selection ends before any authored text in a paragraph. */
	endAtParagraphStart?: boolean;
}

let pendingRestore: InlineTextSelection | null = null;

export function setPendingSelectionRestore(sel: InlineTextSelection | null): void {
	pendingRestore = sel;
}

/** Read-once: clears the stored value after returning it. */
export function getPendingSelectionRestore(): InlineTextSelection | null {
	const sel = pendingRestore;
	pendingRestore = null;
	return sel;
}

/**
 * Read the current browser selection and, if it falls within an
 * `[data-inline-editor]` element, return the segment-level range.
 *
 * Returns `null` when there is no selection, the selection is collapsed
 * (just a cursor), or the selection is outside the inline editor.
 */
export function getInlineEditorSelection(
	segments: TextSegment[] | undefined,
): InlineTextSelection | null {
	if (!segments || segments.length === 0) {
		return null;
	}

	const sel = window.getSelection();
	if (!sel || sel.rangeCount === 0 || sel.isCollapsed) {
		return null;
	}

	const { anchorNode, anchorOffset, focusNode, focusOffset } = sel;
	if (!anchorNode || !focusNode) {
		return null;
	}

	const editor = findEditorContainer(anchorNode);
	if (!editor || !editor.contains(focusNode)) {
		return null;
	}

	const anchorInfo = getSegmentPosition(editor, anchorNode, anchorOffset, segments);
	const focusInfo = getSegmentPosition(editor, focusNode, focusOffset, segments);
	if (!anchorInfo || !focusInfo) {
		return null;
	}

	// Normalize so start <= end.
	const [start, end] =
		anchorInfo.absOffset <= focusInfo.absOffset ? [anchorInfo, focusInfo] : [focusInfo, anchorInfo];

	return {
		startSegIdx: start.segIdx,
		startOffset: start.offsetInSeg,
		endSegIdx: end.segIdx,
		endOffset: end.offsetInSeg,
		...(start.paragraphIndex !== undefined
			? {
					startParagraphIndex: start.paragraphIndex,
					startParagraphOffset: start.paragraphOffset ?? 0,
				}
			: {}),
		...(end.paragraphIndex !== undefined
			? {
					endParagraphIndex: end.paragraphIndex,
					endParagraphOffset: end.paragraphOffset ?? 0,
					endAtParagraphStart: end.atParagraphStart === true,
				}
			: {}),
	};
}

function findEditorContainer(node: Node): Element | null {
	const el = node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
	return el?.closest('[data-inline-editor]') ?? null;
}

function getSegmentPosition(
	editor: Element,
	node: Node,
	offset: number,
	segments: TextSegment[],
): {
	segIdx: number;
	offsetInSeg: number;
	absOffset: number;
	paragraphIndex?: number;
	paragraphOffset?: number;
	atParagraphStart?: boolean;
} | null {
	const el = node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
	if (!el) {
		return null;
	}

	const segSpan = el.closest('[data-seg-idx]');
	if (!segSpan || !editor.contains(segSpan)) {
		return null;
	}

	const segIdx = parseInt(segSpan.getAttribute('data-seg-idx')!, 10);
	if (isNaN(segIdx) || segIdx < 0 || segIdx >= segments.length) {
		return null;
	}

	const offsetInSeg = getTextOffsetWithin(segSpan, node, offset);

	// Absolute offset = sum of all segment text lengths before this one + offset.
	let absOffset = 0;
	for (let i = 0; i < segIdx; i++) {
		absOffset += segments[i].text.length;
	}
	absOffset += offsetInSeg;

	const paragraphPosition = getParagraphPosition(editor, node, offset);
	return { segIdx, offsetInSeg, absOffset, ...paragraphPosition };
}

function getTopLevelParagraphs(editor: Element): HTMLElement[] {
	return Array.from(editor.querySelectorAll<HTMLElement>('[data-pptx-paragraph]')).filter(
		(paragraph) => !paragraph.parentElement?.closest('[data-pptx-paragraph]'),
	);
}

/**
 * Resolve a DOM point against the paragraph wrappers rendered by
 * `renderTextSegments`. The offset deliberately excludes the visible synthetic
 * list marker so it can be projected onto freshly remapped text segments.
 */
function getParagraphPosition(
	editor: Element,
	node: Node,
	offset: number,
): {
	paragraphIndex?: number;
	paragraphOffset?: number;
	atParagraphStart?: boolean;
} {
	const element = node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
	const paragraph = element?.closest<HTMLElement>('[data-pptx-paragraph]');
	if (!paragraph || !editor.contains(paragraph)) {
		return {};
	}
	const paragraphs = getTopLevelParagraphs(editor);
	const paragraphIndex = paragraphs.indexOf(paragraph);
	if (paragraphIndex < 0) {
		return {};
	}

	let prefixText = '';
	try {
		const prefixRange = document.createRange();
		prefixRange.selectNodeContents(paragraph);
		prefixRange.setEnd(node, offset);
		const prefixContainer = document.createElement('div');
		prefixContainer.append(prefixRange.cloneContents());
		const authoredSegments = Array.from(
			prefixContainer.querySelectorAll<HTMLElement>('[data-seg-idx]'),
		).filter((segment) => !segment.parentElement?.closest('[data-seg-idx]'));
		// Ignore formatting whitespace between JSX elements. Authored text always
		// lives inside a segment span, including provisional list continuations.
		prefixText = authoredSegments.map((segment) => segment.textContent ?? '').join('');
	} catch {
		return { paragraphIndex };
	}
	const markerIndex = paragraph.dataset.pptxListSegIdx;
	const marker = markerIndex !== undefined
		? paragraph.querySelector<HTMLElement>(`[data-seg-idx="${markerIndex}"]`)
		: null;
	const markerText = marker?.textContent ?? '';
	if (markerText && prefixText.startsWith(markerText)) {
		prefixText = prefixText.slice(markerText.length);
	}
	const authoredPrefix = prefixText.replaceAll('\u200B', '');
	return {
		paragraphIndex,
		paragraphOffset: authoredPrefix.length,
		atParagraphStart: authoredPrefix.length === 0,
	};
}

/**
 * Compute the character offset of a DOM position (node + offset) relative to
 * the text content of a container element.
 */
function getTextOffsetWithin(container: Element, targetNode: Node, targetOffset: number): number {
	// When the target IS the container (or an element child), offset is a
	// child-index count, not a character count.
	if (targetNode === container || targetNode.nodeType === Node.ELEMENT_NODE) {
		const parent = targetNode === container ? container : targetNode;
		let count = 0;
		for (let i = 0; i < targetOffset && i < parent.childNodes.length; i++) {
			count += (parent.childNodes[i].textContent || '').length;
		}
		return count;
	}

	// Walk text nodes in document order and accumulate lengths until we hit the
	// target text node.
	const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
	let charCount = 0;
	let node: Node | null;
	while ((node = walker.nextNode())) {
		if (node === targetNode) {
			return charCount + targetOffset;
		}
		charCount += (node as Text).length;
	}
	return charCount;
}

/**
 * Apply `updates` only to the characters within the selection range. Segments
 * at the boundaries are split; segments outside the range are copied unchanged.
 * Paragraph-break segments (`\n`) are always passed through unchanged.
 *
 * Returns the new segments array AND the selection coordinates mapped to the
 * new segment indices (for cursor restoration).
 */
export function applyStyleToSelectedSegments(
	segments: TextSegment[],
	selection: InlineTextSelection,
	updates: Partial<TextStyle>,
): { newSegments: TextSegment[]; newSelection: InlineTextSelection } {
	const { startSegIdx, startOffset, endSegIdx, endOffset } = selection;
	const result: TextSegment[] = [];

	// Track new indices for the start/end of the selected text in the new array.
	let newStartSegIdx = -1;
	let newStartOffset = 0;
	let newEndSegIdx = -1;
	let newEndOffset = 0;

	for (let i = 0; i < segments.length; i++) {
		const seg = segments[i];

		// Paragraph-break segments always pass through.
		if (seg.isParagraphBreak || seg.text === '\n') {
			result.push({ ...seg, style: { ...seg.style } });
			continue;
		}

		// Before selection range.
		if (i < startSegIdx) {
			result.push({ ...seg, style: { ...seg.style } });
			continue;
		}

		// After selection range.
		if (i > endSegIdx) {
			result.push({ ...seg, style: { ...seg.style } });
			continue;
		}

		// Single-segment selection (start and end in the same segment).
		if (i === startSegIdx && i === endSegIdx) {
			splitSingle(seg, startOffset, endOffset, updates, result);
			// Track the new selection: the selected part is at `result.length - 1`
			// (or `result.length - 2` if there's an after part).
			const partsAdded = (startOffset > 0 ? 1 : 0) + 1 + (endOffset < seg.text.length ? 1 : 0);
			const selectedPartIdx = result.length - partsAdded + (startOffset > 0 ? 1 : 0);
			newStartSegIdx = selectedPartIdx;
			newStartOffset = 0;
			newEndSegIdx = selectedPartIdx;
			newEndOffset = endOffset - startOffset;
			continue;
		}

		// Start segment of multi-segment selection.
		if (i === startSegIdx) {
			if (startOffset === 0) {
				newStartSegIdx = result.length;
				newStartOffset = 0;
				result.push({ ...seg, style: { ...seg.style, ...updates } });
			} else {
				// Before-selection part.
				const beforeSeg: TextSegment = {
					text: seg.text.slice(0, startOffset),
					style: { ...seg.style },
				};
				if (seg.bulletInfo) {
					beforeSeg.bulletInfo = seg.bulletInfo;
				}
				result.push(beforeSeg);

				// Selected part.
				newStartSegIdx = result.length;
				newStartOffset = 0;
				result.push({
					text: seg.text.slice(startOffset),
					style: { ...seg.style, ...updates },
				});
			}
			continue;
		}

		// End segment of multi-segment selection.
		if (i === endSegIdx) {
			if (endOffset >= seg.text.length) {
				result.push({ ...seg, style: { ...seg.style, ...updates } });
				newEndSegIdx = result.length - 1;
				newEndOffset = seg.text.length;
			} else {
				// Selected part.
				result.push({
					text: seg.text.slice(0, endOffset),
					style: { ...seg.style, ...updates },
				});
				newEndSegIdx = result.length - 1;
				newEndOffset = endOffset;

				// After-selection part.
				result.push({
					text: seg.text.slice(endOffset),
					style: { ...seg.style },
				});
			}
			continue;
		}

		// Middle segment: entirely within selection.
		result.push({ ...seg, style: { ...seg.style, ...updates } });
	}

	return {
		newSegments: result,
		newSelection: {
			startSegIdx: newStartSegIdx,
			startOffset: newStartOffset,
			endSegIdx: newEndSegIdx,
			endOffset: newEndOffset,
		},
	};
}

function splitSingle(
	seg: TextSegment,
	startOffset: number,
	endOffset: number,
	updates: Partial<TextStyle>,
	result: TextSegment[],
): void {
	// Entire segment selected.
	if (startOffset === 0 && endOffset >= seg.text.length) {
		result.push({ ...seg, style: { ...seg.style, ...updates } });
		return;
	}

	// Before-selection part.
	if (startOffset > 0) {
		const beforeSeg: TextSegment = {
			text: seg.text.slice(0, startOffset),
			style: { ...seg.style },
		};
		if (seg.bulletInfo) {
			beforeSeg.bulletInfo = seg.bulletInfo;
		}
		result.push(beforeSeg);
	}

	// Selected part.
	result.push({
		text: seg.text.slice(startOffset, endOffset),
		style: { ...seg.style, ...updates },
	});

	// After-selection part.
	if (endOffset < seg.text.length) {
		result.push({
			text: seg.text.slice(endOffset),
			style: { ...seg.style },
		});
	}
}

/**
 * Set the DOM selection inside the inline editor using segment-level
 * coordinates. Finds `[data-seg-idx]` spans and navigates to the correct
 * text-node offsets.
 */
export function restoreSegmentSelection(
	editor: HTMLElement,
	startSegIdx: number,
	startOffset: number,
	endSegIdx: number,
	endOffset: number,
): void {
	const sel = window.getSelection();
	if (!sel) {
		return;
	}

	const startPos = findTextPositionInSegment(editor, startSegIdx, startOffset);
	const endPos = findTextPositionInSegment(editor, endSegIdx, endOffset);
	if (!startPos || !endPos) {
		return;
	}

	const range = document.createRange();
	range.setStart(startPos.node, startPos.offset);
	range.setEnd(endPos.node, endPos.offset);
	sel.removeAllRanges();
	sel.addRange(range);
}

interface DomPosition {
	node: Node;
	offset: number;
}

function findTextPositionInSegment(
	editor: HTMLElement,
	segIdx: number,
	charOffset: number,
): DomPosition | null {
	const segSpan = editor.querySelector(`[data-seg-idx="${segIdx}"]`);
	if (!segSpan) {
		return null;
	}

	const walker = document.createTreeWalker(segSpan, NodeFilter.SHOW_TEXT);
	let charCount = 0;
	let node: Node | null;
	while ((node = walker.nextNode())) {
		const textLen = (node as Text).length;
		if (charCount + textLen >= charOffset) {
			return { node, offset: charOffset - charCount };
		}
		charCount += textLen;
	}

	// If exact position not found, place at end of last text node.
	const lastTextNode = segSpan.querySelector('*') ? null : segSpan.firstChild;
	if (lastTextNode && lastTextNode.nodeType === Node.TEXT_NODE) {
		return { node: lastTextNode, offset: (lastTextNode as Text).length };
	}

	return null;
}

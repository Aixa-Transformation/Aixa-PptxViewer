import { hasTextProperties } from 'pptx-viewer-core';
import type { BulletInfo, PptxElement, TextSegment, TextStyle } from 'pptx-viewer-core';

import { formatAutoNumber } from './bullet-autonum';
import type { InlineTextSelection } from './inline-selection-utils';

export type ParagraphListType = NonNullable<TextStyle['listType']>;
export type ElementListType = ParagraphListType | 'mixed';

interface IndexedSegment {
	segment: TextSegment;
	oldIndex: number;
}

interface SegmentParagraph {
	items: IndexedSegment[];
	breakItem?: IndexedSegment;
}

export interface ApplyParagraphListTypeInput {
	text?: string;
	textSegments?: TextSegment[];
	fallbackStyle?: TextStyle;
	listType: ParagraphListType;
	selection?: InlineTextSelection | null;
}

export interface ApplyParagraphListTypeResult {
	text: string;
	textSegments: TextSegment[];
	selection?: InlineTextSelection;
}

function isParagraphBreak(segment: TextSegment): boolean {
	return Boolean(segment.isParagraphBreak || (segment.text === '\n' && !segment.isLineBreak));
}

/** Detect a metadata-only list marker produced by the PPTX loader or this editor. */
export function isSyntheticListMarkerSegment(segment: TextSegment): boolean {
	const markerText = String(segment.text ?? '').trim();
	return Boolean(
		segment.bulletInfo &&
		((segment.bulletInfo.none && markerText.length === 0) ||
			(segment.bulletInfo.char && markerText === segment.bulletInfo.char) ||
			(segment.bulletInfo.autoNumType && /^\S+[.)]$/u.test(markerText)) ||
			((segment.bulletInfo.imageDataUrl || segment.bulletInfo.imageRelId) &&
				markerText === '\u{1F4CE}')),
	);
}

function cloneSegment(segment: TextSegment): TextSegment {
	return {
		...segment,
		style: { ...segment.style },
		...(segment.bulletInfo ? { bulletInfo: { ...segment.bulletInfo } } : {}),
		...(segment.paragraphProperties
			? { paragraphProperties: { ...segment.paragraphProperties } }
			: {}),
	};
}

function segmentsFromPlainText(text: string, fallbackStyle: TextStyle): TextSegment[] {
	const lines = text.split('\n');
	const segments: TextSegment[] = [];
	lines.forEach((line, index) => {
		segments.push({ text: line, style: { ...fallbackStyle } });
		if (index < lines.length - 1) {
			segments.push({ text: '\n', style: { ...fallbackStyle }, isParagraphBreak: true });
		}
	});
	return segments;
}

function groupIntoParagraphs(segments: TextSegment[]): SegmentParagraph[] {
	const paragraphs: SegmentParagraph[] = [{ items: [] }];
	segments.forEach((segment, oldIndex) => {
		if (isParagraphBreak(segment)) {
			paragraphs[paragraphs.length - 1].breakItem = { segment, oldIndex };
			paragraphs.push({ items: [] });
			return;
		}
		paragraphs[paragraphs.length - 1].items.push({ segment, oldIndex });
	});
	return paragraphs;
}

function copyParagraphMetadata(from: TextSegment | undefined, to: TextSegment): void {
	if (!from) {
		return;
	}
	if (from.paragraphLevel !== undefined) {
		to.paragraphLevel = from.paragraphLevel;
	}
	if (from.endParaRunProperties !== undefined) {
		to.endParaRunProperties = from.endParaRunProperties;
	}
	if (from.paragraphProperties !== undefined) {
		to.paragraphProperties = { ...from.paragraphProperties };
	}
}

function removeParagraphMetadata(segment: TextSegment): void {
	delete segment.paragraphLevel;
	delete segment.endParaRunProperties;
	delete segment.paragraphProperties;
}

function makeMarkerSegment(
	listType: ParagraphListType,
	style: TextStyle,
	paragraphNumber: number,
	metadataSource: TextSegment | undefined,
): TextSegment {
	let text = '';
	let bulletInfo: BulletInfo;
	if (listType === 'bullet') {
		text = '\u2022 ';
		bulletInfo = { char: '\u2022' };
	} else if (listType === 'numbered') {
		bulletInfo = {
			autoNumType: 'arabicPeriod',
			autoNumStartAt: 1,
			paragraphIndex: paragraphNumber - 1,
		};
		text = `${formatAutoNumber('arabicPeriod', paragraphNumber)} `;
	} else {
		bulletInfo = { none: true };
	}

	const marker: TextSegment = {
		text,
		style: { ...style, listType },
		bulletInfo,
	};
	copyParagraphMetadata(metadataSource, marker);
	return marker;
}

function classifyParagraph(paragraph: SegmentParagraph): ParagraphListType {
	const first = paragraph.items[0]?.segment;
	const bullet = first?.bulletInfo;
	if (bullet?.none || first?.style?.listType === 'none') {
		return 'none';
	}
	if (bullet?.autoNumType || first?.style?.listType === 'numbered') {
		return 'numbered';
	}
	if (
		bullet?.char ||
		bullet?.imageDataUrl ||
		bullet?.imageRelId ||
		first?.style?.listType === 'bullet'
	) {
		return 'bullet';
	}
	return 'none';
}

/** Return the structural list state shown by an element's paragraphs. */
export function getElementListType(element: PptxElement | null | undefined): ElementListType {
	if (!element || !hasTextProperties(element)) {
		return 'none';
	}
	if (!element.textSegments || element.textSegments.length === 0) {
		return element.textStyle?.listType ?? 'none';
	}
	const modes = groupIntoParagraphs(element.textSegments).map(classifyParagraph);
	const firstMode = modes[0] ?? element.textStyle?.listType ?? 'none';
	return modes.every((mode) => mode === firstMode) ? firstMode : 'mixed';
}

function paragraphForOldIndex(
	paragraphs: SegmentParagraph[],
	oldIndex: number,
	fallback: number,
): number {
	for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
		const paragraph = paragraphs[paragraphIndex];
		if (
			paragraph.items.some((item) => item.oldIndex === oldIndex) ||
			paragraph.breakItem?.oldIndex === oldIndex
		) {
			return paragraphIndex;
		}
	}
	return fallback;
}

/**
 * Apply a bulleted, numbered, or explicit no-list state to whole paragraphs.
 * The returned segments carry real DrawingML bullet metadata, so React renders
 * the marker and the PPTX writer emits `a:buChar`, `a:buAutoNum`, or `a:buNone`.
 */
export function applyParagraphListType({
	text = '',
	textSegments,
	fallbackStyle = {},
	listType,
	selection,
}: ApplyParagraphListTypeInput): ApplyParagraphListTypeResult {
	const sourceSegments =
		textSegments && textSegments.length > 0
			? textSegments
			: segmentsFromPlainText(text, fallbackStyle);
	const paragraphs = groupIntoParagraphs(sourceSegments);
	const startParagraph = selection ? paragraphForOldIndex(paragraphs, selection.startSegIdx, 0) : 0;
	const endParagraph = selection
		? paragraphForOldIndex(paragraphs, selection.endSegIdx, paragraphs.length - 1)
		: paragraphs.length - 1;
	const firstTarget = Math.min(startParagraph, endParagraph);
	const lastTarget = Math.max(startParagraph, endParagraph);
	const output: TextSegment[] = [];
	const oldToNew = new Map<number, number>();
	let numberedParagraph = 0;

	paragraphs.forEach((paragraph, paragraphIndex) => {
		const isTarget = paragraphIndex >= firstTarget && paragraphIndex <= lastTarget;
		if (!isTarget) {
			paragraph.items.forEach((item) => {
				oldToNew.set(item.oldIndex, output.length);
				output.push(cloneSegment(item.segment));
			});
		} else {
			const metadataSource = paragraph.items[0]?.segment;
			const content: Array<{ segment: TextSegment; oldIndex: number }> = [];
			const removedMarkerIndices: number[] = [];
			paragraph.items.forEach((item) => {
				if (isSyntheticListMarkerSegment(item.segment)) {
					removedMarkerIndices.push(item.oldIndex);
					return;
				}
				const segment = cloneSegment(item.segment);
				delete segment.bulletInfo;
				delete segment.style.listType;
				content.push({ segment, oldIndex: item.oldIndex });
			});
			if (content.length === 0) {
				content.push({ segment: { text: '', style: { ...fallbackStyle } }, oldIndex: -1 });
			}
			removeParagraphMetadata(content[0].segment);
			if (listType === 'numbered') {
				numberedParagraph += 1;
			}
			const markerIndex = output.length;
			output.push(
				makeMarkerSegment(
					listType,
					metadataSource?.style ?? content[0].segment.style ?? fallbackStyle,
					numberedParagraph,
					metadataSource,
				),
			);
			removedMarkerIndices.forEach((oldIndex) => oldToNew.set(oldIndex, markerIndex));
			content.forEach(({ segment, oldIndex }) => {
				if (oldIndex >= 0) {
					oldToNew.set(oldIndex, output.length);
				}
				output.push(segment);
			});
		}

		if (paragraph.breakItem) {
			oldToNew.set(paragraph.breakItem.oldIndex, output.length);
			output.push(cloneSegment(paragraph.breakItem.segment));
		}
	});

	let mappedSelection: InlineTextSelection | undefined;
	if (selection) {
		const newStartSegIdx = oldToNew.get(selection.startSegIdx) ?? 0;
		const newEndSegIdx = oldToNew.get(selection.endSegIdx) ?? newStartSegIdx;
		mappedSelection = {
			startSegIdx: newStartSegIdx,
			startOffset: Math.min(selection.startOffset, output[newStartSegIdx]?.text.length ?? 0),
			endSegIdx: newEndSegIdx,
			endOffset: Math.min(selection.endOffset, output[newEndSegIdx]?.text.length ?? 0),
		};
	}

	return {
		text: output.map((segment) => segment.text).join(''),
		textSegments: output,
		...(mappedSelection ? { selection: mappedSelection } : {}),
	};
}

/**
 * Remap edited plain-text back onto original rich-text segments, preserving
 * per-segment styles (font, size, colour, bold, italic, …). Pure,
 * framework-agnostic logic shared by every binding (React / Vue / Angular).
 */
import type { TextSegment, TextStyle } from 'pptx-viewer-core';

import { formatAutoNumber } from './bullet-autonum';
import { isSyntheticListMarkerSegment } from './text-list-edit';

export interface ParagraphIndent {
	marginLeft?: number;
	indent?: number;
}

/**
 * Copy segment-level metadata (equation, field) from an original segment onto
 * its remapped counterpart. Without this, entering and leaving inline text
 * editing destroys the data these fields carry even when nothing was typed:
 * an equation collapses to its literal "[Equation]" placeholder text and a
 * slide-number/date field degrades to frozen plain text. Hyperlink and other
 * style-level properties already survive via the copied `style` object.
 */
function copySegmentMetadata(from: TextSegment, to: TextSegment): TextSegment {
	if (from.equationXml !== undefined) {
		to.equationXml = from.equationXml;
	}
	if (from.equationNumber !== undefined) {
		to.equationNumber = from.equationNumber;
	}
	if (from.fieldType !== undefined) {
		to.fieldType = from.fieldType;
	}
	if (from.fieldGuid !== undefined) {
		to.fieldGuid = from.fieldGuid;
	}
	if (from.fieldGuidAttr !== undefined) {
		to.fieldGuidAttr = from.fieldGuidAttr;
	}
	if (from.fieldParagraphPropertiesXml !== undefined) {
		to.fieldParagraphPropertiesXml = from.fieldParagraphPropertiesXml;
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
	return to;
}

function splitOriginalParagraphs(originalSegments: TextSegment[]): TextSegment[][] {
	const paragraphs: TextSegment[][] = [[]];
	for (const segment of originalSegments) {
		if (segment.text === '\n' || segment.isParagraphBreak) {
			paragraphs.push([]);
		} else {
			paragraphs[paragraphs.length - 1].push(segment);
		}
	}
	return paragraphs;
}

function normalizeParagraphForMatching(text: string, stripListMarker: boolean): string {
	const content = stripListMarker
		? text.replace(/^\s*(?:(?:\S+[.)])|[\u2022\u25CF\u25A0\u25AA\u2013\u2014\-\u{1F4CE}])\s+/u, '')
		: text;
	return content.replace(/\s+/gu, ' ').trim();
}

/**
 * Match edited paragraphs back to their authored source paragraphs. An exact
 * LCS keeps unchanged paragraphs attached to their original style/indent when
 * Enter inserts a paragraph in the middle; unmatched inserted paragraphs
 * inherit the closest preceding paragraph (or the following one at the start).
 */
function mapEditedParagraphsToOriginal(
	newParagraphTexts: string[],
	originalParagraphs: TextSegment[][],
): number[] {
	if (originalParagraphs.length === 0) {
		return newParagraphTexts.map(() => 0);
	}
	if (newParagraphTexts.length === originalParagraphs.length) {
		return newParagraphTexts.map((_, index) => index);
	}

	const originalHasList = originalParagraphs.some((paragraph) =>
		paragraph.some(isSyntheticListMarkerSegment),
	);
	const originalText = originalParagraphs.map((paragraph) =>
		normalizeParagraphForMatching(
			paragraph
				.filter((segment) => !isSyntheticListMarkerSegment(segment))
				.map((segment) => segment.text)
				.join(''),
			false,
		),
	);
	const editedText = newParagraphTexts.map((text) =>
		normalizeParagraphForMatching(text, originalHasList),
	);
	const rows = originalText.length + 1;
	const columns = editedText.length + 1;
	const lcs = Array.from({ length: rows }, () => Array<number>(columns).fill(0));
	for (let oldIndex = originalText.length - 1; oldIndex >= 0; oldIndex--) {
		for (let newIndex = editedText.length - 1; newIndex >= 0; newIndex--) {
			lcs[oldIndex][newIndex] =
				originalText[oldIndex] === editedText[newIndex] && originalText[oldIndex].length > 0
					? lcs[oldIndex + 1][newIndex + 1] + 1
					: Math.max(lcs[oldIndex + 1][newIndex], lcs[oldIndex][newIndex + 1]);
		}
	}

	const mapping = Array<number | undefined>(editedText.length).fill(undefined);
	let oldIndex = 0;
	let newIndex = 0;
	while (oldIndex < originalText.length && newIndex < editedText.length) {
		if (
			originalText[oldIndex] === editedText[newIndex] &&
			originalText[oldIndex].length > 0
		) {
			mapping[newIndex] = oldIndex;
			oldIndex += 1;
			newIndex += 1;
		} else if (lcs[oldIndex + 1][newIndex] >= lcs[oldIndex][newIndex + 1]) {
			oldIndex += 1;
		} else {
			newIndex += 1;
		}
	}

	for (let index = 0; index < mapping.length; index++) {
		if (mapping[index] !== undefined) {
			continue;
		}
		let previousIndex = index - 1;
		while (previousIndex >= 0 && mapping[previousIndex] === undefined) {
			previousIndex -= 1;
		}
		let followingIndex = index + 1;
		while (followingIndex < mapping.length && mapping[followingIndex] === undefined) {
			followingIndex += 1;
		}
		mapping[index] =
			(previousIndex >= 0 ? mapping[previousIndex] : undefined) ??
			(followingIndex < mapping.length ? mapping[followingIndex] : undefined) ??
			Math.min(index, originalParagraphs.length - 1);
	}

	return mapping.map((index) => index ?? 0);
}

/** Keep per-paragraph hanging indents aligned after paragraph insertion/removal. */
export function remapParagraphIndents(
	newText: string,
	originalSegments: TextSegment[] | undefined,
	originalIndents: ParagraphIndent[] | undefined,
): ParagraphIndent[] | undefined {
	if (!originalIndents || originalIndents.length === 0) {
		return originalIndents;
	}
	if (!originalSegments || originalSegments.length === 0) {
		return newText.split('\n').map((_, index) => ({
			...(originalIndents[Math.min(index, originalIndents.length - 1)] ?? {}),
		}));
	}
	const originalParagraphs = splitOriginalParagraphs(originalSegments);
	const mapping = mapEditedParagraphsToOriginal(newText.split('\n'), originalParagraphs);
	return mapping.map((sourceIndex) => ({
		...(originalIndents[sourceIndex] ??
			originalIndents[Math.min(sourceIndex, originalIndents.length - 1)] ??
			{}),
	}));
}

/**
 * Strategy:
 * 1. Split both original segments and new text into paragraphs by "\n".
 * 2. Distribute new characters proportionally across segments.
 * 3. Extra chars go to last segment, extra paragraphs inherit last style.
 * 4. Re-insert paragraph-break markers between paragraphs.
 */
export function remapTextToSegments(
	newText: string,
	originalSegments: TextSegment[] | undefined,
	elementTextStyle: TextStyle | undefined,
): TextSegment[] {
	const fallbackStyle: TextStyle = { ...elementTextStyle };

	if (!originalSegments || originalSegments.length === 0) {
		return [{ text: newText, style: fallbackStyle }];
	}

	// Split original segments into paragraphs by paragraph-break markers.
	const originalParagraphs = splitOriginalParagraphs(originalSegments);

	const newParagraphTexts = newText.split('\n');
	const paragraphSourceIndexes = mapEditedParagraphsToOriginal(
		newParagraphTexts,
		originalParagraphs,
	);

	const firstContentSeg = originalParagraphs
		.flat()
		.find((s) => !isSyntheticListMarkerSegment(s) && s.text.trim().length > 0);
	const baseFallbackStyle: TextStyle = firstContentSeg?.style
		? { ...firstContentSeg.style }
		: fallbackStyle;
	let activeAutoNumberType: string | undefined;
	let activeAutoNumberIndex = 0;
	let activeAutoNumberStartAt = 1;

	function stripVisibleMarker(
		paragraphText: string,
		markerText: string | undefined,
		bulletInfo: TextSegment['bulletInfo'],
	): string {
		if (markerText && paragraphText.startsWith(markerText)) {
			return paragraphText.slice(markerText.length);
		}
		if (bulletInfo?.char) {
			const escaped = bulletInfo.char.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&');
			return paragraphText.replace(new RegExp(`^${escaped}\\s?`, 'u'), '');
		}
		if (bulletInfo?.autoNumType) {
			return paragraphText.replace(/^\s*\S+[.)]\s+/u, '');
		}
		if (bulletInfo?.imageDataUrl || bulletInfo?.imageRelId) {
			return paragraphText.replace(/^\u{1F4CE}\s?/u, '');
		}
		return paragraphText;
	}

	function buildStructuralMarker(
		originalMarker: TextSegment | undefined,
		paragraphBulletInfo: NonNullable<TextSegment['bulletInfo']>,
		contentStyle: TextStyle,
	): TextSegment {
		const bulletInfo = { ...paragraphBulletInfo };
		let markerText = '';
		if (bulletInfo.none) {
			activeAutoNumberType = undefined;
			activeAutoNumberIndex = 0;
			activeAutoNumberStartAt = 1;
		} else if (bulletInfo.char) {
			activeAutoNumberType = undefined;
			activeAutoNumberIndex = 0;
			activeAutoNumberStartAt = 1;
			markerText = `${bulletInfo.char} `;
		} else if (bulletInfo.autoNumType) {
			if (activeAutoNumberType === bulletInfo.autoNumType) {
				activeAutoNumberIndex += 1;
			} else {
				activeAutoNumberType = bulletInfo.autoNumType;
				activeAutoNumberIndex = 0;
				activeAutoNumberStartAt = bulletInfo.autoNumStartAt ?? 1;
			}
			bulletInfo.paragraphIndex = activeAutoNumberIndex;
			markerText = `${formatAutoNumber(
				bulletInfo.autoNumType,
				activeAutoNumberStartAt + activeAutoNumberIndex,
			)} `;
		} else if (bulletInfo.imageDataUrl || bulletInfo.imageRelId) {
			activeAutoNumberType = undefined;
			activeAutoNumberIndex = 0;
			activeAutoNumberStartAt = 1;
			markerText = '\u{1F4CE} ';
		}

		return {
			...(originalMarker ?? {}),
			text: markerText,
			style: { ...(originalMarker?.style ?? contentStyle) },
			bulletInfo,
		};
	}

	function remapParagraph(paraNewText: string, paraOrigSegments: TextSegment[]): TextSegment[] {
		if (paraOrigSegments.length === 0) {
			activeAutoNumberType = undefined;
			activeAutoNumberIndex = 0;
			activeAutoNumberStartAt = 1;
			return paraNewText.length > 0
				? [{ text: paraNewText, style: { ...baseFallbackStyle } }]
				: [{ text: '', style: { ...baseFallbackStyle } }];
		}

		const metadataSource = paraOrigSegments[0];
		const paragraphBulletInfo = metadataSource.bulletInfo;
		const originalMarker = paraOrigSegments.find(isSyntheticListMarkerSegment);
		const hasStructuralMarker = Boolean(
			paragraphBulletInfo &&
			(paragraphBulletInfo.none ||
				paragraphBulletInfo.char ||
				paragraphBulletInfo.autoNumType ||
				paragraphBulletInfo.imageDataUrl ||
				paragraphBulletInfo.imageRelId),
		);
		if (!hasStructuralMarker) {
			activeAutoNumberType = undefined;
			activeAutoNumberIndex = 0;
			activeAutoNumberStartAt = 1;
		}
		const contentOriginalSegments = paraOrigSegments
			.filter((segment) => segment !== originalMarker)
			.map((segment, index) => {
				if (index !== 0 || !hasStructuralMarker || !segment.bulletInfo) {
					return segment;
				}
				const withoutBullet = { ...segment };
				delete withoutBullet.bulletInfo;
				return withoutBullet;
			});
		const contentText = hasStructuralMarker
			? stripVisibleMarker(paraNewText, originalMarker?.text, paragraphBulletInfo)
			: paraNewText;
		const contentStyle =
			contentOriginalSegments[0]?.style ?? originalMarker?.style ?? baseFallbackStyle;

		if (contentText.length === 0) {
			const emptyStyle = { ...contentStyle };
			const result: TextSegment[] = [{ text: '', style: emptyStyle }];
			if (paragraphBulletInfo && !hasStructuralMarker) {
				result[0].bulletInfo = paragraphBulletInfo;
			}
			return hasStructuralMarker && paragraphBulletInfo
				? [buildStructuralMarker(originalMarker, paragraphBulletInfo, contentStyle), ...result]
				: result;
		}

		const totalOrigLen = contentOriginalSegments.reduce((sum, s) => sum + s.text.length, 0);

		if (totalOrigLen === 0) {
			const result: TextSegment[] = [
				copySegmentMetadata(contentOriginalSegments[0] ?? metadataSource, {
					text: contentText,
					style: { ...contentStyle },
				}),
			];
			if (paragraphBulletInfo && !hasStructuralMarker) {
				result[0].bulletInfo = paragraphBulletInfo;
			}
			if (hasStructuralMarker && paragraphBulletInfo) {
				delete result[0].paragraphLevel;
				delete result[0].endParaRunProperties;
				delete result[0].paragraphProperties;
				return [
					buildStructuralMarker(originalMarker, paragraphBulletInfo, contentStyle),
					...result,
				];
			}
			return result;
		}

		const remapped: TextSegment[] = [];
		let newPos = 0;

		for (let i = 0; i < contentOriginalSegments.length; i++) {
			const origSeg = contentOriginalSegments[i];
			const isLastSeg = i === contentOriginalSegments.length - 1;
			const origLen = origSeg.text.length;

			if (newPos >= contentText.length) {
				break;
			}

			let segText: string;
			if (isLastSeg) {
				segText = contentText.slice(newPos);
			} else {
				segText = contentText.slice(newPos, newPos + origLen);
			}

			if (segText.length > 0) {
				const outSeg: TextSegment = copySegmentMetadata(origSeg, {
					text: segText,
					style: { ...origSeg.style },
				});
				if (remapped.length === 0 && paragraphBulletInfo && !hasStructuralMarker) {
					outSeg.bulletInfo = paragraphBulletInfo;
				}
				remapped.push(outSeg);
			}

			newPos += isLastSeg ? segText.length : origLen;
		}

		if (remapped.length === 0) {
			const fallback: TextSegment = copySegmentMetadata(
				contentOriginalSegments[0] ?? metadataSource,
				{
					text: contentText,
					style: { ...contentStyle },
				},
			);
			if (paragraphBulletInfo && !hasStructuralMarker) {
				fallback.bulletInfo = paragraphBulletInfo;
			}
			return [fallback];
		}

		if (hasStructuralMarker && paragraphBulletInfo) {
			delete remapped[0].paragraphLevel;
			delete remapped[0].endParaRunProperties;
			delete remapped[0].paragraphProperties;
			return [
				buildStructuralMarker(originalMarker, paragraphBulletInfo, contentStyle),
				...remapped,
			];
		}

		if (paragraphBulletInfo && remapped[0].bulletInfo === undefined) {
			remapped[0].bulletInfo = paragraphBulletInfo;
		}

		return remapped;
	}

	const output: TextSegment[] = [];
	const lastOrigPara = originalParagraphs[originalParagraphs.length - 1];

	for (let pi = 0; pi < newParagraphTexts.length; pi++) {
		if (pi > 0) {
			const precedingOrigPara = originalParagraphs[paragraphSourceIndexes[pi - 1]] ?? [];
			const breakStyle = precedingOrigPara[0]?.style
				? { ...precedingOrigPara[0].style }
				: { ...baseFallbackStyle };
			output.push({ text: '\n', style: breakStyle, isParagraphBreak: true });
		}

		const origPara = originalParagraphs[paragraphSourceIndexes[pi]] ?? lastOrigPara ?? [];
		const paraSegments = remapParagraph(newParagraphTexts[pi], origPara);
		output.push(...paraSegments);
	}

	return output.length > 0 ? output : [{ text: '', style: { ...baseFallbackStyle } }];
}

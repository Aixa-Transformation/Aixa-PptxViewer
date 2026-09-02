import type { TextSegment } from 'pptx-viewer-core';
import { formatAutoNumber } from 'pptx-viewer-shared';

/**
 * Visible marker used by the provisional paragraph that Enter creates while a
 * list is being edited. The commit path still rebuilds structural BulletInfo;
 * this marker only keeps the live contentEditable DOM faithful until commit.
 */
export function getListContinuationMarker(
	segment: TextSegment | undefined,
	/** The number currently displayed by a provisional DOM paragraph. */
	currentAutoNumber?: number,
): string | null {
	const bulletInfo = segment?.bulletInfo;
	if (!bulletInfo || bulletInfo.none) {
		return null;
	}
	if (bulletInfo.char) {
		return `${bulletInfo.char} `;
	}
	if (bulletInfo.autoNumType) {
		const startAt = bulletInfo.autoNumStartAt ?? 1;
		const currentNumber =
			typeof currentAutoNumber === 'number' && Number.isFinite(currentAutoNumber)
				? currentAutoNumber
				: startAt + (bulletInfo.paragraphIndex ?? 0);
		return `${formatAutoNumber(bulletInfo.autoNumType, currentNumber + 1)} `;
	}
	if (bulletInfo.imageDataUrl || bulletInfo.imageRelId) {
		return '\u{1F4CE} ';
	}
	return null;
}

import type { TextSegment } from 'pptx-viewer-core';
import { formatAutoNumber } from 'pptx-viewer-shared';

/**
 * Visible marker used by the provisional paragraph that Enter creates while a
 * list is being edited. The commit path still rebuilds structural BulletInfo;
 * this marker only keeps the live contentEditable DOM faithful until commit.
 */
export function getListContinuationMarker(segment: TextSegment | undefined): string | null {
	const bulletInfo = segment?.bulletInfo;
	if (!bulletInfo || bulletInfo.none) {
		return null;
	}
	if (bulletInfo.char) {
		return `${bulletInfo.char} `;
	}
	if (bulletInfo.autoNumType) {
		const startAt = bulletInfo.autoNumStartAt ?? 1;
		const nextIndex = (bulletInfo.paragraphIndex ?? 0) + 1;
		return `${formatAutoNumber(bulletInfo.autoNumType, startAt + nextIndex)} `;
	}
	if (bulletInfo.imageDataUrl || bulletInfo.imageRelId) {
		return '\u{1F4CE} ';
	}
	return null;
}

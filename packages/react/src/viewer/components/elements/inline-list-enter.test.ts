import type { TextSegment } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { getListContinuationMarker } from './inline-list-enter';

function marker(bulletInfo: NonNullable<TextSegment['bulletInfo']>): TextSegment {
	return { text: '', style: {}, bulletInfo };
}

describe('getListContinuationMarker', () => {
	it('continues a character bullet with one separating space', () => {
		expect(getListContinuationMarker(marker({ char: '\u2022' }))).toBe('\u2022 ');
	});

	it('continues a numbered list from its structural paragraph index', () => {
		expect(
			getListContinuationMarker(
				marker({ autoNumType: 'arabicPeriod', autoNumStartAt: 1, paragraphIndex: 1 }),
			),
		).toBe('3. ');
	});

	it('continues a custom numbered start value', () => {
		expect(
			getListContinuationMarker(
				marker({ autoNumType: 'arabicPeriod', autoNumStartAt: 5, paragraphIndex: 0 }),
			),
		).toBe('6. ');
	});

	it('does not continue an explicit no-list paragraph', () => {
		expect(getListContinuationMarker(marker({ none: true }))).toBeNull();
	});
});

import { describe, expect, it } from 'vitest';

import type { PptxElement } from '../types';
import {
	createBackgroundPreservedArtwork,
	createBackgroundOverridePlaceholder,
	isBackgroundPreservedArtwork,
	isBackgroundOverridePlaceholder,
	isFullSlidePicturePlaceholder,
} from './background-override-placeholder';

function picturePlaceholder(overrides: Partial<PptxElement> = {}): PptxElement {
	return {
		id: 'layout-placeholder-picture',
		type: 'shape',
		x: 0,
		y: 0,
		width: 1280,
		height: 720,
		text: '',
		promptText: 'Click icon to add picture',
		shapeStyle: { fillColor: '#DDE3E7' },
		rawXml: {
			'p:nvSpPr': {
				'p:cNvPr': { '@_id': '13', '@_name': 'Picture Placeholder 8' },
				'p:nvPr': { 'p:ph': { '@_type': 'pic', '@_idx': '13' } },
			},
			'p:spPr': { 'a:solidFill': { 'a:srgbClr': { '@_val': 'DDE3E7' } } },
			'p:txBody': { 'a:bodyPr': {}, 'a:lstStyle': {}, 'a:p': { 'a:r': { 'a:t': 'Prompt' } } },
		},
		...overrides,
	} as PptxElement;
}

describe('background override placeholder', () => {
	it('identifies full-slide picture placeholders and ordinary layout pictures', () => {
		expect(isFullSlidePicturePlaceholder(picturePlaceholder())).toBe(true);
		expect(
			isFullSlidePicturePlaceholder(
				picturePlaceholder({
					id: 'layout-background-picture',
					type: 'picture',
					rawXml: {
						'p:nvPicPr': {
							'p:cNvPr': { '@_id': '13', '@_name': 'Picture 5' },
							'p:nvPr': {},
						},
					},
				}),
			),
		).toBe(true);
		expect(
			isFullSlidePicturePlaceholder(
				picturePlaceholder({ x: 700, y: 40, width: 500, height: 600 }),
			),
		).toBe(false);
	});

	it('creates a transparent, empty slide-owned placeholder binding', () => {
		const override = createBackgroundOverridePlaceholder(picturePlaceholder());
		expect(override).not.toBeNull();
		expect(override?.id.startsWith('aixa-background-override-')).toBe(true);
		expect(isBackgroundOverridePlaceholder(override!)).toBe(true);
		expect(override?.text).toBe('');
		expect(override?.promptText).toBeUndefined();
		expect(override?.shapeStyle).toMatchObject({
			fillMode: 'none',
			fillColor: 'transparent',
			strokeColor: 'transparent',
			strokeWidth: 0,
		});
	});

	it('creates ordinary slide-owned artwork without a placeholder binding', () => {
		const preserved = createBackgroundPreservedArtwork(
			picturePlaceholder({ id: 'layout-logo', name: 'Swisscom', type: 'picture' }),
		);
		expect(isBackgroundPreservedArtwork(preserved)).toBe(true);
		expect(preserved.type).toBe('shape');
		expect(preserved.shapeId).toBeUndefined();
		expect(
			(preserved.rawXml?.['p:nvSpPr'] as Record<string, unknown>)?.['p:nvPr'],
		).not.toHaveProperty('p:ph');
	});
});

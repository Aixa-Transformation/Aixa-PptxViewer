import { describe, expect, it } from 'vitest';

import type { XmlObject } from '../../types';
import { shouldRenderLayoutPlaceholderArtwork } from './layout-placeholder-artwork';

function placeholderShape(type: string, spPr: XmlObject = {}): XmlObject {
	return {
		'p:nvSpPr': {
			'p:nvPr': {
				'p:ph': { '@_type': type },
			},
		},
		'p:spPr': spPr,
	};
}

describe('layout placeholder artwork detection', () => {
	it('keeps picture placeholders even when their authored geometry has no fill', () => {
		expect(
			shouldRenderLayoutPlaceholderArtwork(
				placeholderShape('pic', {
					'a:custGeom': {},
					'a:noFill': {},
				}),
			),
		).toBe(true);
	});

	it('keeps image-filled placeholders used as layout logos', () => {
		expect(
			shouldRenderLayoutPlaceholderArtwork(
				placeholderShape('body', {
					'a:blipFill': { 'a:blip': {} },
				}),
			),
		).toBe(true);
	});

	it('keeps solid placeholder surfaces used as picture backings', () => {
		expect(
			shouldRenderLayoutPlaceholderArtwork(
				placeholderShape('pic', {
					'a:solidFill': { 'a:srgbClr': { '@_val': 'DDE3E7' } },
				}),
			),
		).toBe(true);
	});

	it('still excludes ordinary text placeholders from the inherited layer', () => {
		expect(shouldRenderLayoutPlaceholderArtwork(placeholderShape('title'))).toBe(false);
		expect(
			shouldRenderLayoutPlaceholderArtwork(
				placeholderShape('body', {
					'a:noFill': {},
					'a:ln': { 'a:noFill': {} },
				}),
			),
		).toBe(false);
	});
});

import { describe, expect, it } from 'vitest';

import { getDeckBackgroundPatch, getMasterBackgroundTargetPaths } from './SlideBackgroundPanel';

describe('getMasterBackgroundTargetPaths', () => {
	const masters = [
		{ path: 'ppt/slideMasters/slideMaster1.xml' },
		{ path: 'ppt/slideMasters/slideMaster2.xml' },
	];

	it('targets only the first master when requested', () => {
		expect(getMasterBackgroundTargetPaths(masters, false)).toEqual([
			'ppt/slideMasters/slideMaster1.xml',
		]);
	});

	it('targets every master when requested', () => {
		expect(getMasterBackgroundTargetPaths(masters, true)).toEqual([
			'ppt/slideMasters/slideMaster1.xml',
			'ppt/slideMasters/slideMaster2.xml',
		]);
	});

	it('returns no targets when a deck has no parsed masters', () => {
		expect(getMasterBackgroundTargetPaths(undefined, true)).toEqual([]);
	});
});

describe('getDeckBackgroundPatch', () => {
	it('copies an uploaded image and its fallback colour to the whole deck', () => {
		expect(
			getDeckBackgroundPatch({
				id: 'slide-1',
				elements: [],
				backgroundColor: '#102030',
				backgroundImage: 'data:image/png;base64,AAAA',
			}),
		).toEqual({
			backgroundColor: '#102030',
			backgroundImage: 'data:image/png;base64,AAAA',
			backgroundGradient: undefined,
		});
	});

	it('marks an old image for removal when applying a solid colour', () => {
		expect(
			getDeckBackgroundPatch({
				id: 'slide-1',
				elements: [],
				backgroundColor: '#AABBCC',
			}),
		).toEqual({
			backgroundColor: '#AABBCC',
			backgroundImage: '',
			backgroundGradient: undefined,
		});
	});
});

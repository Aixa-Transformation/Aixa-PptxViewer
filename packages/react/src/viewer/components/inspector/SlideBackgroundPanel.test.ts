import { describe, expect, it } from 'vitest';

import { getDeckBackgroundPatch } from './SlideBackgroundPanel';

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
			backgroundSource: 'slide',
			showMasterShapes: true,
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
			backgroundSource: 'slide',
			showMasterShapes: true,
		});
	});

	it('supports clearing a background and copying background-graphics visibility', () => {
		expect(getDeckBackgroundPatch({ id: 's', elements: [], showMasterShapes: false })).toEqual({
			backgroundColor: undefined,
			backgroundImage: '',
			backgroundGradient: undefined,
			backgroundSource: 'inherited',
			showMasterShapes: false,
		});
	});

	it('retains an explicit hide-background-graphics choice with a slide background', () => {
		expect(
			getDeckBackgroundPatch({
				id: 's',
				elements: [],
				backgroundColor: '#AABBCC',
				backgroundSource: 'slide',
				showMasterShapes: false,
			}),
		).toMatchObject({ backgroundSource: 'slide', showMasterShapes: false });
	});
});

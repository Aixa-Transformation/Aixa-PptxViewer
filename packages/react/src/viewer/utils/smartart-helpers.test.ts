import { describe, expect, it } from 'vitest';

import { contrastingTextColor } from './smartart-helpers';

describe('contrastingTextColor', () => {
	it('uses dark text on white and light SmartArt panels', () => {
		expect(contrastingTextColor('#FFFFFF')).toBe('#000000');
		expect(contrastingTextColor('#D8B11D')).toBe('#000000');
	});

	it('uses light text on dark SmartArt heading shapes', () => {
		expect(contrastingTextColor('#2C2EA2')).toBe('#FFFFFF');
		expect(contrastingTextColor('#196B24')).toBe('#FFFFFF');
	});

	it('handles shorthand hex and unknown fills safely', () => {
		expect(contrastingTextColor('#fff')).toBe('#000000');
		expect(contrastingTextColor(undefined)).toBe('#000000');
		expect(contrastingTextColor('url(#gradient)')).toBe('#000000');
	});
});

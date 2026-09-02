import { describe, it, expect } from 'vitest';

import { getRenderedAutoFitScale, isTextAutoFitEnabled } from './text-autofit-dom';

describe('isTextAutoFitEnabled', () => {
	it('enables shrink-to-fit for normAutofit and spAutoFit bodies', () => {
		expect(isTextAutoFitEnabled({ autoFit: true })).toBe(true);
		expect(isTextAutoFitEnabled({ autoFitMode: 'normal' })).toBe(true);
		expect(isTextAutoFitEnabled({ autoFitMode: 'shrink' })).toBe(true);
	});

	it('leaves noAutofit and unspecified bodies at their authored size', () => {
		expect(isTextAutoFitEnabled({ autoFitMode: 'none' })).toBe(false);
		expect(isTextAutoFitEnabled({ autoFit: true, autoFitMode: 'none' })).toBe(false);
		expect(isTextAutoFitEnabled({})).toBe(false);
		expect(isTextAutoFitEnabled(undefined)).toBe(false);
	});
});

describe('getRenderedAutoFitScale', () => {
	it('returns the scale already baked into the rendered markup', () => {
		expect(getRenderedAutoFitScale({ autoFitFontScale: 0.75 })).toBe(0.75);
	});

	it('falls back to 1 for missing or out-of-range scales', () => {
		expect(getRenderedAutoFitScale(undefined)).toBe(1);
		expect(getRenderedAutoFitScale({})).toBe(1);
		expect(getRenderedAutoFitScale({ autoFitFontScale: 0 })).toBe(1);
		expect(getRenderedAutoFitScale({ autoFitFontScale: 1 })).toBe(1);
	});
});

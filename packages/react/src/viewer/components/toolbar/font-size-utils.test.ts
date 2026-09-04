import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import {
	getAuthoredFontSizeForToolbar,
	getEffectiveToolbarFontSize,
	getToolbarAutoFitScale,
} from './font-size-utils';

const autoFitElement = {
	id: 'shape-1',
	type: 'shape',
	textStyle: { fontSize: 24, autoFitFontScale: 0.75 },
} as PptxElement;

describe('toolbar font size autofit synchronization', () => {
	it('shows the effective on-slide font size', () => {
		expect(getToolbarAutoFitScale(autoFitElement)).toBe(0.75);
		expect(getEffectiveToolbarFontSize(autoFitElement, 24)).toBe(18);
		expect(getEffectiveToolbarFontSize(autoFitElement, 58.67)).toBe(44);
	});

	it('converts a toolbar choice back to the authored size without changing the visual result', () => {
		const authored = getAuthoredFontSizeForToolbar(autoFitElement, 20);
		expect(authored).toBeCloseTo(26.67, 2);
		expect(getEffectiveToolbarFontSize(autoFitElement, authored)).toBe(20);
	});

	it('prefers the live editing scale before the edit is committed', () => {
		expect(getToolbarAutoFitScale(autoFitElement, 0.5)).toBe(0.5);
		expect(getEffectiveToolbarFontSize(autoFitElement, 24, 0.5)).toBe(12);
		const authored = getAuthoredFontSizeForToolbar(autoFitElement, 18, 0.5);
		expect(authored).toBe(36);
	});

	it('never exposes a decimal font size in the toolbar', () => {
		for (const authored of [10.25, 19.99, 26.67, 44.5, 58.67]) {
			const displayed = getEffectiveToolbarFontSize(autoFitElement, authored, 0.73);
			expect(Number.isInteger(displayed)).toBe(true);
		}
	});
});

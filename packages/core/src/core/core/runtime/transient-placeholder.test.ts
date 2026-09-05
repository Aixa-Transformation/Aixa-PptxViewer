import { describe, expect, it } from 'vitest';
import type { PptxElement } from '../../types';
import {
	isEmptyGeneratedPlaceholder,
	markGeneratedPlaceholder,
	persistGeneratedPlaceholder,
	shouldPersistGeneratedPlaceholder,
} from './transient-placeholder';

function emptyTextElement(): PptxElement {
	return {
		type: 'text',
		id: 'placeholder',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		text: '',
	};
}

describe('generated layout placeholder persistence', () => {
	it('keeps load-only empty prompt placeholders transient', () => {
		const element = markGeneratedPlaceholder(emptyTextElement());
		expect(isEmptyGeneratedPlaceholder(element)).toBe(true);
		expect(shouldPersistGeneratedPlaceholder(element)).toBe(false);
	});

	it('persists empty bindings created by an explicit layout switch', () => {
		const element = persistGeneratedPlaceholder(emptyTextElement());
		expect(isEmptyGeneratedPlaceholder(element)).toBe(true);
		expect(shouldPersistGeneratedPlaceholder(element)).toBe(true);
	});

	it('does not consider a generated placeholder empty after text is entered', () => {
		const element = persistGeneratedPlaceholder(emptyTextElement());
		(element as PptxElement & { text?: string }).text = 'Content';
		expect(isEmptyGeneratedPlaceholder(element)).toBe(false);
	});
});

// @vitest-environment happy-dom

import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';

import type { ElementOperations, UseElementOperationsInput } from './useElementOperations';
import { useElementOperations } from './useElementOperations';

describe('useElementOperations live list editing', () => {
	it('receives the live inline text through its input instead of reading a free variable', () => {
		const selectedElement = {
			id: 'shape-1',
			type: 'shape',
			x: 0,
			y: 0,
			width: 200,
			height: 100,
			rotation: 0,
			flipHorizontal: false,
			flipVertical: false,
			hidden: false,
			opacity: 1,
			rawXml: {},
			text: 'Existing paragraph',
			textSegments: [{ text: 'Existing paragraph', style: {} }],
			textStyle: { fontSize: 20 },
		} as PptxElement;
		const activeSlide = {
			id: 'slide-1',
			rId: '',
			slideNumber: 1,
			elements: [selectedElement],
		} as PptxSlide;
		const setInlineEditingText = vi.fn();
		const input: UseElementOperationsInput = {
			slides: [activeSlide],
			activeSlide,
			activeSlideIndex: 0,
			selectedElement,
			selectedElementId: selectedElement.id,
			editTemplateMode: false,
			templateElements: [],
			history: { markDirty: vi.fn() } as UseElementOperationsInput['history'],
			setSlides: vi.fn(),
			setTemplateElementsBySlideId: vi.fn(),
			setSelectedElementId: vi.fn(),
			setSelectedElementIds: vi.fn(),
			setInlineEditingElementId: vi.fn(),
			inlineEditingElementId: selectedElement.id,
			inlineEditingText: 'Existing paragraph\nNewly typed paragraph',
			setInlineEditingText,
			setContextMenuState: vi.fn(),
		};
		let operations: ElementOperations | undefined;
		const Harness = (): null => {
			operations = useElementOperations(input);
			return null;
		};
		const container = document.createElement('div');
		document.body.appendChild(container);
		const root = createRoot(container);

		act(() => root.render(createElement(Harness)));
		expect(() => {
			act(() => operations?.updateSelectedTextStyle({ listType: 'numbered' }));
		}).not.toThrow();
		expect(setInlineEditingText).toHaveBeenCalledWith(
			expect.stringContaining('Newly typed paragraph'),
		);

		act(() => root.unmount());
		container.remove();
	});
});

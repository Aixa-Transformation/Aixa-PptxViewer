// @vitest-environment happy-dom
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { translationsEn } from 'pptx-viewer-shared/i18n';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { SHAPE_QUICK_STYLES } from '../../constants';

vi.mock(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
	}),
}));

const { DrawingGroup } = await import('./DrawingGroup');
type DrawingGroupProps = import('./DrawingGroup').DrawingGroupProps;

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

function createProps(
	onUpdateElementStyle: (style: Partial<ShapeStyle>) => void,
): DrawingGroupProps {
	return {
		canEdit: true,
		selectedElement: { id: 'shape-1', type: 'shape' } as PptxElement,
		newShapeType: 'rect',
		onSetNewShapeType: vi.fn(),
		onAddShape: vi.fn(),
		onMoveLayer: vi.fn(),
		onMoveLayerToEdge: vi.fn(),
		onUpdateElementStyle,
	};
}

function renderGroup(onUpdateElementStyle: (style: Partial<ShapeStyle>) => void): void {
	act(() => {
		root.render(<DrawingGroup {...createProps(onUpdateElementStyle)} />);
	});
}

function clickButton(selector: string): void {
	const button = container.querySelector<HTMLButtonElement>(selector);
	expect(button).not.toBeNull();
	act(() => button?.click());
}

describe('DrawingGroup shape styling', () => {
	it('applies a solid shape fill using schema-native style fields', () => {
		const onUpdateElementStyle = vi.fn<(style: Partial<ShapeStyle>) => void>();
		renderGroup(onUpdateElementStyle);

		clickButton('button[title="Shape Fill"]');
		clickButton('button[aria-label="Fill colour #ff0000"]');

		expect(onUpdateElementStyle).toHaveBeenCalledOnce();
		expect(onUpdateElementStyle).toHaveBeenCalledWith({
			fillMode: 'solid',
			fillColor: '#ff0000',
		});
	});

	it('applies a solid shape outline using schema-native style fields', () => {
		const onUpdateElementStyle = vi.fn<(style: Partial<ShapeStyle>) => void>();
		renderGroup(onUpdateElementStyle);

		clickButton('button[title="Shape Outline"]');
		clickButton('button[aria-label="Outline colour #0000ff"]');

		expect(onUpdateElementStyle).toHaveBeenCalledOnce();
		expect(onUpdateElementStyle).toHaveBeenCalledWith({
			strokeFillMode: 'solid',
			strokeColor: '#0000ff',
		});
	});

	it('applies the selected shared quick-style preset', () => {
		const onUpdateElementStyle = vi.fn<(style: Partial<ShapeStyle>) => void>();
		const quickStyle = SHAPE_QUICK_STYLES[0];
		renderGroup(onUpdateElementStyle);

		clickButton('button[title="Quick Styles"]');
		clickButton(`button[aria-label="${quickStyle.name}"]`);

		expect(onUpdateElementStyle).toHaveBeenCalledOnce();
		expect(onUpdateElementStyle).toHaveBeenCalledWith(quickStyle.style);
	});
});

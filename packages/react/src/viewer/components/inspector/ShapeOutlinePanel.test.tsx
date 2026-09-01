// @vitest-environment happy-dom
import type { PptxTheme, ShapeStyle } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import {
	createNoOutlinePatch,
	createSolidOutlinePatch,
	outlinePixelsToPoints,
	outlinePointsToPixels,
	ShapeOutlinePanel,
} from './ShapeOutlinePanel';

vi.mock(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string, fallback?: string) => fallback ?? key,
	}),
}));

const theme: PptxTheme = {
	colorScheme: {
		dk1: '#000000',
		lt1: '#FFFFFF',
		dk2: '#44546A',
		lt2: '#E7E6E6',
		accent1: '#4472C4',
		accent2: '#ED7D31',
		accent3: '#A5A5A5',
		accent4: '#FFC000',
		accent5: '#5B9BD5',
		accent6: '#70AD47',
		hlink: '#0563C1',
		folHlink: '#954F72',
	},
};

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

function renderPanel(
	style: ShapeStyle,
	onUpdateShapeStyle: (patch: Partial<ShapeStyle>) => void,
	showArrows = false,
): void {
	act(() => {
		root.render(
			<ShapeOutlinePanel
				style={style}
				theme={theme}
				canEdit
				showArrows={showArrows}
				onUpdateShapeStyle={onUpdateShapeStyle}
			/>,
		);
	});
}

function changeSelect(label: string, value: string): void {
	const select = container.querySelector<HTMLSelectElement>(`select[aria-label="${label}"]`);
	expect(select).not.toBeNull();
	act(() => {
		select!.value = value;
		select!.dispatchEvent(new Event('change', { bubbles: true }));
	});
}

describe('ShapeOutlinePanel', () => {
	it('converts PowerPoint point weights to the internal pixel unit', () => {
		expect(outlinePointsToPixels(3)).toBe(4);
		expect(outlinePixelsToPoints(4)).toBe(3);
	});

	it('creates schema-native solid and no-outline patches', () => {
		expect(createSolidOutlinePatch('#ff0000', { strokeWidth: 2 })).toStrictEqual({
			strokeFillMode: 'solid',
			strokeColor: '#ff0000',
			strokeWidth: 2,
		});
		expect(createSolidOutlinePatch('#ff0000', { strokeWidth: 0 }).strokeWidth).toBeCloseTo(
			4 / 3,
		);
		expect(createNoOutlinePatch()).toStrictEqual({
			strokeFillMode: 'none',
			strokeColor: 'transparent',
			strokeWidth: 0,
		});
	});

	it('applies standard colours, weight, dashes, and no outline', () => {
		const update = vi.fn<(patch: Partial<ShapeStyle>) => void>();
		renderPanel({ strokeColor: '#000000', strokeWidth: 4 / 3 }, update);

		act(() => {
			container
				.querySelector<HTMLButtonElement>('button[aria-label="Outline colour #FF0000"]')!
				.click();
		});
		expect(update).toHaveBeenLastCalledWith({
			strokeFillMode: 'solid',
			strokeColor: '#FF0000',
			strokeWidth: 4 / 3,
		});

		changeSelect('Outline weight', '3');
		expect(update).toHaveBeenLastCalledWith({ strokeWidth: 4 });

		changeSelect('Outline dashes', 'dashDot');
		expect(update).toHaveBeenLastCalledWith({ strokeDash: 'dashDot' });

		const noOutline = [...container.querySelectorAll<HTMLButtonElement>('button')].find((button) =>
			button.textContent?.includes('No Outline'),
		);
		act(() => noOutline!.click());
		expect(update).toHaveBeenLastCalledWith(createNoOutlinePatch());
	});

	it('shows arrow controls only for lines and updates schema arrow fields', () => {
		const update = vi.fn<(patch: Partial<ShapeStyle>) => void>();
		renderPanel({ strokeColor: '#000000', strokeWidth: 1 }, update, true);

		const selects = container.querySelectorAll<HTMLSelectElement>('select');
		const startArrow = [...selects].find((select) =>
			select.parentElement?.textContent?.includes('pptx.connectorArrows.startArrow'),
		);
		expect(startArrow).toBeDefined();
		act(() => {
			startArrow!.value = 'triangle';
			startArrow!.dispatchEvent(new Event('change', { bubbles: true }));
		});
		expect(update).toHaveBeenLastCalledWith({ connectorStartArrow: 'triangle' });
	});
});

/**
 * Tests for pure logic and type contract of useThemeHandlers (GAP-E3).
 *
 * The theme switching coordination logic delegates to PptxHandler.
 * These tests verify the handler result shape and edge-case handling
 * without mounting React components.
 */
import { describe, it, expect, vi } from 'vitest';

import type { PptxElement, PptxThemeColorScheme } from 'pptx-viewer-core';

import {
	reResolveTemplateElementsBySlideId,
	type ThemeHandlersResult,
	type UseThemeHandlersInput,
} from './useThemeHandlers';

// ---------------------------------------------------------------------------
// Type-level assertions: ensure the new methods exist in the result type
// ---------------------------------------------------------------------------

describe('themeHandlersResult type contract', () => {
	it('should include handleGetAvailableThemes in the result interface', () => {
		// Compile-time check: the property must exist on ThemeHandlersResult
		const check: keyof ThemeHandlersResult = 'handleGetAvailableThemes';
		expect(check).toBe('handleGetAvailableThemes');
	});

	it('should include handleSwitchTheme in the result interface', () => {
		const check: keyof ThemeHandlersResult = 'handleSwitchTheme';
		expect(check).toBe('handleSwitchTheme');
	});

	it('should include all original handlers in the result interface', () => {
		const keys: Array<keyof ThemeHandlersResult> = [
			'handleApplyTheme',
			'handleUpdateThemeColorScheme',
			'handleUpdateThemeFontScheme',
			'handleUpdateThemeName',
			'handleApplyThemeToPresentation',
			'handleApplyThemeData',
			'handleSetTemplateBackground',
			'handleGetTemplateBackgroundColor',
			'handleGetAvailableThemes',
			'handleSwitchTheme',
		];
		expect(keys).toHaveLength(10);
	});
});

// ---------------------------------------------------------------------------
// Input type validation
// ---------------------------------------------------------------------------

describe('useThemeHandlersInput contract', () => {
	it('should accept all required input fields', () => {
		const input: UseThemeHandlersInput = {
			handlerRef: { current: null },
			serializeSlides: vi.fn<() => void>().mockResolvedValue(null),
			setContent: vi.fn<() => void>(),
			onContentChange: undefined,
			setTheme: vi.fn<() => void>(),
			setSlideMasters: vi.fn<() => void>(),
			slideMasters: [],
			history: {
				markDirty: vi.fn<() => void>(),
			} as unknown as UseThemeHandlersInput['history'],
			setSlides: vi.fn<() => void>(),
			setTemplateElementsBySlideId: vi.fn<() => void>(),
			theme: undefined,
			bumpHistory: vi.fn<() => void>(),
		};

		expect(input.handlerRef.current).toBeNull();
		expect(input.slideMasters).toStrictEqual([]);
	});
});

describe('reResolveTemplateElementsBySlideId', () => {
	it('re-colours master/layout elements stored outside slide.elements', () => {
		const elements = {
			'slide-1': [
				{
					id: 'master-background',
					type: 'shape',
					shapeType: 'rect',
					x: 0,
					y: 0,
					width: 1280,
					height: 720,
					shapeStyle: { fillColor: '#94B6D2' },
				} as PptxElement,
			],
		};
		const scheme = {
			dk1: '#000000',
			lt1: '#FFFFFF',
			dk2: '#222222',
			lt2: '#EEEEEE',
			accent1: '#4472C4',
			accent2: '#ED7D31',
			accent3: '#A5A5A5',
			accent4: '#FFC000',
			accent5: '#5B9BD5',
			accent6: '#70AD47',
			hlink: '#0563C1',
			folHlink: '#954F72',
		} satisfies PptxThemeColorScheme;

		const result = reResolveTemplateElementsBySlideId(
			elements,
			{ accent1: '94B6D2' },
			scheme,
		);

		expect(result['slide-1'][0].shapeStyle?.fillColor).toBe('#4472C4');
	});
});

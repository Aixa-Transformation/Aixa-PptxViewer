import { describe, it, expect } from 'vitest';

import {
	MOBILE_BREAKPOINT,
	TABLET_BREAKPOINT,
	MIN_TOUCH_TARGET,
	DESKTOP_ONLY_VIEWER_LAYOUT,
	deriveViewerBreakpoints,
	isMobileRuntimePlatform,
} from './useIsMobile';

// ---------------------------------------------------------------------------
// Since useIsMobile is a React hook, we test the pure helper constants and
// the breakpoint logic by extracting the derivation into plain functions.
// For the hook itself we verify the exported constants.
// ---------------------------------------------------------------------------

describe('useIsMobile constants', () => {
	it('mOBILE_BREAKPOINT is 768', () => {
		expect(MOBILE_BREAKPOINT).toBe(768);
	});

	it('tABLET_BREAKPOINT is 1024', () => {
		expect(TABLET_BREAKPOINT).toBe(1024);
	});

	it('mIN_TOUCH_TARGET is 44px per WCAG guidelines', () => {
		expect(MIN_TOUCH_TARGET).toBe(44);
	});

	it('keeps the full desktop editor layout enabled on every device', () => {
		expect(DESKTOP_ONLY_VIEWER_LAYOUT).toBe(true);
	});
});

describe('breakpoint derivation logic', () => {
	const deriveBreakpoints = (viewportWidth: number) =>
		deriveViewerBreakpoints(viewportWidth, 800, false);

	it('classifies 320px as mobile', () => {
		const result = deriveBreakpoints(320);
		expect(result).toStrictEqual({
			isMobile: true,
			isTablet: false,
			isDesktop: false,
		});
	});

	it('classifies 375px (iPhone) as mobile', () => {
		const result = deriveBreakpoints(375);
		expect(result).toStrictEqual({
			isMobile: true,
			isTablet: false,
			isDesktop: false,
		});
	});

	it('classifies 767px as mobile (just below breakpoint)', () => {
		const result = deriveBreakpoints(767);
		expect(result).toStrictEqual({
			isMobile: true,
			isTablet: false,
			isDesktop: false,
		});
	});

	it('classifies 768px as tablet (exactly at mobile breakpoint)', () => {
		const result = deriveBreakpoints(768);
		expect(result).toStrictEqual({
			isMobile: false,
			isTablet: true,
			isDesktop: false,
		});
	});

	it('classifies 900px as tablet', () => {
		const result = deriveBreakpoints(900);
		expect(result).toStrictEqual({
			isMobile: false,
			isTablet: true,
			isDesktop: false,
		});
	});

	it('classifies 1023px as tablet (just below desktop breakpoint)', () => {
		const result = deriveBreakpoints(1023);
		expect(result).toStrictEqual({
			isMobile: false,
			isTablet: true,
			isDesktop: false,
		});
	});

	it('classifies 1024px as desktop (exactly at tablet breakpoint)', () => {
		const result = deriveBreakpoints(1024);
		expect(result).toStrictEqual({
			isMobile: false,
			isTablet: false,
			isDesktop: true,
		});
	});

	it('classifies 1920px as desktop', () => {
		const result = deriveBreakpoints(1920);
		expect(result).toStrictEqual({
			isMobile: false,
			isTablet: false,
			isDesktop: true,
		});
	});

	it('classifies 0px as mobile', () => {
		const result = deriveBreakpoints(0);
		expect(result).toStrictEqual({
			isMobile: true,
			isTablet: false,
			isDesktop: false,
		});
	});

	it('exactly one flag is true for any width', () => {
		// Test a broad range of widths to ensure mutual exclusivity
		for (const width of [0, 100, 320, 767, 768, 900, 1023, 1024, 2560]) {
			const result = deriveBreakpoints(width);
			const trueCount = [result.isMobile, result.isTablet, result.isDesktop].filter(Boolean).length;
			expect(trueCount).toBe(1);
		}
	});

	it('keeps desktop chrome when an embedded host container becomes narrow', () => {
		// Breakpoints intentionally use the 1440px browser viewport. A host
		// sidebar may leave only 640px for the editor, but that container resize
		// must not replace the desktop ribbon with the mobile toolbar.
		const narrowContainerWidth = 640;
		expect(narrowContainerWidth).toBeLessThan(MOBILE_BREAKPOINT);
		expect(deriveViewerBreakpoints(1440, 900, false)).toStrictEqual({
			isMobile: false,
			isTablet: false,
			isDesktop: true,
		});
	});

	it('still selects mobile chrome for a real phone viewport', () => {
		expect(deriveViewerBreakpoints(390, 844, true, true)).toStrictEqual({
			isMobile: true,
			isTablet: false,
			isDesktop: false,
		});
	});

	it('keeps desktop chrome in a narrow Windows desktop browser', () => {
		expect(deriveViewerBreakpoints(700, 800, true, false)).toStrictEqual({
			isMobile: false,
			isTablet: false,
			isDesktop: true,
		});
	});
});

describe('mobile runtime platform detection', () => {
	it('does not treat a Windows touch desktop as mobile', () => {
		expect(
			isMobileRuntimePlatform(
				'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
				'Win32',
				10,
			),
		).toBeFalsy();
	});

	it('recognizes Android, iPhone, and modern iPadOS runtimes', () => {
		expect(isMobileRuntimePlatform('Mozilla/5.0 (Linux; Android 15)', 'Linux armv8l', 5)).toBeTruthy();
		expect(isMobileRuntimePlatform('Mozilla/5.0 (iPhone)', 'iPhone', 5)).toBeTruthy();
		expect(isMobileRuntimePlatform('Mozilla/5.0 (Macintosh)', 'MacIntel', 5)).toBeTruthy();
	});

	it('prefers the user-agent client hint when it is available', () => {
		expect(isMobileRuntimePlatform('Desktop UA', 'Win32', 0, true)).toBeTruthy();
		expect(isMobileRuntimePlatform('Android', 'Linux armv8l', 5, false)).toBeFalsy();
	});
});

describe('virtual keyboard detection logic', () => {
	function isVirtualKeyboardOpen(initialHeight: number, currentHeight: number): boolean {
		const shrinkRatio = currentHeight / initialHeight;
		return shrinkRatio < 0.7;
	}

	it('detects keyboard open when viewport shrinks by more than 30%', () => {
		expect(isVirtualKeyboardOpen(800, 400)).toBeTruthy();
		expect(isVirtualKeyboardOpen(800, 300)).toBeTruthy();
	});

	it('does not detect keyboard when viewport barely shrinks', () => {
		expect(isVirtualKeyboardOpen(800, 700)).toBeFalsy();
		expect(isVirtualKeyboardOpen(800, 800)).toBeFalsy();
	});

	it('edge case: exactly 70% ratio is not open', () => {
		// 560 / 800 = 0.7, not less than 0.7
		expect(isVirtualKeyboardOpen(800, 560)).toBeFalsy();
	});

	it('edge case: just below 70% ratio is open', () => {
		expect(isVirtualKeyboardOpen(800, 559)).toBeTruthy();
	});
});

describe('orientation derivation logic', () => {
	function deriveOrientation(width: number, height: number): 'portrait' | 'landscape' {
		return height > width ? 'portrait' : 'landscape';
	}

	it('tall viewport is portrait', () => {
		expect(deriveOrientation(375, 812)).toBe('portrait');
	});

	it('wide viewport is landscape', () => {
		expect(deriveOrientation(812, 375)).toBe('landscape');
	});

	it('square viewport is landscape', () => {
		expect(deriveOrientation(500, 500)).toBe('landscape');
	});
});

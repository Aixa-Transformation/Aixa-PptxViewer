import type { PptxElement, PptxSlide, PptxLayoutOption, PptxHandler } from 'pptx-viewer-core';
/**
 * useLayoutSwitching -- Hook for switching an existing slide's layout.
 *
 * Wraps the core `getAvailableLayoutsForSlide` and `applyLayoutToSlide`
 * APIs and exposes them as React-friendly callbacks with loading state.
 */
import { useState, useCallback, useRef } from 'react';
import type React from 'react';

import type { EditorHistoryResult } from './useEditorHistory';
import type { ElementOperations } from './useElementOperations';

/**
 * Input for {@link useLayoutSwitching}.
 */
export interface UseLayoutSwitchingInput {
	/** Current PPTX handler instance (may be null before load). */
	handler: PptxHandler | null;
	/** Current slides array. */
	slides: PptxSlide[];
	/** Index of the currently active slide. */
	activeSlideIndex: number;
	/** Element operations for updating the slides array. */
	ops: ElementOperations;
	/** Editor history for marking dirty state. */
	history: EditorHistoryResult;
	/** Template elements shown behind each slide; refreshed after applying a layout. */
	setTemplateElementsBySlideId: React.Dispatch<
		React.SetStateAction<Record<string, PptxElement[]>>
	>;
}

/**
 * Result returned by {@link useLayoutSwitching}.
 */
export interface LayoutSwitchingResult {
	/** Available layouts for the active slide (populated after calling `loadAvailableLayouts`). */
	availableLayouts: PptxLayoutOption[];
	/** Whether a layout operation is in progress. */
	isLoading: boolean;
	/** Load the available layouts for the current slide. */
	loadAvailableLayouts: () => Promise<void>;
	/** Apply a layout to the active slide by its archive path. */
	applyLayout: (layoutPath: string) => Promise<void>;
	/** The current slide's layout path (if known). */
	currentLayoutPath: string | undefined;
}

/**
 * Hook that provides layout switching capabilities for the editor.
 *
 * @example
 * ```tsx
 * const { availableLayouts, applyLayout, loadAvailableLayouts } =
 *   useLayoutSwitching({ handler, slides, activeSlideIndex, ops, history });
 *
 * // Load layouts when the user opens a layout picker
 * await loadAvailableLayouts();
 *
 * // Switch to a different layout
 * await applyLayout("ppt/slideLayouts/slideLayout3.xml");
 * ```
 */
export function useLayoutSwitching(input: UseLayoutSwitchingInput): LayoutSwitchingResult {
	const { handler, slides, activeSlideIndex, ops, history, setTemplateElementsBySlideId } = input;

	const [availableLayouts, setAvailableLayouts] = useState<PptxLayoutOption[]>([]);
	const [isLoading, setIsLoading] = useState(false);

	// Keep a ref to the latest slides so callbacks don't go stale
	const slidesRef = useRef(slides);
	slidesRef.current = slides;

	const currentLayoutPath = slides[activeSlideIndex]?.layoutPath;

	const loadAvailableLayouts = useCallback(async () => {
		if (!handler) {
			return;
		}
		setIsLoading(true);
		try {
			const layouts = await handler.getAvailableLayoutsForSlide(
				activeSlideIndex,
				slidesRef.current,
			);
			setAvailableLayouts(layouts);
		} finally {
			setIsLoading(false);
		}
	}, [handler, activeSlideIndex]);

	const applyLayout = useCallback(
		async (layoutPath: string) => {
			if (!handler) {
				return;
			}
			setIsLoading(true);
			try {
				const updated = await handler.applyLayoutToSlide(
					activeSlideIndex,
					layoutPath,
					slidesRef.current,
				);
				// Replace the slide in the slides array via ops
				ops.updateSlides((prev) => {
					const next = [...prev];
					next[activeSlideIndex] = updated;
					return next;
				});
				// Template/layout artwork is kept outside slide.elements in the React
				// viewer. Refresh that store after the relationship changes; otherwise
				// the canvas continues showing the old layout until the file is reopened.
				const templateElements = await handler.getTemplateElementsForSlide(updated.id);
				setTemplateElementsBySlideId((prev) => ({
					...prev,
					[updated.id]: templateElements,
				}));
				history.markDirty();
			} finally {
				setIsLoading(false);
			}
		},
		[handler, activeSlideIndex, ops, history, setTemplateElementsBySlideId],
	);

	return {
		availableLayouts,
		isLoading,
		loadAvailableLayouts,
		applyLayout,
		currentLayoutPath,
	};
}

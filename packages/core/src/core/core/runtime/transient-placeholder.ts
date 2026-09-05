import type { PptxElement } from '../../types';

/**
 * Marker set on placeholder elements the viewer materialises from a layout so
 * empty text areas are visible and clickable. They exist only in the loaded
 * model: the save writer drops them again unless the user typed something.
 */
export interface GeneratedPlaceholderMeta {
	_layoutSwitchGenerated?: boolean;
	/**
	 * PowerPoint creates slide-owned placeholder bindings when a layout is
	 * explicitly applied.  Those bindings can be visually empty while still
	 * being required for the layout's picture/fill artwork to render in the
	 * desktop app, so the save writer must retain them.
	 */
	_layoutSwitchPersistOnSave?: boolean;
}

/** Mark an element as a layout-derived placeholder that carries no content yet. */
export function markGeneratedPlaceholder(element: PptxElement): PptxElement {
	(element as PptxElement & GeneratedPlaceholderMeta)._layoutSwitchGenerated = true;
	return element;
}

/** Mark an explicitly applied layout placeholder as required in the PPTX. */
export function persistGeneratedPlaceholder(element: PptxElement): PptxElement {
	const metadata = element as PptxElement & GeneratedPlaceholderMeta;
	metadata._layoutSwitchGenerated = true;
	metadata._layoutSwitchPersistOnSave = true;
	return element;
}

/** Whether an otherwise-empty generated placeholder must be written. */
export function shouldPersistGeneratedPlaceholder(element: PptxElement): boolean {
	return Boolean((element as PptxElement & GeneratedPlaceholderMeta)._layoutSwitchPersistOnSave);
}

/**
 * Whether a generated placeholder is still empty, and therefore safe to drop on
 * save or to rebuild on the next layout change. Once the user types into it the
 * element becomes ordinary slide content.
 */
export function isEmptyGeneratedPlaceholder(element: PptxElement): boolean {
	if (!(element as PptxElement & GeneratedPlaceholderMeta)._layoutSwitchGenerated) {
		return false;
	}
	const withText = element as PptxElement & {
		text?: string;
		textSegments?: Array<{ text?: string }>;
	};
	if (withText.text && withText.text.trim().length > 0) {
		return false;
	}
	return !withText.textSegments?.some((segment) => (segment?.text ?? '').trim().length > 0);
}

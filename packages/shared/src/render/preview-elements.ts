import {
	isFullSlidePicturePlaceholder,
	type PptxElement,
	type PptxSlide,
} from 'pptx-viewer-core';

/**
 * Composition helper for slide previews and sidebar thumbnails.
 *
 * Every binding paints a preview from the same two sources as a real save:
 * the inherited layout/master (template) elements first, then the slide-owned
 * elements on top. Keeping that merge + cap in one place stops each binding's
 * thumbnail path from drifting away from `buildSaveSlides` ordering.
 */

/**
 * Default cap on the number of elements a preview renders. This guards against
 * pathological decks (thousands of shapes on one slide) blowing up a tiny
 * off-screen thumbnail; ordinary slides sit far below it, so normal content is
 * never dropped.
 */
export const DEFAULT_PREVIEW_ELEMENT_CAP = 500;

export interface BuildPreviewElementsOptions {
	/**
	 * Maximum number of elements to include. Defaults to
	 * {@link DEFAULT_PREVIEW_ELEMENT_CAP}. A value <= 0 disables the cap.
	 */
	cap?: number;
}

/**
 * Resolve the template layer shown for a slide. An explicit slide background
 * replaces the inherited background itself, so master/layout graphics remain
 * visible above it (including logos). Without an override, the native PPTX
 * showMasterShapes flag is honoured as usual.
 */
export function getVisibleTemplateElements(
	slide: PptxSlide | undefined,
	templateElements: readonly PptxElement[] = [],
): readonly PptxElement[] {
	const hasExplicitBackground =
		slide?.backgroundSource === 'slide' ||
		(slide?.backgroundSource === undefined &&
			Boolean(slide?.backgroundColor || slide?.backgroundImage || slide?.backgroundGradient));
	if (slide?.showMasterShapes === false) return [];
	if (hasExplicitBackground) {
		return templateElements.filter((element) => !isFullSlidePicturePlaceholder(element));
	}
	return templateElements;
}

/**
 * Ordered, capped element list for a slide preview/thumbnail. Inherited
 * template (layout/master) elements come first so slide-owned elements paint
 * on top, matching {@link import('./template-editing').buildSaveSlides}.
 */
export function buildPreviewElements(
	slide: PptxSlide,
	templateElements: readonly PptxElement[] = [],
	options?: BuildPreviewElementsOptions,
): PptxElement[] {
	const cap = options?.cap ?? DEFAULT_PREVIEW_ELEMENT_CAP;
	const inherited = getVisibleTemplateElements(slide, templateElements);
	const hasExplicitBackground =
		slide.backgroundSource === 'slide' ||
		(slide.backgroundSource === undefined &&
			Boolean(slide.backgroundColor || slide.backgroundImage || slide.backgroundGradient));
	const slideOwned = hasExplicitBackground
		? slide.elements.filter(
				(element) =>
					!element.id.startsWith('slide-layout-artwork-') ||
					!isFullSlidePicturePlaceholder(element),
			)
		: slide.elements;
	const merged = [...inherited, ...slideOwned];
	if (cap > 0 && merged.length > cap) {
		return merged.slice(0, cap);
	}
	return merged;
}

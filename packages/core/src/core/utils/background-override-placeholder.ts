import type { PptxElement, XmlObject } from '../types';

export const BACKGROUND_OVERRIDE_PLACEHOLDER_NAME_PREFIX =
	'AIXA_BACKGROUND_OVERRIDE_PLACEHOLDER__';
export const BACKGROUND_PRESERVED_ARTWORK_NAME_PREFIX =
	'AIXA_BACKGROUND_PRESERVED_ARTWORK__';

type PlaceholderElement = PptxElement & { placeholderType?: string };

function getPlaceholderNode(element: PptxElement): XmlObject | undefined {
	const raw = element.rawXml;
	if (!raw || typeof raw !== 'object') return undefined;
	const nonVisual = (raw['p:nvSpPr'] || raw['p:nvPicPr']) as XmlObject | undefined;
	return (nonVisual?.['p:nvPr'] as XmlObject | undefined)?.['p:ph'] as XmlObject | undefined;
}

function getPlaceholderType(element: PptxElement): string {
	const explicit = (element as PlaceholderElement).placeholderType;
	const fromXml = getPlaceholderNode(element)?.['@_type'];
	return String(explicit || fromXml || '').trim().toLowerCase();
}

/**
 * Detect picture-backed layout artwork that covers the slide. Some templates
 * use a picture placeholder for this layer, while others (including the
 * Swisscom title layouts) use an ordinary `p:pic`. Both must yield to an
 * explicit slide background; partial content-picture placeholders remain
 * available to the editor.
 */
export function isFullSlidePicturePlaceholder(element: PptxElement): boolean {
	const type = getPlaceholderType(element);
	const isPictureElement = element.type === 'picture' || element.type === 'image';
	return (
		(isPictureElement || type === 'pic' || type === 'sldimg') &&
		Math.abs(element.x) <= 2 &&
		Math.abs(element.y) <= 2 &&
		element.width >= 600 &&
		element.height >= 350
	);
}

/** Internal transparent slide binding written to suppress the layout backing. */
export function isBackgroundOverridePlaceholder(element: PptxElement): boolean {
	return String(element.name || '').startsWith(BACKGROUND_OVERRIDE_PLACEHOLDER_NAME_PREFIX);
}

/** Internal slide-owned copy of artwork retained while template graphics are hidden. */
export function isBackgroundPreservedArtwork(element: PptxElement): boolean {
	return String(element.name || '').startsWith(BACKGROUND_PRESERVED_ARTWORK_NAME_PREFIX);
}

/**
 * Copy inherited artwork onto a slide as an ordinary shape. This is used when
 * PowerPoint's all-or-nothing `showMasterSp` switch must hide a full-slide
 * layout picture placeholder without also losing logos and brand artwork.
 */
export function createBackgroundPreservedArtwork(element: PptxElement): PptxElement {
	const preserved = structuredClone(element) as PptxElement & {
		_layoutSwitchGenerated?: boolean;
		_layoutSwitchPersistOnSave?: boolean;
	};
	preserved.id = `aixa-background-preserved-${element.id}`;
	preserved.name = `${BACKGROUND_PRESERVED_ARTWORK_NAME_PREFIX}${element.name || 'Artwork'}`;
	preserved.shapeId = undefined;
	delete preserved._layoutSwitchGenerated;
	delete preserved._layoutSwitchPersistOnSave;

	const raw = preserved.rawXml;
	if (raw && typeof raw === 'object') {
		// Image-filled layout artwork is modelled as a picture for browser
		// rendering even when its native OOXML node is still a `p:sp`. Preserve
		// that native node kind when the artwork becomes slide-owned.
		if (raw['p:nvSpPr'] && (preserved.type === 'picture' || preserved.type === 'image')) {
			preserved.type = 'shape';
		}
		const nonVisual = (raw['p:nvSpPr'] || raw['p:nvPicPr'] || raw['p:nvGrpSpPr']) as
			| XmlObject
			| undefined;
		const cNvPr = nonVisual?.['p:cNvPr'] as XmlObject | undefined;
		if (cNvPr) cNvPr['@_name'] = preserved.name;
		const nvPr = nonVisual?.['p:nvPr'] as XmlObject | undefined;
		if (nvPr) delete nvPr['p:ph'];
	}
	return preserved;
}

/**
 * Build a slide-owned transparent placeholder with the same picture binding.
 * PowerPoint then treats the layout placeholder as fulfilled without painting
 * its grey/glass fill or custom prompt above the explicit slide background.
 */
export function createBackgroundOverridePlaceholder(element: PptxElement): PptxElement | null {
	if (!isFullSlidePicturePlaceholder(element)) return null;
	if (element.type !== 'shape' && element.type !== 'text') return null;

	const override = structuredClone(element) as PlaceholderElement & {
		text?: string;
		textSegments?: unknown[];
	};
	override.id = `aixa-background-override-${element.id}`;
	override.name = `${BACKGROUND_OVERRIDE_PLACEHOLDER_NAME_PREFIX}${element.name || 'Picture'}`;
	override.shapeId = undefined;
	override.text = '';
	override.textSegments = undefined;
	override.promptText = undefined;
	override.shapeStyle = {
		...override.shapeStyle,
		fillMode: 'none',
		fillColor: 'transparent',
		strokeColor: 'transparent',
		strokeWidth: 0,
	};

	const raw = override.rawXml;
	if (raw && typeof raw === 'object') {
		const nonVisual = (raw['p:nvSpPr'] || raw['p:nvPicPr']) as XmlObject | undefined;
		const cNvPr = nonVisual?.['p:cNvPr'] as XmlObject | undefined;
		if (cNvPr) cNvPr['@_name'] = override.name;
	}
	return override;
}

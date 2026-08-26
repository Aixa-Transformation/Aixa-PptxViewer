import type { PptxElement } from '../../types';

/**
 * Snapshot the editable model for a layout/master element.
 *
 * `rawXml` is excluded because save-time schema normalisation mutates it even
 * when the user made no change. Rebuilding such untouched template parts can
 * corrupt valid interleaved custom geometry, so dirty detection must use only
 * the editable model.
 */
export function fingerprintTemplateElement(element: PptxElement): string {
	return JSON.stringify(element, (key, value: unknown) =>
		key === 'rawXml' || key === 'shapeId' ? undefined : value,
	);
}

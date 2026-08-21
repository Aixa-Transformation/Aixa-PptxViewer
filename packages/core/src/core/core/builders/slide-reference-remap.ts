import type { PptxSlide, XmlObject } from '../../types';
import type { PptxSlideReferenceRemap } from '../../utils/presentation-collections';

export interface SlideReferenceRemapInput {
	/** Reconciled (final) slide list, each carrying its current id/path and rId. */
	slides: PptxSlide[];
	/** Load-time map: original slide rId -> slide path it pointed at. */
	originalRIdToPath: Map<string, string>;
	/** Load-time map: original numeric slide id -> slide path it pointed at. */
	originalSldIdToPath: Map<string, string>;
	/** Rebuilt `p:sldId` entries (each with `@_r:id` and `@_id`). */
	rebuiltSlideIds: XmlObject[];
}

/**
 * Derive the old->new slide-reference remapping from the load-time topology
 * and the reconciled slide list. Surviving slides get an entry in the rId /
 * sldId maps (identity when unchanged); slides that are gone are recorded in
 * the removed sets so their references can be dropped. `changed` stays false
 * for an unmodified round-trip so callers can keep byte-stable output.
 */
export function buildSlideReferenceRemap(init: SlideReferenceRemapInput): PptxSlideReferenceRemap {
	const pathToNewRId = new Map<string, string>();
	for (const slide of init.slides) {
		pathToNewRId.set(slide.id, slide.rId);
	}

	const newRIdToNumeric = new Map<string, string>();
	for (const entry of init.rebuiltSlideIds) {
		const relationshipId = String(entry?.['@_r:id'] ?? '');
		const numericSlideId = String(entry?.['@_id'] ?? '');
		if (relationshipId.length > 0 && numericSlideId.length > 0) {
			newRIdToNumeric.set(relationshipId, numericSlideId);
		}
	}
	const orderedSldIds = init.slides
		.map((slide) => newRIdToNumeric.get(slide.rId))
		.filter((id): id is string => id !== undefined);
	const originalPaths = new Set(init.originalSldIdToPath.values());
	const newSldIds = new Set<string>();
	for (const slide of init.slides) {
		if (!originalPaths.has(slide.id)) {
			const numericId = newRIdToNumeric.get(slide.rId);
			if (numericId) {
				newSldIds.add(numericId);
			}
		}
	}

	const rIdByOldRId = new Map<string, string>();
	const removedRIds = new Set<string>();
	let changed = false;
	for (const [oldRId, path] of init.originalRIdToPath.entries()) {
		const newRId = pathToNewRId.get(path);
		if (newRId === undefined) {
			removedRIds.add(oldRId);
			changed = true;
			continue;
		}
		rIdByOldRId.set(oldRId, newRId);
		if (newRId !== oldRId) {
			changed = true;
		}
	}

	const sldIdByOldSldId = new Map<string, string>();
	const removedSldIds = new Set<string>();
	for (const [oldSldId, path] of init.originalSldIdToPath.entries()) {
		const newRId = pathToNewRId.get(path);
		const newSldId = newRId === undefined ? undefined : newRIdToNumeric.get(newRId);
		if (newSldId === undefined) {
			removedSldIds.add(oldSldId);
			changed = true;
			continue;
		}
		sldIdByOldSldId.set(oldSldId, newSldId);
		if (newSldId !== oldSldId) {
			changed = true;
		}
	}

	changed ||= newSldIds.size > 0;
	return {
		rIdByOldRId,
		sldIdByOldSldId,
		removedRIds,
		removedSldIds,
		orderedSldIds,
		newSldIds,
		changed,
	};
}
// Modified by Aixa Ltd from the original ChristopherVR/pptx-viewer project.

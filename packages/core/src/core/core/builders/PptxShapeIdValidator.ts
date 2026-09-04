import type { XmlObject } from '../../types';
import {
	createShapeIdAllocator,
	parseShapeId,
	remapShapeIdReferences,
} from '../../utils/shape-ids';

/**
 * Shape ID uniqueness validator for OOXML slide shape trees.
 *
 * OpenXML requires that every `p:cNvPr/@id` within a single slide's
 * `p:spTree` is unique. Duplicate IDs can corrupt files in MS Office.
 * This validator scans the tree and reassigns duplicate IDs.
 */

/** Recursively collect all cNvPr nodes from a shape tree. */
function collectCnvPrNodes(
	node: XmlObject,
	results: XmlObject[],
	ensureArray: (value: unknown) => unknown[],
): void {
	// Check direct cNvPr references in nvSpPr, nvPicPr, nvCxnSpPr, nvGrpSpPr, nvGraphicFramePr
	const nvContainers = [
		'p:nvSpPr',
		'p:nvPicPr',
		'p:nvCxnSpPr',
		'p:nvGrpSpPr',
		'p:nvGraphicFramePr',
		'p:nvContentPartPr',
	];
	for (const nvKey of nvContainers) {
		const nvNode = node[nvKey] as XmlObject | undefined;
		if (nvNode?.['p:cNvPr']) {
			results.push(nvNode['p:cNvPr'] as XmlObject);
		}
	}

	// Recurse into shape lists
	const shapeLists = ['p:sp', 'p:pic', 'p:cxnSp', 'p:graphicFrame', 'p:grpSp', 'p:contentPart'];
	for (const listKey of shapeLists) {
		const children = ensureArray(node[listKey]) as XmlObject[];
		for (const child of children) {
			collectCnvPrNodes(child, results, ensureArray);
		}
	}

	// Modern ink and several other Office extensions place their real element
	// and fallback shape in mc:AlternateContent branches. Those nodes are still
	// part of the slide's non-visual ID space and must participate in duplicate
	// detection; otherwise a Draw operation can introduce a repeated id that
	// makes desktop PowerPoint repair or reject the deck.
	for (const alternate of ensureArray(node['mc:AlternateContent']) as XmlObject[]) {
		for (const branchKey of ['mc:Choice', 'mc:Fallback']) {
			for (const branch of ensureArray(alternate[branchKey]) as XmlObject[]) {
				collectCnvPrNodes(branch, results, ensureArray);
			}
		}
	}
}

export interface IPptxShapeIdValidator {
	validateAndDeduplicateIds(
		spTree: XmlObject,
		ensureArray: (value: unknown) => unknown[],
		referenceRoot?: XmlObject,
		remappedIds?: Map<string, string>,
	): number;
}

/**
 * Validates shape IDs in a slide's spTree and reassigns duplicates.
 * Returns the number of IDs that were reassigned.
 */
export class PptxShapeIdValidator implements IPptxShapeIdValidator {
	public validateAndDeduplicateIds(
		spTree: XmlObject,
		ensureArray: (value: unknown) => unknown[],
		referenceRoot: XmlObject = spTree,
		remappedIds: Map<string, string> = new Map(),
	): number {
		const cNvPrNodes: XmlObject[] = [];
		collectCnvPrNodes(spTree, cNvPrNodes, ensureArray);

		if (cNvPrNodes.length === 0) {
			return 0;
		}

		// Reserve every valid ID before allocating. Invalid IDs must never seed
		// allocation: one timestamp used to make all duplicate/group IDs overflow.
		const allocate = createShapeIdAllocator(cNvPrNodes.map((node) => node['@_id']));
		const usedIds = new Set<number>();
		let reassigned = 0;
		for (const cNvPr of cNvPrNodes) {
			const oldId = String(cNvPr['@_id'] ?? '').trim();
			const id = parseShapeId(oldId);
			if (id === undefined || usedIds.has(id)) {
				const replacement = allocate();
				cNvPr['@_id'] = replacement;
				reassigned += 1;
				// Duplicate valid IDs are ambiguous: references keep targeting the
				// first original shape. For invalid IDs, consistently target the first
				// repaired occurrence. Zero/missing IDs are unassigned sentinels.
				if (id === undefined && oldId && oldId !== '0' && !remappedIds.has(oldId)) {
					remappedIds.set(oldId, replacement);
				}
			} else {
				usedIds.add(id);
			}
		}
		remapShapeIdReferences(referenceRoot, remappedIds);
		return reassigned;
	}
}

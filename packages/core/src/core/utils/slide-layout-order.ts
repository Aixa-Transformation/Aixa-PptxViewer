import type { XmlObject } from '../types';
import { xmlAttr, xmlChild } from './xml-access';

function asArray(value: unknown): XmlObject[] {
	if (value === undefined || value === null) return [];
	return Array.isArray(value) ? (value as XmlObject[]) : [value as XmlObject];
}

/** Resolve layouts in the p:sldLayoutIdLst order used by PowerPoint's gallery. */
export function resolveSlideLayoutOrder(
	sldMaster: XmlObject,
	relationships: XmlObject[],
	resolveTarget: (target: string) => string,
): string[] {
	const layoutRelationships = relationships.filter((relationship) =>
		String(relationship['@_Type'] ?? '').includes('/slideLayout'),
	);
	const targetById = new Map(
		layoutRelationships.map((relationship) => [
			String(relationship['@_Id'] ?? ''),
			String(relationship['@_Target'] ?? ''),
		]),
	);
	const orderedIds = asArray(xmlChild(sldMaster, 'p:sldLayoutIdLst')?.['p:sldLayoutId'])
		.map((layoutId) => xmlAttr(layoutId, 'r:id'))
		.filter((id): id is string => Boolean(id));
	const orderedTargets: string[] = [];
	const usedIds = new Set<string>();
	for (const id of orderedIds) {
		const target = targetById.get(id);
		if (target) {
			orderedTargets.push(resolveTarget(target));
			usedIds.add(id);
		}
	}
	// Keep orphaned relationships at the end so malformed legacy decks remain lossless.
	for (const relationship of layoutRelationships) {
		const id = String(relationship['@_Id'] ?? '');
		const target = String(relationship['@_Target'] ?? '');
		if (target && !usedIds.has(id)) orderedTargets.push(resolveTarget(target));
	}
	return orderedTargets;
}

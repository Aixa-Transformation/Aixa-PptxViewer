import type { PptxElement, XmlObject } from '../types';

export const MAX_SHAPE_ID = 0xffffffff;

/** Native DrawingML IDs are UInt32s, not editor UUIDs or timestamps. */
export function parseShapeId(value: unknown, allowZero = false): number | undefined {
	const text = String(value ?? '').trim();
	if (!/^\d+$/.test(text)) return undefined;
	const id = Number(text);
	// Zero is a valid UInt32 in imported XML, but is reserved by our allocator
	// and treated as an unassigned sentinel by the existing element writers.
	return Number.isSafeInteger(id) && id >= (allowZero ? 0 : 1) && id <= MAX_SHAPE_ID
		? id
		: undefined;
}

/** Preserve existing IDs; wrap into free gaps when the UInt32 ceiling is occupied. */
export function createShapeIdAllocator(values: Iterable<unknown>): () => string {
	const used = new Set<number>();
	let next = 1;
	for (const value of values) {
		const id = parseShapeId(value);
		if (id !== undefined) {
			used.add(id);
			next = Math.max(next, id + 1);
		}
	}
	return () => {
		if (next > MAX_SHAPE_ID) next = 1;
		while (used.has(next)) next += 1;
		if (next > MAX_SHAPE_ID) throw new Error('No free UInt32 PowerPoint shape IDs remain.');
		used.add(next);
		return String(next++);
	};
}

export function visitXmlObjects(
	value: unknown,
	visit: (node: XmlObject, tag: string) => void,
	tag = '',
): void {
	if (Array.isArray(value)) {
		for (const item of value) visitXmlObjects(item, visit, tag);
	} else if (value && typeof value === 'object') {
		const node = value as XmlObject;
		visit(node, tag);
		for (const [key, child] of Object.entries(node)) {
			if (!key.startsWith('@_')) visitXmlObjects(child, visit, key);
		}
	}
}

export function elementShapeIds(elements: readonly PptxElement[]): unknown[] {
	const ids: unknown[] = [1]; // Reserve the implicit shape-tree group.
	for (const element of elements) {
		ids.push(element.shapeId);
		visitXmlObjects(element.rawXml, (node, tag) => {
			if (tag === 'p:cNvPr') ids.push(node['@_id']);
		});
		if (element.type === 'group') ids.push(...elementShapeIds(element.children));
	}
	return ids;
}

/** Update only shape-reference attributes, never relationship/slide/timing-node IDs. */
export function remapShapeIdReferences(root: unknown, ids: ReadonlyMap<string, string>): void {
	if (ids.size === 0) return;
	visitXmlObjects(root, (node, tag) => {
		const attrs: string[] = [];
		if (tag.startsWith('p:')) attrs.push('@_spid');
		if (tag === 'a:stCxn' || tag === 'a:endCxn') attrs.push('@_id');
		if (tag === 'pptx:animation') attrs.push('@_elementId', '@_triggerShapeId');
		for (const attr of attrs) {
			const replacement = ids.get(String(node[attr] ?? '').trim());
			if (replacement !== undefined) node[attr] = replacement;
		}
	});
}

/** Keep live model IDs consistent with a repair, including the next save without reload. */
export function remapElementShapeIds(
	elements: readonly PptxElement[],
	ids: ReadonlyMap<string, string>,
): void {
	if (ids.size === 0) return;
	for (const element of elements) {
		if (element.shapeId !== undefined)
			element.shapeId = ids.get(element.shapeId) ?? element.shapeId;
		visitXmlObjects(element.rawXml, (node, tag) => {
			if (tag === 'p:cNvPr') node['@_id'] = ids.get(String(node['@_id'])) ?? node['@_id'];
		});
		remapShapeIdReferences(element.rawXml, ids);
		if (element.type === 'connector') {
			for (const connection of [
				element.shapeStyle?.connectorStartConnection,
				element.shapeStyle?.connectorEndConnection,
			]) {
				if (connection?.shapeId)
					connection.shapeId = ids.get(connection.shapeId) ?? connection.shapeId;
			}
		}
		if (element.type === 'group') remapElementShapeIds(element.children, ids);
	}
}

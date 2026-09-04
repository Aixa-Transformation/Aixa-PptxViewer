import { describe, it, expect } from 'vitest';

import type { XmlObject } from '../../types';
import { PptxShapeIdValidator } from './PptxShapeIdValidator';

const ensureArray = (value: unknown): unknown[] => {
	if (Array.isArray(value)) {
		return value;
	}
	if (value === undefined || value === null) {
		return [];
	}
	return [value];
};

describe('pptxShapeIdValidator', () => {
	const validator = new PptxShapeIdValidator();

	it('repairs timestamps without spreading them to groups/duplicates and preserves references', () => {
		const ids = ['1', '1788524999615', '0', '4294967295', '2', '2', '2.5', '-1'];
		const shapes = ids.map((id) => ({ 'p:nvSpPr': { 'p:cNvPr': { '@_id': id } } }));
		const tree = { 'p:sp': shapes };
		const root = {
			'p:cSld': { 'p:spTree': tree },
			'p:timing': { 'p:spTgt': { '@_spid': ids[1] }, 'p:cTn': { '@_id': ids[1] } },
			'a:stCxn': { '@_id': ids[1] },
			'a:endCxn': { '@_id': '2' },
		};
		const remapped = new Map<string, string>();
		expect(validator.validateAndDeduplicateIds(tree, ensureArray, root, remapped)).toBe(5);
		const repaired = shapes.map((s) => s['p:nvSpPr']['p:cNvPr']['@_id']);
		expect(new Set(repaired).size).toBe(ids.length);
		expect(
			repaired.every((id) => /^\d+$/.test(id) && Number(id) > 0 && Number(id) <= 4294967295),
		).toBe(true);
		expect(root['p:timing']['p:spTgt']['@_spid']).toBe(repaired[1]);
		expect(root['a:stCxn']['@_id']).toBe(repaired[1]);
		expect(root['a:endCxn']['@_id']).toBe('2');
		expect(root['p:timing']['p:cTn']['@_id']).toBe(ids[1]);
		expect(validator.validateAndDeduplicateIds(tree, ensureArray, root)).toBe(0);
	});

	it('should return 0 when all IDs are unique', () => {
		const spTree: XmlObject = {
			'p:sp': [
				{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '2', '@_name': 'Shape 1' } } },
				{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '3', '@_name': 'Shape 2' } } },
			],
		};
		const result = validator.validateAndDeduplicateIds(spTree, ensureArray);
		expect(result).toBe(0);
	});

	it('should reassign duplicate IDs', () => {
		const spTree: XmlObject = {
			'p:sp': [
				{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '2', '@_name': 'Shape 1' } } },
				{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '2', '@_name': 'Shape 2' } } },
			],
		};
		const result = validator.validateAndDeduplicateIds(spTree, ensureArray);
		expect(result).toBe(1);

		const shapes = spTree['p:sp'] as XmlObject[];
		const id1 = (shapes[0]['p:nvSpPr'] as XmlObject)['p:cNvPr']['@_id'];
		const id2 = (shapes[1]['p:nvSpPr'] as XmlObject)['p:cNvPr']['@_id'];
		expect(id1).not.toBe(id2);
	});

	it('should reassign zero IDs', () => {
		const spTree: XmlObject = {
			'p:sp': [
				{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '0', '@_name': 'Shape 1' } } },
				{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '5', '@_name': 'Shape 2' } } },
			],
		};
		const result = validator.validateAndDeduplicateIds(spTree, ensureArray);
		expect(result).toBe(1);

		const shapes = spTree['p:sp'] as XmlObject[];
		const id1 = (shapes[0]['p:nvSpPr'] as XmlObject)['p:cNvPr']['@_id'];
		expect(id1).toBe('6');
	});

	it('should handle mixed element types (shapes, pics, connectors)', () => {
		const spTree: XmlObject = {
			'p:sp': { 'p:nvSpPr': { 'p:cNvPr': { '@_id': '2', '@_name': 'Shape' } } },
			'p:pic': { 'p:nvPicPr': { 'p:cNvPr': { '@_id': '2', '@_name': 'Pic' } } },
			'p:cxnSp': { 'p:nvCxnSpPr': { 'p:cNvPr': { '@_id': '2', '@_name': 'Connector' } } },
		};
		const result = validator.validateAndDeduplicateIds(spTree, ensureArray);
		expect(result).toBe(2);
	});

	it('should handle nested group shapes', () => {
		const spTree: XmlObject = {
			'p:grpSp': {
				'p:nvGrpSpPr': { 'p:cNvPr': { '@_id': '2', '@_name': 'Group' } },
				'p:sp': [
					{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '2', '@_name': 'Child1' } } },
					{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '3', '@_name': 'Child2' } } },
				],
			},
		};
		const result = validator.validateAndDeduplicateIds(spTree, ensureArray);
		expect(result).toBe(1);
	});

	it('deduplicates content-part and fallback ids inside AlternateContent', () => {
		const spTree: XmlObject = {
			'p:sp': {
				'p:nvSpPr': { 'p:cNvPr': { '@_id': '2', '@_name': 'Existing shape' } },
			},
			'mc:AlternateContent': {
				'mc:Choice': {
					'p:contentPart': {
						'p:nvContentPartPr': {
							'p:cNvPr': { '@_id': '2', '@_name': 'Ink content' },
						},
					},
				},
				'mc:Fallback': {
					'p:sp': {
						'p:nvSpPr': { 'p:cNvPr': { '@_id': '0', '@_name': 'Ink fallback' } },
					},
				},
			},
		};
		const result = validator.validateAndDeduplicateIds(spTree, ensureArray);
		expect(result).toBe(2);

		const alternate = spTree['mc:AlternateContent'] as XmlObject;
		const choice = alternate['mc:Choice'] as XmlObject;
		const contentPart = choice['p:contentPart'] as XmlObject;
		const contentNv = contentPart['p:nvContentPartPr'] as XmlObject;
		const fallback = alternate['mc:Fallback'] as XmlObject;
		const fallbackShape = fallback['p:sp'] as XmlObject;
		const fallbackNv = fallbackShape['p:nvSpPr'] as XmlObject;
		const ids = [
			'2',
			String((contentNv['p:cNvPr'] as XmlObject)['@_id']),
			String((fallbackNv['p:cNvPr'] as XmlObject)['@_id']),
		];
		expect(new Set(ids).size).toBe(3);
		expect(ids).not.toContain('0');
	});

	it('should return 0 for empty spTree', () => {
		const spTree: XmlObject = {};
		const result = validator.validateAndDeduplicateIds(spTree, ensureArray);
		expect(result).toBe(0);
	});

	it('should handle cloned shapes with all duplicate IDs', () => {
		const spTree: XmlObject = {
			'p:sp': [
				{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '5', '@_name': 'Original' } } },
				{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '5', '@_name': 'Clone 1' } } },
				{ 'p:nvSpPr': { 'p:cNvPr': { '@_id': '5', '@_name': 'Clone 2' } } },
			],
		};
		const result = validator.validateAndDeduplicateIds(spTree, ensureArray);
		expect(result).toBe(2);

		const shapes = spTree['p:sp'] as XmlObject[];
		const ids = shapes.map((s) => (s['p:nvSpPr'] as XmlObject)['p:cNvPr']['@_id']);
		const uniqueIds = new Set(ids);
		expect(uniqueIds.size).toBe(3);
	});
});

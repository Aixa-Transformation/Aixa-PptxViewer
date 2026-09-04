import { describe, expect, it } from 'vitest';
import { createShapeIdAllocator, parseShapeId, remapShapeIdReferences } from './shape-ids';
import { canonicalPlaceholderType, canonicalizePlaceholderTypes } from './placeholder-validation';

describe('native shape IDs and placeholder XML', () => {
	it('accepts imported UInt32 zero without allocating it to new shapes', () => {
		expect(parseShapeId('0', true)).toBe(0);
		expect(parseShapeId('0')).toBeUndefined();
		expect(createShapeIdAllocator(['0'])()).toBe('1');
	});
	it.each(['1788524999615', '4294967296', '-1', '2.5', '2x', '', '0', 'NaN', '1e3'])(
		'rejects invalid ID %s',
		(id) => expect(parseShapeId(id)).toBeUndefined(),
	);
	it('allocates bounded, unique IDs without a timestamp or overflow', () => {
		const allocate = createShapeIdAllocator(['1', '2', '4294967295', '1788524999615']);
		expect([allocate(), allocate(), allocate()]).toEqual(['3', '4', '5']);
		expect(parseShapeId('4294967295')).toBe(4294967295);
	});
	it.each([
		['ctrtitle', 'ctrTitle'],
		['subtitle', 'subTitle'],
		['sldnum', 'sldNum'],
		['sldimg', 'sldImg'],
		['clipart', 'clipArt'],
	])('canonicalizes %s to %s', (input, output) =>
		expect(canonicalPlaceholderType(input)).toBe(output),
	);
	it('does not change non-placeholder types or erase unknown extensions', () => {
		const root = {
			'p:sp': { 'p:ph': { '@_type': 'ctrtitle' } },
			'a:prstGeom': { '@_prst': 'rect' },
		};
		expect(canonicalizePlaceholderTypes(root)).toBe(1);
		expect(canonicalizePlaceholderTypes(root)).toBe(0);
		expect(root['a:prstGeom']['@_prst']).toBe('rect');
		expect(canonicalPlaceholderType('vendorThing')).toBe('vendorThing');
	});
	it('only remaps shape references, not timing IDs or relationship IDs', () => {
		const root = {
			'p:spTgt': { '@_spid': '12345' },
			'p:cTn': { '@_id': '12345' },
			'a:stCxn': { '@_id': '12345' },
			'a:blip': { '@_r:embed': '12345' },
			'pptx:animation': { '@_elementId': '12345' },
		};
		remapShapeIdReferences(root, new Map([['12345', '2']]));
		expect(root['p:spTgt']['@_spid']).toBe('2');
		expect(root['a:stCxn']['@_id']).toBe('2');
		expect(root['pptx:animation']['@_elementId']).toBe('2');
		expect(root['p:cTn']['@_id']).toBe('12345');
		expect(root['a:blip']['@_r:embed']).toBe('12345');
	});
});

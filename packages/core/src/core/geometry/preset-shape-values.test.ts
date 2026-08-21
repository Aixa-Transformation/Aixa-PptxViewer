import { describe, expect, it } from 'vitest';

import {
	DRAWINGML_PRESET_SHAPE_VALUES,
	normalizeDrawingmlPresetShape,
	sanitizeDrawingmlPresetGeometryTree,
} from './preset-shape-values';

describe('DrawingML preset shape values', () => {
	it('contains the complete unique Open XML ShapeTypeValues set', () => {
		expect(DRAWINGML_PRESET_SHAPE_VALUES.size).toBe(187);
		expect(DRAWINGML_PRESET_SHAPE_VALUES.has('cloudCallout')).toBe(true);
		expect(DRAWINGML_PRESET_SHAPE_VALUES.has('chartPlus')).toBe(true);
		expect(DRAWINGML_PRESET_SHAPE_VALUES.has('custom')).toBe(false);
	});

	it('normalizes aliases and rejects non-schema values', () => {
		expect(normalizeDrawingmlPresetShape('cylinder')).toBe('can');
		expect(normalizeDrawingmlPresetShape('rightTriangle')).toBe('rtTriangle');
		expect(normalizeDrawingmlPresetShape('custom')).toBe('rect');
		expect(normalizeDrawingmlPresetShape('unknownShape')).toBe('rect');
	});

	it('repairs invalid preset geometry recursively in preserved raw XML', () => {
		const rawXml = {
			'p:grpSp': {
				'p:sp': [
					{ 'p:spPr': { 'a:prstGeom': { '@_prst': 'custom' } } },
					{ 'p:spPr': { 'a:prstGeom': { '@_prst': 'ellipse' } } },
				],
			},
		};
		sanitizeDrawingmlPresetGeometryTree(rawXml);
		const shapes = rawXml['p:grpSp']['p:sp'];
		expect(shapes[0]['p:spPr']['a:prstGeom']['@_prst']).toBe('rect');
		expect(shapes[1]['p:spPr']['a:prstGeom']['@_prst']).toBe('ellipse');
	});
});

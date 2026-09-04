import { describe, expect, it } from 'vitest';

import type { PptxElement, XmlObject } from '../../types';
import { PptxHandlerRuntime } from './PptxHandlerRuntimeImplementation';

class TestRuntime extends PptxHandlerRuntime {
	public parseTestShape(shape: XmlObject): PptxElement | null {
		return this.parseShape(shape, 'shape-1');
	}
}

function shape(useBgFill?: string): XmlObject {
	return {
		...(useBgFill === undefined ? {} : { '@_useBgFill': useBgFill }),
		'p:nvSpPr': {
			'p:cNvPr': { '@_id': '1', '@_name': 'Background fill shape' },
			'p:cNvSpPr': {},
			'p:nvPr': {},
		},
		'p:spPr': {
			'a:xfrm': {
				'a:off': { '@_x': '0', '@_y': '0' },
				'a:ext': { '@_cx': '9144000', '@_cy': '5143500' },
			},
			'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} },
		},
		'p:style': {
			'a:fillRef': {
				'@_idx': '1',
				'a:schemeClr': { '@_val': 'accent1' },
			},
		},
	};
}

describe('p:sp useBgFill parsing', () => {
	it('preserves the semantic background-fill marker on the parsed shape', () => {
		const element = new TestRuntime().parseTestShape(shape('1'));

		expect(element).not.toBeNull();
		expect(element && 'shapeStyle' in element ? element.shapeStyle?.useBackgroundFill : false).toBe(
			true,
		);
	});

	it('does not mark ordinary shapes as background-fill shapes', () => {
		const element = new TestRuntime().parseTestShape(shape('0'));

		expect(element).not.toBeNull();
		expect(element && 'shapeStyle' in element ? element.shapeStyle?.useBackgroundFill : undefined).toBe(
			undefined,
		);
	});
});

import { describe, expect, it } from 'vitest';

import type { XmlObject } from '../types';
import { resolveSlideLayoutOrder } from './slide-layout-order';

describe('resolveSlideLayoutOrder', () => {
	it('uses master layout-id order instead of relationship order', () => {
		const master: XmlObject = {
			'p:sldLayoutIdLst': {
				'p:sldLayoutId': [
					{ '@_r:id': 'rId3' },
					{ '@_r:id': 'rId1' },
					{ '@_r:id': 'rId2' },
				],
			},
		};
		const relationships: XmlObject[] = [
			{ '@_Id': 'rId1', '@_Type': 'x/slideLayout', '@_Target': 'layout1.xml' },
			{ '@_Id': 'rId2', '@_Type': 'x/slideLayout', '@_Target': 'layout2.xml' },
			{ '@_Id': 'rId3', '@_Type': 'x/slideLayout', '@_Target': 'layout3.xml' },
		];
		expect(resolveSlideLayoutOrder(master, relationships, (target) => `ppt/${target}`)).toEqual([
			'ppt/layout3.xml',
			'ppt/layout1.xml',
			'ppt/layout2.xml',
		]);
	});

	it('appends orphaned layout relationships without dropping them', () => {
		const master: XmlObject = {
			'p:sldLayoutIdLst': { 'p:sldLayoutId': { '@_r:id': 'rId2' } },
		};
		const relationships: XmlObject[] = [
			{ '@_Id': 'rId1', '@_Type': 'x/slideLayout', '@_Target': 'orphan.xml' },
			{ '@_Id': 'rId2', '@_Type': 'x/slideLayout', '@_Target': 'first.xml' },
		];
		expect(resolveSlideLayoutOrder(master, relationships, (target) => target)).toEqual([
			'first.xml',
			'orphan.xml',
		]);
	});
});

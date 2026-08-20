import { describe, expect, it } from 'vitest';

import { getVisibleInspectorTabs } from './InspectorPaneHeader';

describe('getVisibleInspectorTabs', () => {
	it('shows every inspector tab by default', () => {
		expect(getVisibleInspectorTabs().map(({ key }) => key)).toEqual([
			'elements',
			'properties',
			'comments',
		]);
	});

	it('can restrict the inspector to Properties', () => {
		expect(
			getVisibleInspectorTabs(['inspectorElements', 'inspectorComments']).map(({ key }) => key),
		).toEqual(['properties']);
	});
});

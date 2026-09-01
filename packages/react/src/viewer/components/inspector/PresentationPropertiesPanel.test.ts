import { describe, expect, it } from 'vitest';

import { getPresentationPropertiesVisibility } from './PresentationPropertiesPanel';

describe('getPresentationPropertiesVisibility', () => {
	it('shows every optional presentation property card by default', () => {
		expect(getPresentationPropertiesVisibility()).toEqual({
			notesHandout: true,
			document: true,
			tags: true,
			slideSize: true,
			slideSummary: true,
		});
	});

	it('can independently hide Notes & Handout, Document, and Tags', () => {
		expect(
			getPresentationPropertiesVisibility([
				'inspectorNotesHandout',
				'inspectorDocument',
				'inspectorTags',
				'inspectorSlideSize',
				'inspectorSlideSummary',
			]),
		).toEqual({
			notesHandout: false,
			document: false,
			tags: false,
			slideSize: false,
			slideSummary: false,
		});
	});

	it('does not let unrelated visibility flags hide these cards', () => {
		expect(getPresentationPropertiesVisibility(['inspectorComments'])).toEqual({
			notesHandout: true,
			document: true,
			tags: true,
			slideSize: true,
			slideSummary: true,
		});
	});
});

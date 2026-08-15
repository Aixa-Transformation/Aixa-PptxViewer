import { describe, expect, it } from 'vitest';

import { isFullSlideBackgroundPicture } from './SlideCanvas';

describe('isFullSlideBackgroundPicture', () => {
	const canvas = { width: 1280, height: 720 };

	it('detects an edge-to-edge picture used as slide artwork', () => {
		expect(
			isFullSlideBackgroundPicture(
				{ type: 'picture', x: 0, y: 0, width: 1280, height: 720 },
				canvas,
			),
		).toBe(true);
	});

	it('does not disable ordinary pictures or non-picture elements', () => {
		expect(
			isFullSlideBackgroundPicture(
				{ type: 'picture', x: 640, y: 0, width: 640, height: 720 },
				canvas,
			),
		).toBe(false);
		expect(
			isFullSlideBackgroundPicture(
				{ type: 'text', x: 0, y: 0, width: 1280, height: 720 },
				canvas,
			),
		).toBe(false);
	});
});

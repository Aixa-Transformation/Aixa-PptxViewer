import type { MediaPptxElement, ShapePptxElement } from 'pptx-viewer-core';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { STATIC_ELEMENTS_ARE_PRESENTATION_PASSIVE } from '../StaticElementRenderer';
import { renderMediaElement } from '../../utils/media-render';
import { getChevronTextFrameStyle, shouldRenderTextBody } from './TextElementBody';

describe('passive slide fidelity', () => {
	it('treats inherited static elements as passive', () => {
		expect(STATIC_ELEMENTS_ARE_PRESENTATION_PASSIVE).toBe(true);
	});

	it('hides placeholder prompt text outside the editable canvas', () => {
		expect(shouldRenderTextBody(true, false, 'Click to edit Master title style', true)).toBe(false);
		expect(shouldRenderTextBody(true, false, 'Click to add title', false)).toBe(true);
		expect(shouldRenderTextBody(true, true, 'Click to add title', true)).toBe(true);
		expect(shouldRenderTextBody(true, true, undefined, true)).toBe(true);
	});

	it('keeps chevron text between the rear notch and arrow tip', () => {
		const element = {
			id: 'chevron-1',
			type: 'shape',
			shapeType: 'chevron',
			x: 0,
			y: 0,
			width: 400,
			height: 100,
		} as ShapePptxElement;

		expect(getChevronTextFrameStyle(element)).toStrictEqual({
			position: 'absolute',
			left: 50,
			top: 0,
			width: 300,
			height: 100,
		});
	});

	it('renders a valid poster at full opacity when linked media is unavailable', () => {
		const element = {
			id: 'video-1',
			type: 'media',
			x: 0,
			y: 0,
			width: 320,
			height: 180,
			mediaType: 'video',
			mediaMissing: true,
			posterFrameData: 'data:image/png;base64,AAAA',
		} as MediaPptxElement;

		const passive = renderToStaticMarkup(
			<>{renderMediaElement(element, new Map(), { isPresentationMode: true })}</>,
		);
		expect(passive).toContain('data:image/png;base64,AAAA');
		expect(passive).not.toContain('opacity-50');

		const editable = renderToStaticMarkup(<>{renderMediaElement(element, new Map())}</>);
		expect(editable).toContain('opacity-50');
	});
});

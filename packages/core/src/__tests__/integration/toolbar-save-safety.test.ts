import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxElement, PptxSlide } from '../../core/types';

const TINY_PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==';

async function saveAndReload(handler: PptxHandler, slides: PptxSlide[]) {
	const bytes = await handler.save(slides);
	const nextHandler = new PptxHandler();
	const nextData = await nextHandler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	return { bytes, handler: nextHandler, slides: nextData.slides };
}

function firstTextElement(slide: PptxSlide): PptxElement & Record<string, any> {
	const element = slide.elements.find((candidate) => 'text' in candidate);
	if (!element) {
		throw new Error('Expected a text-capable element');
	}
	return element as PptxElement & Record<string, any>;
}

describe('toolbar mutations remain save-safe', () => {
	it('survives formatting, insert, arrange, animation and slide CRUD saves', async () => {
		const created = await PresentationBuilder.create();
		const slides = [
			created.createSlide('Blank').addText('Alpha', { x: 40, y: 40, width: 280, height: 60 }).build(),
			created
				.createSlide('Blank')
				.addText('Beta', { x: 40, y: 40, width: 280, height: 60 })
				.addShape('rect', { x: 80, y: 140, width: 180, height: 100 })
				.addImage(TINY_PNG, { x: 300, y: 140, width: 80, height: 80 })
				.build(),
			created.createSlide('Blank').addText('Gamma', { x: 40, y: 40, width: 280, height: 60 }).build(),
		];

		let roundTrip = await saveAndReload(created.handler, slides);

		// Home/paragraph/font toolbar operations.
		const text = firstTextElement(roundTrip.slides[0]);
		text.text = 'Edited from toolbar';
		text.textStyle = {
			...(text.textStyle ?? {}),
			fontFamily: 'Arial',
			fontSize: 26,
			bold: true,
			italic: true,
			underline: true,
			color: '#C62828',
			align: 'center',
			vAlign: 'middle',
		};
		text.rotation = 12;
		text.x += 15;
		text.y += 10;
		text.width += 30;
		roundTrip.slides[0].isDirty = true;
		roundTrip = await saveAndReload(roundTrip.handler, roundTrip.slides);
		expect(firstTextElement(roundTrip.slides[0]).text).toBe('Edited from toolbar');

		// Drawing/arrange/animation/transition operations plus element deletion.
		const shape = roundTrip.slides[1].elements.find((element) => element.type === 'shape') as
			| (PptxElement & Record<string, any>)
			| undefined;
		expect(shape).toBeDefined();
		if (shape) {
			shape.fill = { color: '#1565C0' };
			shape.stroke = { color: '#FFB300', width: 3 };
			shape.rotation = 28;
			shape.x += 20;
			shape.y += 15;
			roundTrip.slides[1].animations = [
				{ elementId: shape.id, entrance: 'fade', durationMs: 500, order: 0, trigger: 'onClick' },
			];
		}
		roundTrip.slides[1].transition = { type: 'fade', durationMs: 650 };
		roundTrip.slides[1].elements = roundTrip.slides[1].elements.filter(
			(element) => element.type !== 'picture',
		);
		roundTrip.slides[1].isDirty = true;
		roundTrip = await saveAndReload(roundTrip.handler, roundTrip.slides);
		expect(roundTrip.slides[1].elements.some((element) => element.type === 'picture')).toBeFalsy();

		// Slide toolbar operations in one save: add in the middle, reorder, delete.
		const inserted = created
			.createSlide('Blank')
			.addText('Inserted', { x: 60, y: 60, width: 240, height: 60 })
			.build();
		roundTrip.slides.splice(1, 0, inserted);
		roundTrip.slides.unshift(roundTrip.slides.pop()!);
		roundTrip.slides.splice(2, 1);
		roundTrip = await saveAndReload(roundTrip.handler, roundTrip.slides);

		expect(roundTrip.slides).toHaveLength(3);
		expect(roundTrip.slides.map((slide) => slide.slideNumber)).toStrictEqual([1, 2, 3]);
		expect(roundTrip.bytes.byteLength).toBeGreaterThan(0);
	});
});

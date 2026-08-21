import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import type { PicturePptxElement, PptxSlide } from '../../core/types';

const TINY_PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==';

function asArrayBuffer(bytes: Uint8Array): ArrayBuffer {
	return bytes.buffer.slice(
		bytes.byteOffset,
		bytes.byteOffset + bytes.byteLength,
	) as ArrayBuffer;
}

async function buildPictureDeck(): Promise<Uint8Array> {
	const created = await PresentationBuilder.create();
	const slide = created
		.createSlide('Blank')
		.addImage(TINY_PNG, { x: 80, y: 60, width: 240, height: 180 })
		.build();
	return created.handler.save([slide]);
}

async function slideXml(bytes: Uint8Array): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	return zip.file('ppt/slides/slide1.xml')!.async('string');
}

function pictureFrom(slide: PptxSlide): PicturePptxElement {
	const picture = slide.elements.find((element) => element.type === 'picture');
	if (!picture || picture.type !== 'picture') {
		throw new Error('Expected a picture element');
	}
	return picture;
}

describe('preset geometry save safety', () => {
	it('never serializes the custom sentinel as a preset geometry value', async () => {
		const source = await buildPictureDeck();
		const handler = new PptxHandler();
		const data = await handler.load(asArrayBuffer(source));
		const picture = pictureFrom(data.slides[0]);

		picture.shapeType = 'custom';
		picture.customGeometryPaths = undefined;
		data.slides[0].isDirty = true;

		const saved = await handler.save(data.slides);
		const xml = await slideXml(saved);
		expect(xml).not.toMatch(/<a:prstGeom\b[^>]*\bprst="custom"/u);
		expect(xml).toMatch(/<a:prstGeom\b[^>]*\bprst="rect"/u);
	});

	it('round-trips real custom picture geometry as a:custGeom', async () => {
		const source = await buildPictureDeck();
		const sourceZip = await JSZip.loadAsync(source);
		const originalXml = await sourceZip.file('ppt/slides/slide1.xml')!.async('string');
		const customGeometry = [
			'<a:custGeom>',
			'<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>',
			'<a:rect l="l" t="t" r="r" b="b"/>',
			'<a:pathLst><a:path w="100" h="100">',
			'<a:moveTo><a:pt x="0" y="0"/></a:moveTo>',
			'<a:lnTo><a:pt x="100" y="0"/></a:lnTo>',
			'<a:lnTo><a:pt x="50" y="100"/></a:lnTo>',
			'<a:close/>',
			'</a:path></a:pathLst>',
			'</a:custGeom>',
		].join('');
		const injectedXml = originalXml.replace(
			/<a:prstGeom\b[^>]*>(?:[\s\S]*?)<\/a:prstGeom>|<a:prstGeom\b[^>]*\/>/u,
			customGeometry,
		);
		expect(injectedXml).not.toBe(originalXml);
		sourceZip.file('ppt/slides/slide1.xml', injectedXml);
		const injected = await sourceZip.generateAsync({ type: 'uint8array' });

		const handler = new PptxHandler();
		const data = await handler.load(asArrayBuffer(injected));
		const picture = pictureFrom(data.slides[0]);
		expect(picture.shapeType).toBe('custom');
		expect(picture.customGeometryPaths?.[0]?.segments).toHaveLength(4);

		data.slides[0].isDirty = true;
		const saved = await handler.save(data.slides);
		const xml = await slideXml(saved);
		expect(xml).toContain('<a:custGeom>');
		expect(xml).not.toMatch(/<a:prstGeom\b[^>]*\bprst="custom"/u);
	});
});

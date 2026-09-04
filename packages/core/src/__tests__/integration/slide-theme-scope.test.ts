import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';

const parser = new XMLParser({ ignoreAttributes: false });
const array = <T>(value: T | T[] | undefined): T[] =>
	value === undefined ? [] : Array.isArray(value) ? value : [value];

async function fixture() {
	const { handler, data, createSlide } = await PresentationBuilder.create({
		theme: { colors: { accent1: '#FF0000' } },
	});
	data.slides = [0, 1, 2].map(() => createSlide('Blank').build());
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const theme = await zip.file('ppt/theme/theme1.xml')!.async('string');
	zip.file('ppt/theme/theme2.xml', theme.replace(/FF0000/gi, '0000FF'));
	const types = await zip.file('[Content_Types].xml')!.async('string');
	zip.file(
		'[Content_Types].xml',
		types.replace(
			'</Types>',
			'<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
		),
	);
	for (const file of zip.file(/^ppt\/slides\/slide\d+\.xml$/)) {
		const xml = await file.async('string');
		zip.file(
			file.name,
			xml.replace(
				'</p:spTree>',
				'<p:sp><p:nvSpPr><p:cNvPr id="30" name="Theme rectangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr></p:sp></p:spTree>',
			),
		);
	}
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	const reader = new PptxHandler();
	return { handler: reader, data: await reader.load(bytes), zip };
}

function overrideTarget(zip: JSZip, path: string) {
	return zip
		.file(path.replace(/\/([^/]+)$/, '/_rels/$1.rels'))!
		.async('string')
		.then((xml) => {
			const rels = array(parser.parse(xml).Relationships.Relationship) as Record<string, string>[];
			return rels.find((r) => r['@_Type'].endsWith('/themeOverride'))?.['@_Target'];
		});
}

describe('slide-scoped theme Apply', () => {
	it('changes only the selected non-first slide, persists, and leaves shared masters alone', async () => {
		const { handler, data, zip } = await fixture();
		await handler.setSlidesTheme('ppt/theme/theme2.xml', [1]);
		const bytes = await handler.save(data.slides);
		const saved = await JSZip.loadAsync(bytes);
		expect(await overrideTarget(saved, data.slides[0].id)).toBeUndefined();
		expect(await overrideTarget(saved, data.slides[1].id)).toContain('themeOverride');
		expect(await overrideTarget(saved, data.slides[2].id)).toBeUndefined();
		for (const file of zip.file(/^ppt\/slideMasters\/_rels\/.+\.rels$/)) {
			expect(await saved.file(file.name)!.async('string')).toBe(await file.async('string'));
		}
		const reloaded = await new PptxHandler().load(bytes);
		expect(
			reloaded.slides.map((s) =>
				s.elements.find((e) => e.shapeStyle?.fillColor)?.shapeStyle?.fillColor?.toUpperCase(),
			),
		).toEqual(['#FF0000', '#0000FF', '#FF0000']);
	});

	it('applies to every slide including one with an earlier override; repeated Apply does not add parts', async () => {
		const { handler, data } = await fixture();
		await handler.setSlidesTheme('ppt/theme/theme1.xml', [1]);
		await handler.setSlidesTheme('ppt/theme/theme2.xml');
		await handler.setSlidesTheme('ppt/theme/theme2.xml');
		const bytes = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(bytes);
		expect(zip.file(/^ppt\/theme\/themeOverride\d+\.xml$/)).toHaveLength(3);
		const reloaded = await new PptxHandler().load(bytes);
		expect(
			reloaded.slides.map((s) =>
				s.elements.find((e) => e.shapeStyle?.fillColor)?.shapeStyle?.fillColor?.toUpperCase(),
			),
		).toEqual(['#0000FF', '#0000FF', '#0000FF']);
	});

	it('rejects invalid themes/indices without partially applying a selection', async () => {
		const { handler, data } = await fixture();
		await expect(handler.setSlidesTheme('missing.xml', [0])).rejects.toThrow();
		await expect(handler.setSlidesTheme('ppt/theme/theme2.xml', [0, 99])).rejects.toThrow();
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		expect(zip.file(/^ppt\/theme\/themeOverride\d+\.xml$/)).toHaveLength(0);
	});

	it('can save/reopen and switch again without baking theme colours into fixed RGB', async () => {
		const { handler, data } = await fixture();
		await handler.setSlidesTheme('ppt/theme/theme2.xml', [1]);
		const reader = new PptxHandler();
		const first = await reader.load(await handler.save(data.slides));
		// Normal autosave after reopening, then a second Apply using the other theme.
		await reader.save(first.slides);
		await reader.setSlidesTheme('ppt/theme/theme1.xml', [1]);
		const final = await new PptxHandler().load(await reader.save(first.slides));
		expect(
			final.slides.map((s) =>
				s.elements.find((e) => e.shapeStyle?.fillColor)?.shapeStyle?.fillColor?.toUpperCase(),
			),
		).toEqual(['#FF0000', '#FF0000', '#FF0000']);
	});
});

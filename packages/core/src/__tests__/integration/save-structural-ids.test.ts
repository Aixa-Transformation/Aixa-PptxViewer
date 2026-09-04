import fs from 'node:fs';
import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxElement, XmlObject } from '../../core/types';
import { parseShapeId, visitXmlObjects } from '../../core/utils/shape-ids';

const parser = new XMLParser({ ignoreAttributes: false, parseAttributeValue: false });
const buffer = (bytes: Uint8Array): ArrayBuffer =>
	bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
const own = (element: PptxElement) => !/^(master|layout)-/.test(element.id);

async function validateArchive(bytes: Uint8Array): Promise<JSZip> {
	const zip = await JSZip.loadAsync(bytes);
	for (const path of Object.keys(zip.files).filter((name) =>
		/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name),
	)) {
		visitXmlObjects(parser.parse(await zip.file(path)!.async('string')), (node, tag) => {
			if (tag === 'p:cNvPr')
				expect(parseShapeId(node['@_id'], true), `${path}: ${node['@_id']}`).toBeDefined();
			if (tag === 'p:ph')
				expect(['ctrtitle', 'subtitle', 'sldnum', 'sldimg', 'clipart']).not.toContain(
					node['@_type'],
				);
		});
	}
	return zip;
}

describe('PPTX save structural corruption regression', () => {
	it('creates valid placeholders and retains canonical types through edit/save/reopen', async () => {
		const seed = await PptxHandler.createBlank({ initialSlideCount: 1 });
		const layout = seed.handler.getLayoutOptions().find((option) => option.name === 'Title Slide')!;
		expect(layout).toBeDefined();
		const zip = await JSZip.loadAsync(await seed.handler.save(seed.data.slides));
		// The SDK's minimal default layouts carry no authored placeholders.
		// Seed real layout XML so this test exercises the production load/switch path.
		zip.file(
			layout.path,
			`<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld name="Title Slide"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${['ctrTitle', 'subTitle'].map((type, idx) => `<p:sp><p:nvSpPr><p:cNvPr id="${idx + 2}" name="${type}"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}" idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="${100000 + idx * 1000000}"/><a:ext cx="5000000" cy="800000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr sz="2400"/></a:p></p:txBody></p:sp>`).join('')}</p:spTree></p:cSld></p:sldLayout>`,
		);
		const handler = new PptxHandler();
		const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
		data.slides[0].elements = [];
		data.slides[0] = await handler.applyLayoutToSlide(0, layout.path, data.slides);
		const placeholders = data.slides[0].elements.filter(own);
		expect(placeholders.length).toBeGreaterThanOrEqual(2);
		for (const element of placeholders) {
			expect(parseShapeId(element.shapeId)).toBeDefined();
			if (element.type === 'text') element.text = `Typed ${element.id}`;
		}
		expect(new Set(placeholders.map((e) => e.shapeId)).size).toBe(placeholders.length);
		const first = await handler.save(data.slides);
		await validateArchive(first);
		await validateArchive(await handler.save(data.slides));
		const reopened = await new PptxHandler().load(buffer(first));
		expect(
			reopened.slides[0].elements.some((el) => 'text' in el && el.text?.startsWith('Typed ')),
		).toBe(true);
	});

	it('repairs oversized targets, group IDs, connectors and animations across consecutive saves', async () => {
		const { handler, data, createSlide } = await PptxHandler.createBlank();
		const slide = createSlide('Blank')
			.addText('Target', { x: 20, y: 20, width: 250, height: 80 })
			.build();
		const target = slide.elements[0];
		target.shapeId = '1788524999615';
		const child: PptxElement = {
			type: 'shape',
			id: 'child',
			shapeId: '1788524999616',
			x: 0,
			y: 0,
			width: 50,
			height: 50,
			shapeType: 'rect',
		};
		const group: PptxElement = {
			type: 'group',
			id: 'group',
			shapeId: '1788524999617',
			x: 300,
			y: 20,
			width: 100,
			height: 100,
			children: [child],
		};
		const connector: PptxElement = {
			type: 'connector',
			id: 'connector',
			shapeId: '8',
			x: 20,
			y: 20,
			width: 100,
			height: 30,
			shapeStyle: {
				connectorStartConnection: { shapeId: target.shapeId, connectionSiteIndex: 0 },
				connectorEndConnection: { shapeId: group.shapeId, connectionSiteIndex: 1 },
			},
		};
		slide.elements.push(group, connector);
		slide.animations = [target, child, group].map((el) => ({
			elementId: el.id,
			entrance: 'fadeIn',
		}));
		data.slides.push(slide);
		for (let cycle = 0; cycle < 3; cycle += 1) {
			const zip = await validateArchive(await handler.save(data.slides));
			const xml = parser.parse(await zip.file(slide.id)!.async('string')) as XmlObject;
			const ids = new Set<string>();
			const references: string[] = [];
			visitXmlObjects(xml, (node, tag) => {
				if (tag === 'p:cNvPr') ids.add(String(node['@_id']));
				if (tag === 'p:spTgt' || tag === 'p:bldP') references.push(String(node['@_spid']));
				if (tag === 'a:stCxn' || tag === 'a:endCxn') references.push(String(node['@_id']));
			});
			expect(references.length).toBeGreaterThanOrEqual(5);
			for (const reference of references)
				expect(ids.has(reference), `cycle ${cycle}, reference ${reference}`).toBe(true);
			for (const el of [target, child, group]) expect(parseShapeId(el.shapeId)).toBeDefined();
			expect(connector.shapeStyle?.connectorStartConnection?.shapeId).toBe(target.shapeId);
		}
	});

	it.each(['1788524999615', '0'])(
		'validates imported ID %s even when the slide is marked unchanged',
		async (id) => {
			const { handler, data, createSlide } = await PptxHandler.createBlank();
			data.slides.push(createSlide('Blank').addText('Retained text').build());
			const zip = await JSZip.loadAsync(await handler.save(data.slides));
			const path = data.slides[0].id;
			const xml = await zip.file(path)!.async('string');
			zip.file(
				path,
				xml.replace(/(<p:cNvPr\b[^>]*\bid=")[^"]+/, (_match, prefix) => prefix + id),
			);
			const reopenedHandler = new PptxHandler();
			const reopened = await reopenedHandler.load(await zip.generateAsync({ type: 'arraybuffer' }));
			reopened.slides[0].isDirty = false;
			await validateArchive(await reopenedHandler.save(reopened.slides));
		},
	);
});

// Opt-in real-file regression: fixtures stay outside Git and are never overwritten.
const brokenPath = process.env.PPTX_CORRUPT_FIXTURE;
const masterPath = process.env.PPTX_MASTER_FIXTURE;
describe.skipIf(!brokenPath || !masterPath)('supplied Swisscom decks', () => {
	it('repairs the broken file and retains layout9 XML and relationships', async () => {
		const warning = vi.spyOn(console, 'warn').mockImplementation(() => {});
		const info = vi.spyOn(console, 'info').mockImplementation(() => {});
		try {
			const original = fs.readFileSync(brokenPath!);
			const source = await JSZip.loadAsync(original);
			let handler = new PptxHandler();
			let data = await handler.load(buffer(original));
			for (let cycle = 0; cycle < 3; cycle += 1) {
				const saved = await handler.save(data.slides);
				const zip = await validateArchive(saved);
				for (const path of [
					'ppt/slideLayouts/slideLayout9.xml',
					'ppt/slideLayouts/_rels/slideLayout9.xml.rels',
				]) {
					expect(await zip.file(path)!.async('string')).toBe(
						await source.file(path)!.async('string'),
					);
				}
				handler = new PptxHandler();
				data = await handler.load(buffer(saved));
				expect(data.slides).toHaveLength(3);
			}
		} finally {
			warning.mockRestore();
			info.mockRestore();
		}
	}, 60_000);

	it('switches through every master layout, types into placeholders, and saves valid XML', async () => {
		const warning = vi.spyOn(console, 'warn').mockImplementation(() => {});
		const info = vi.spyOn(console, 'info').mockImplementation(() => {});
		try {
			const original = fs.readFileSync(masterPath!);
			const handler = new PptxHandler();
			const data = await handler.load(buffer(original));
			const layouts = await handler.getAvailableLayoutsForSlide(0, data.slides);
			expect(layouts).toHaveLength(30);
			for (const layout of layouts) {
				data.slides[0].elements = [];
				data.slides[0] = await handler.applyLayoutToSlide(0, layout.path, data.slides);
				for (const el of data.slides[0].elements.filter(own)) {
					if (el.type === 'text') el.text = `Typed placeholder for ${layout.name}`;
				}
				await validateArchive(await handler.save(data.slides));
			}
		} finally {
			warning.mockRestore();
			info.mockRestore();
		}
	}, 120_000);
});

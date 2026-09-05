import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import type { PptxElement } from '../../core/types';

async function savedSlideXml(options: {
	backgroundSource: 'slide' | 'inherited';
	backgroundColor?: string;
	includeFullSlidePicturePlaceholder?: boolean;
	showMasterShapes?: boolean;
}): Promise<string> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	try {
		const slide = createSlide('Blank').build();
		slide.backgroundSource = options.backgroundSource;
		slide.backgroundColor = options.backgroundColor;
		slide.showMasterShapes = options.showMasterShapes ?? true;
		if (options.includeFullSlidePicturePlaceholder) {
			slide.elements.unshift(
			{
				id: 'layout-placeholder-picture',
				type: 'shape',
				x: 0,
				y: 0,
				width: 1280,
				height: 720,
				text: '',
				promptText: 'Long layout authoring instructions that must not cover the background',
				shapeStyle: { fillColor: '#DDE3E7' },
				rawXml: {
					'p:nvSpPr': {
						'p:cNvPr': { '@_id': '13', '@_name': 'Picture Placeholder 8' },
						'p:cNvSpPr': {},
						'p:nvPr': { 'p:ph': { '@_type': 'pic', '@_idx': '13' } },
					},
					'p:spPr': {
						'a:xfrm': {
							'a:off': { '@_x': '0', '@_y': '0' },
							'a:ext': { '@_cx': '12192000', '@_cy': '6858000' },
						},
						'a:prstGeom': { '@_prst': 'rect', 'a:avLst': '' },
						'a:solidFill': { 'a:srgbClr': { '@_val': 'DDE3E7' } },
					},
					'p:txBody': {
						'a:bodyPr': {},
						'a:lstStyle': {},
						'a:p': { 'a:r': { 'a:t': 'Long layout authoring instructions' } },
					},
				},
			} as PptxElement,
			{
				id: 'layout-logo',
				name: 'Swisscom',
				type: 'shape',
				x: 20,
				y: 20,
				width: 200,
				height: 60,
				text: '',
				shapeStyle: { fillColor: '#FF0000' },
				rawXml: {
					'p:nvSpPr': {
						'p:cNvPr': { '@_id': '14', '@_name': 'Swisscom' },
						'p:cNvSpPr': {},
						'p:nvPr': {},
					},
					'p:spPr': {
						'a:xfrm': {
							'a:off': { '@_x': '190500', '@_y': '190500' },
							'a:ext': { '@_cx': '1905000', '@_cy': '571500' },
						},
						'a:prstGeom': { '@_prst': 'rect', 'a:avLst': '' },
						'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } },
					},
				},
			} as PptxElement,
			);
		}
		data.slides.push(slide);

		const bytes = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(bytes);
		const xml = await zip.file(slide.id)?.async('string');
		if (!xml) throw new Error(`Saved slide part is missing: ${slide.id}`);
		return xml;
	} finally {
		handler.dispose();
	}
}

describe('slide background save', () => {
	it('suppresses a full-slide layout fill while preserving the other artwork', async () => {
		const xml = await savedSlideXml({
			backgroundSource: 'slide',
			backgroundColor: '#102030',
			includeFullSlidePicturePlaceholder: true,
		});
		expect(xml).toContain('showMasterSp="0"');
		expect(xml).not.toContain('showMasterSp="1"');
		expect(xml).toContain('AIXA_BACKGROUND_PRESERVED_ARTWORK__Swisscom');
		expect(xml).not.toContain('AIXA_BACKGROUND_OVERRIDE_PLACEHOLDER__');
		expect(xml).not.toContain('Long layout authoring instructions');
	});

	it('keeps template graphics enabled when no full-slide picture overlay exists', async () => {
		const xml = await savedSlideXml({
			backgroundSource: 'slide',
			backgroundColor: '#102030',
		});
		expect(xml).toContain('showMasterSp="1"');
		expect(xml).not.toContain('showMasterSp="0"');
	});

	it('preserves native hide-background-graphics behavior without a slide override', async () => {
		const xml = await savedSlideXml({
			backgroundSource: 'inherited',
			showMasterShapes: false,
		});
		expect(xml).toContain('showMasterSp="0"');
	});

	it('honours an explicit hide-background-graphics choice with a slide override', async () => {
		const xml = await savedSlideXml({
			backgroundSource: 'slide',
			backgroundColor: '#102030',
			includeFullSlidePicturePlaceholder: true,
			showMasterShapes: false,
		});
		expect(xml).toContain('showMasterSp="0"');
		expect(xml).not.toContain('AIXA_BACKGROUND_PRESERVED_ARTWORK__');
	});
});

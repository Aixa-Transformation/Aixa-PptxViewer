import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';

describe('authored notes package round-trip', () => {
	it('creates notes master and slide OPC references for a new presentation', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		const slide = createSlide('Blank').build();
		slide.notes = 'A newly authored speaker note';
		data.slides.push(slide);

		const saved = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(saved);
		const requiredParts = [
			'ppt/notesMasters/notesMaster1.xml',
			'ppt/notesMasters/_rels/notesMaster1.xml.rels',
			'ppt/notesSlides/notesSlide1.xml',
			'ppt/notesSlides/_rels/notesSlide1.xml.rels',
		];
		for (const path of requiredParts) {
			expect(zip.file(path), `${path} should exist`).not.toBeNull();
		}

		const presentation = await zip.file('ppt/presentation.xml')!.async('string');
		expect(presentation).toContain('p:notesMasterIdLst');
		const presentationRels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string');
		expect(presentationRels).toContain('relationships/notesMaster');
		expect(presentationRels).toContain('notesMasters/notesMaster1.xml');

		const slideRels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
		expect(slideRels).toContain('relationships/notesSlide');
		const notesRels = await zip.file('ppt/notesSlides/_rels/notesSlide1.xml.rels')!.async('string');
		expect(notesRels).toContain('relationships/notesMaster');
		expect(notesRels).toContain('relationships/slide');

		const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
		expect(contentTypes).toContain('/ppt/notesMasters/notesMaster1.xml');
		expect(contentTypes).toContain('/ppt/notesSlides/notesSlide1.xml');

		const reloader = new PptxHandler();
		const reloaded = await reloader.load(saved.buffer as ArrayBuffer);
		expect(reloaded.slides[0].notes).toBe(slide.notes);
		expect(reloaded.notesMaster?.path).toBe('ppt/notesMasters/notesMaster1.xml');

		const resaved = await reloader.save(reloaded.slides, { notesMaster: reloaded.notesMaster });
		const finalLoader = new PptxHandler();
		const final = await finalLoader.load(resaved.buffer as ArrayBuffer);
		expect(final.slides[0].notes).toBe(slide.notes);
		expect(final.notesMaster).toBeDefined();

		// Some valid generators (including PptxGenJS) place sldIdLst before
		// notesMasterIdLst. Preserve that source-specific ordering on save:
		// normalizing it to the Office-authored order makes desktop PowerPoint
		// report ERROR_FILE_CORRUPT for those decks.
		const generatorZip = await JSZip.loadAsync(saved);
		const generatorPresentation = await generatorZip.file('ppt/presentation.xml')!.async('string');
		const slideList = generatorPresentation.match(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/)?.[0];
		expect(slideList).toBeDefined();
		const generatorOrdered = generatorPresentation
			.replace(slideList!, '')
			.replace('</p:sldMasterIdLst>', `</p:sldMasterIdLst>${slideList!}`);
		generatorZip.file('ppt/presentation.xml', generatorOrdered);
		const generatorBytes = await generatorZip.generateAsync({ type: 'uint8array' });
		const generatorLoader = new PptxHandler();
		const generatorData = await generatorLoader.load(generatorBytes.buffer as ArrayBuffer);
		const generatorResaved = await generatorLoader.save(generatorData.slides, {
			notesMaster: generatorData.notesMaster,
		});
		const generatorResavedZip = await JSZip.loadAsync(generatorResaved);
		const generatorResavedPresentation = await generatorResavedZip
			.file('ppt/presentation.xml')!
			.async('string');
		expect(generatorResavedPresentation.indexOf('<p:sldIdLst')).toBeLessThan(
			generatorResavedPresentation.indexOf('<p:notesMasterIdLst'),
		);
	});
});

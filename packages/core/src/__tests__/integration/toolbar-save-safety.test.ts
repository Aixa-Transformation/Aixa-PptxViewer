import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
	createShapeElement,
	createTextElement,
} from '../../core/builders/sdk/ElementFactory';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import type { InkPptxElement, PptxElement, PptxSlide, ShapePptxElement } from '../../core/types';
import { validatePptx } from '../../core/utils/pptx-validator';

const TINY_PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==';

const CORRUPTION_WARNING_CODES = new Set([
	'CONTENT_TYPE_MISSING_PART',
	'DANGLING_RELATIONSHIP',
	'MISSING_MEDIA',
	'MISSING_THEME',
]);

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

async function expectHealthyPackage(bytes: Uint8Array, expectedSlideCount: number): Promise<void> {
	const buffer = toArrayBuffer(bytes);
	const validation = await validatePptx(buffer);
	const fatalIssues = validation.issues.filter(
		(issue) => issue.severity === 'error' || CORRUPTION_WARNING_CODES.has(issue.code),
	);
	expect(fatalIssues, fatalIssues.map((issue) => `${issue.code}: ${issue.message}`).join('\n')).toStrictEqual(
		[],
	);
	expect(validation.valid).toBeTruthy();

	const zip = await JSZip.loadAsync(bytes);
	expect(Object.values(zip.files).some((entry) => entry.dir)).toBeFalsy();
	const slideParts = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path));
	expect(slideParts).toHaveLength(expectedSlideCount);

	const presentationXml = await zip.file('ppt/presentation.xml')!.async('string');
	const presentationRels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string');
	expect(presentationXml.match(/<p:sldId\b/g) ?? []).toHaveLength(expectedSlideCount);
	expect(
		presentationRels.match(/relationships\/slide["']/g) ?? [],
	).toHaveLength(expectedSlideCount);
}

async function saveAndReload(handler: PptxHandler, slides: PptxSlide[]) {
	const firstBytes = await handler.save(slides);
	await expectHealthyPackage(firstBytes, slides.length);
	const nextHandler = new PptxHandler();
	const nextData = await nextHandler.load(
		toArrayBuffer(firstBytes),
	);

	// A deck that only survives one save is still unsafe: stale relationships
	// and duplicate identifiers often surface on the next editor save.
	const secondBytes = await nextHandler.save(nextData.slides);
	await expectHealthyPackage(secondBytes, nextData.slides.length);
	const finalHandler = new PptxHandler();
	const finalData = await finalHandler.load(toArrayBuffer(secondBytes));
	return { bytes: secondBytes, handler: finalHandler, slides: finalData.slides };
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
			// These are the same model fields used by the Home ribbon's Shape Fill,
			// Shape Outline and Quick Styles controls.
			shape.shapeStyle = {
				...(shape.shapeStyle ?? {}),
				fillMode: 'solid',
				fillColor: '#1565C0',
				strokeFillMode: 'solid',
				strokeColor: '#FFB300',
				strokeWidth: 3,
				shadowColor: '#000000',
				shadowOpacity: 0.35,
				shadowBlur: 6,
				shadowOffsetX: 2,
				shadowOffsetY: 2,
			};
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
		const savedShape = roundTrip.slides[1].elements.find(
			(element) => element.type === 'shape',
		) as (PptxElement & Record<string, any>) | undefined;
		expect(savedShape?.shapeStyle?.fillColor).toBe('#1565C0');
		expect(savedShape?.shapeStyle?.strokeColor).toBe('#FFB300');
		expect(savedShape?.shapeStyle?.strokeWidth).toBeCloseTo(3, 1);

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

	it('survives Insert and Drawing content across two independent saves', async () => {
		const created = await PresentationBuilder.create();
		const groupedText = createTextElement('Grouped text', {
			x: 20,
			y: 20,
			width: 140,
			height: 36,
			bold: true,
		});
		const groupedShape = createShapeElement('diamond', {
			x: 180,
			y: 20,
			width: 50,
			height: 50,
			fill: { type: 'solid', color: '#7B1FA2' },
		});
		const slide = created
			.createSlide('Blank')
			.addText(
				[
					{ text: 'Rich ', style: { bold: true, color: '#1A237E' } },
					{ text: 'text', style: { italic: true, underline: true, color: '#C62828' } },
				],
				{ x: 30, y: 20, width: 260, height: 55, alignment: 'center' },
			)
			.addShape('roundRect', {
				x: 30,
				y: 95,
				width: 170,
				height: 90,
				fill: { type: 'gradient', angle: 45, stops: [
					{ color: '#42A5F5', position: 0 },
					{ color: '#0D47A1', position: 1 },
				] },
				stroke: { color: '#FFB300', width: 3, dash: 'dash' },
				shadow: { color: '#000000', blur: 5, offsetX: 2, offsetY: 3, opacity: 0.3 },
				text: 'Styled shape',
			})
			.addConnector({
				x: 210,
				y: 135,
				width: 120,
				height: 0,
				type: 'bent',
				stroke: { color: '#00897B', width: 2 },
				startArrow: 'oval',
				endArrow: 'triangle',
			})
			.addImage(TINY_PNG, {
				x: 350,
				y: 25,
				width: 100,
				height: 100,
				altText: 'Tiny test image',
				cropLeft: 0.05,
				cropRight: 0.05,
			})
			.addGroup([groupedText, groupedShape], {
				x: 470,
				y: 25,
				width: 250,
				height: 100,
			})
			.addTable(
				{
					rows: [
						{ cells: [
							{ text: 'Metric', style: { bold: true, color: '#FFFFFF' }, fill: { type: 'solid', color: '#1565C0' } },
							{ text: 'Value', style: { bold: true, color: '#FFFFFF' }, fill: { type: 'solid', color: '#1565C0' } },
						] },
						{ cells: [{ text: 'Revenue' }, { text: '$125' }] },
					],
					firstRow: true,
					bandRows: true,
				},
				{ x: 30, y: 220, width: 330, height: 130 },
			)
			.addChart(
				'bar',
				{
					series: [
						{ name: 'Revenue', values: [10, 20, 35], color: '#43A047' },
						{ name: 'Cost', values: [8, 12, 18], color: '#E53935' },
					],
					categories: ['Q1', 'Q2', 'Q3'],
					title: 'Quarterly result',
					hasLegend: true,
				},
				{ x: 390, y: 180, width: 390, height: 260 },
			)
			.build();

		let roundTrip = await saveAndReload(created.handler, [slide]);
		const elementTypes = new Set(roundTrip.slides[0].elements.map((element) => element.type));
		for (const expectedType of ['text', 'shape', 'connector', 'group', 'table', 'chart']) {
			expect(elementTypes.has(expectedType as PptxElement['type']), expectedType).toBeTruthy();
		}
		expect(elementTypes.has('image') || elementTypes.has('picture'), 'image/picture').toBeTruthy();

		const table = roundTrip.slides[0].elements.find((element) => element.type === 'table') as
			| (PptxElement & Record<string, any>)
			| undefined;
		const chart = roundTrip.slides[0].elements.find((element) => element.type === 'chart') as
			| (PptxElement & Record<string, any>)
			| undefined;
		const image = roundTrip.slides[0].elements.find(
			(element) => element.type === 'image' || element.type === 'picture',
		) as
			| (PptxElement & Record<string, any>)
			| undefined;
		expect(table?.tableData).toBeDefined();
		expect(chart?.chartData).toBeDefined();
		expect(image).toBeDefined();
		if (table?.tableData) {
			table.tableData.rows[1].cells[1].text = '$150';
		}
		if (chart?.chartData?.series?.[0]?.values) {
			chart.chartData.series[0].values[2] = 40;
		}
		if (image) {
			image.altText = 'Updated accessible description';
			image.rotation = 15;
			image.opacity = 0.8;
		}
		roundTrip.slides[0].isDirty = true;
		roundTrip = await saveAndReload(roundTrip.handler, roundTrip.slides);

		const savedTable = roundTrip.slides[0].elements.find((element) => element.type === 'table') as
			| (PptxElement & Record<string, any>)
			| undefined;
		const savedImage = roundTrip.slides[0].elements.find(
			(element) => element.type === 'image' || element.type === 'picture',
		) as
			| (PptxElement & Record<string, any>)
			| undefined;
		expect(savedTable?.tableData?.rows[1].cells[1].text).toBe('$150');
		expect(savedImage?.altText).toBe('Updated accessible description');
	});

	it('survives Design, presentation and Review metadata across two saves', async () => {
		const created = await PresentationBuilder.create({
			title: 'Editor save safety',
			creator: 'Aixa Ltd',
			theme: {
				name: 'Aixa test theme',
				colors: { accent1: '#1565C0', accent2: '#F57C00' },
				fonts: { majorFont: 'Aptos Display', minorFont: 'Aptos' },
			},
		});
		const slide = created
			.createSlide('Blank')
			.addText('Review this slide', { x: 50, y: 50, width: 350, height: 60 })
			.setBackground({ type: 'solid', color: '#F5F7FA' })
			.setTransition({ type: 'push', duration: 725, direction: 'l', advanceAfterMs: 4500 })
			.setNotes('Presenter notes survive editing and download.')
			.setHidden(true)
			.setSection('Review', 'section-review')
			.build();
		slide.comments = [{
			id: 'comment-1',
			author: 'Reviewer',
			createdAt: '2026-08-21T10:00:00Z',
			text: 'Please verify this content.',
			x: 80,
			y: 90,
		}];

		let roundTrip = await saveAndReload(created.handler, [slide]);
		expect(roundTrip.slides[0].backgroundColor).toBe('#F5F7FA');
		expect(roundTrip.slides[0].transition?.type).toBe('push');
		expect(roundTrip.slides[0].notes).toContain('Presenter notes');
		expect(roundTrip.slides[0].hidden).toBeTruthy();
		expect(roundTrip.slides[0].comments?.[0].text).toBe('Please verify this content.');

		roundTrip.slides[0].backgroundColor = '#FFF8E1';
		roundTrip.slides[0].transition = { type: 'wipe', durationMs: 900, direction: 'd' };
		roundTrip.slides[0].notes = 'Updated presenter notes.';
		roundTrip.slides[0].comments![0].text = 'Resolved after review.';
		roundTrip.slides[0].hidden = false;
		roundTrip.slides[0].isDirty = true;
		roundTrip = await saveAndReload(roundTrip.handler, roundTrip.slides);

		expect(roundTrip.slides[0].backgroundColor).toBe('#FFF8E1');
		expect(roundTrip.slides[0].transition?.type).toBe('wipe');
		expect(roundTrip.slides[0].notes).toContain('Updated presenter notes');
		expect(roundTrip.slides[0].hidden).toBeFalsy();
		expect(roundTrip.slides[0].comments?.[0].text).toBe('Resolved after review.');
	});

	it('survives Draw pen, highlighter, freeform sketch and eraser operations', async () => {
		const created = await PresentationBuilder.create();
		const baseSlide = created
			.createSlide('Blank')
			.addText('Original slide content', { x: 40, y: 35, width: 300, height: 50 })
			.build();
		let roundTrip = await saveAndReload(created.handler, [baseSlide]);

		const pen: InkPptxElement = {
			id: 'draw-pen',
			type: 'ink',
			x: 100,
			y: 120,
			width: 180,
			height: 80,
			inkPaths: ['M 0 5 L 35 25 L 80 10 L 140 65'],
			inkColors: ['#111111'],
			inkWidths: [3],
			inkOpacities: [1],
			inkPointPressures: [[0.2, 0.5, 0.8, 1]],
			inkTool: 'pen',
		};
		const highlighter: InkPptxElement = {
			id: 'draw-highlighter',
			type: 'ink',
			x: 90,
			y: 235,
			width: 240,
			height: 30,
			inkPaths: ['M 0 15 L 220 15'],
			inkColors: ['#FFF200'],
			inkWidths: [14],
			inkOpacities: [0.4],
			inkTool: 'highlighter',
		};
		const freeform: ShapePptxElement = {
			id: 'draw-freeform',
			type: 'shape',
			x: 380,
			y: 115,
			width: 190,
			height: 150,
			shapeType: 'custom',
			shapeStyle: {
				fillMode: 'none',
				strokeColor: '#C62828',
				strokeWidth: 4,
			},
			customGeometryPaths: [{
				width: 19000,
				height: 15000,
				fillMode: 'none',
				segments: [
					{ type: 'moveTo', pt: { x: 0, y: 3000 } },
					{ type: 'lineTo', pt: { x: 5000, y: 0 } },
					{ type: 'lineTo', pt: { x: 12000, y: 12000 } },
					{ type: 'lineTo', pt: { x: 19000, y: 4000 } },
				],
			}],
		};

		roundTrip.slides[0].elements.push(pen, highlighter, freeform);
		roundTrip.slides[0].isDirty = true;
		roundTrip = await saveAndReload(roundTrip.handler, roundTrip.slides);

		expect(roundTrip.slides[0].elements.filter((element) => element.type === 'contentPart')).toHaveLength(2);
		expect(roundTrip.slides[0].elements.some((element) => element.type === 'shape' && element.shapeType === 'custom')).toBeTruthy();
		const zip = await JSZip.loadAsync(roundTrip.bytes);
		const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		expect(slideXml).toContain('<p:contentPart');
		expect(slideXml).toContain('<a:custGeom>');
		expect(slideXml).not.toContain('drawing/2010/ink');
		expect(slideXml).not.toMatch(/<a:prstGeom\b[^>]*\bprst="custom"/u);

		// Simulate the Draw eraser: remove one complete ink element and verify
		// that its relationship/InkML part disappears without harming the deck.
		const firstInk = roundTrip.slides[0].elements.find((element) => element.type === 'contentPart');
		expect(firstInk).toBeDefined();
		roundTrip.slides[0].elements = roundTrip.slides[0].elements.filter(
			(element) => element.id !== firstInk?.id,
		);
		roundTrip.slides[0].isDirty = true;
		roundTrip = await saveAndReload(roundTrip.handler, roundTrip.slides);
		expect(roundTrip.slides[0].elements.filter((element) => element.type === 'contentPart')).toHaveLength(1);
		expect(firstTextElement(roundTrip.slides[0]).text).toBe('Original slide content');
	});
});

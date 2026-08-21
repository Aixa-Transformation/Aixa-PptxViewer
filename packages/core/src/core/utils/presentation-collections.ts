import type { IPptxXmlLookupService } from '../services';
import type { PptxCustomShow, PptxSection, XmlObject } from '../types';
import { cloneXmlObject } from './clone-utils';

const SECTION_EXTENSION_URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';

/**
 * Slide-reference remapping produced by the presentation-slides reconciler.
 *
 * When slides are reordered, added, or removed the reconciler may reassign a
 * slide relationship id (`r:id`) and mints fresh numeric slide ids for new
 * slides. Custom shows reference slides by relationship id and sections by
 * numeric slide id, so those references must be rewritten to the current
 * values (and references to removed slides dropped) before serialization.
 */
export interface PptxSlideReferenceRemap {
	/** Old presentation-rels slide rId -> current rId, for surviving slides. */
	rIdByOldRId: Map<string, string>;
	/** Old numeric slide id -> current numeric slide id, for surviving slides. */
	sldIdByOldSldId: Map<string, string>;
	/** Old slide rIds whose slide was removed (references must be dropped). */
	removedRIds: Set<string>;
	/** Old numeric slide ids whose slide was removed (references must be dropped). */
	removedSldIds: Set<string>;
	/** Numeric slide ids in their final presentation order. */
	orderedSldIds: string[];
	/** Numeric ids minted for slides that did not exist in the source deck. */
	newSldIds: Set<string>;
	/** True when any reference changed or any slide was removed. */
	changed: boolean;
}

function remapReferenceList(
	references: readonly string[],
	mapping: Map<string, string>,
	removed: Set<string>,
): string[] {
	const result: string[] = [];
	for (const reference of references) {
		const mapped = mapping.get(reference);
		if (mapped !== undefined) {
			result.push(mapped);
			continue;
		}
		if (removed.has(reference)) {
			continue;
		}
		// Unknown reference (never a tracked slide): leave it untouched.
		result.push(reference);
	}
	return result;
}

function localName(key: string): string {
	return key.split(':').pop() ?? key;
}

function childKey(parent: XmlObject, name: string): string | undefined {
	return Object.keys(parent).find((key) => localName(key) === name);
}

function attributeKey(node: XmlObject, name: string): string | undefined {
	return Object.keys(node).find((key) => {
		if (!key.startsWith('@_')) {
			return false;
		}
		return localName(key.slice(2)) === name;
	});
}

function setAttribute(node: XmlObject, name: string, value: string, fallbackKey: string): void {
	node[attributeKey(node, name) ?? fallbackKey] = value;
}

function relationshipId(node: XmlObject): string {
	const key = attributeKey(node, 'id');
	return String((key && node[key]) || '').trim();
}

export function parseCustomShows(
	presentationData: XmlObject | null,
	lookup: IPptxXmlLookupService,
): PptxCustomShow[] | undefined {
	if (!presentationData) {
		return undefined;
	}
	const presentation = lookup.getChildByLocalName(presentationData, 'presentation');
	const list = lookup.getChildByLocalName(presentation, 'custShowLst');
	const shows = lookup.getChildrenArrayByLocalName(list, 'custShow');
	if (shows.length === 0) {
		return undefined;
	}

	return shows.map((show) => {
		const slideList = lookup.getChildByLocalName(show, 'sldLst');
		return {
			name: String(show[attributeKey(show, 'name') ?? ''] || ''),
			id: String(show[attributeKey(show, 'id') ?? ''] || ''),
			slideRIds: lookup
				.getChildrenArrayByLocalName(slideList, 'sld')
				.map(relationshipId)
				.filter(Boolean),
			rawXml: cloneXmlObject(show),
		};
	});
}

function replaceChildren(
	parent: XmlObject,
	name: string,
	value: XmlObject | XmlObject[],
	fallbackKey: string,
): void {
	parent[childKey(parent, name) ?? fallbackKey] = value;
}

function updateCustomShow(show: PptxCustomShow, existing?: XmlObject): XmlObject {
	const node = cloneXmlObject(show.rawXml) ?? cloneXmlObject(existing) ?? {};
	setAttribute(node, 'name', show.name, '@_name');
	setAttribute(node, 'id', String(show.id), '@_id');
	const oldListKey = childKey(node, 'sldLst');
	const oldList = oldListKey ? (node[oldListKey] as XmlObject) : undefined;
	const slideList = cloneXmlObject(oldList) ?? {};
	const oldSlidesKey = childKey(slideList, 'sld');
	const oldSlides = oldSlidesKey ? slideList[oldSlidesKey] : undefined;
	const candidates = (Array.isArray(oldSlides) ? oldSlides : oldSlides ? [oldSlides] : []).filter(
		(entry): entry is XmlObject => typeof entry === 'object' && entry !== null,
	);
	const slides = show.slideRIds.map((rId) => {
		const prior = candidates.find((entry) => relationshipId(entry) === rId);
		const slide = cloneXmlObject(prior) ?? {};
		setAttribute(slide, 'id', rId, '@_r:id');
		return slide;
	});
	replaceChildren(slideList, 'sld', slides, 'p:sld');
	replaceChildren(node, 'sldLst', slideList, 'p:sldLst');
	return node;
}

export function applyCustomShows(
	presentation: XmlObject,
	shows: PptxCustomShow[] | undefined,
	lookup: IPptxXmlLookupService,
	remap?: PptxSlideReferenceRemap,
): void {
	if (shows === undefined) {
		return;
	}
	const key = childKey(presentation, 'custShowLst');
	if (shows.length === 0) {
		if (key) {
			delete presentation[key];
		}
		return;
	}
	const oldList = key ? (presentation[key] as XmlObject) : undefined;
	const list = cloneXmlObject(oldList) ?? {};
	const oldShows = lookup.getChildrenArrayByLocalName(oldList, 'custShow');
	const updated = shows.map((show) => {
		const effective =
			remap && remap.changed
				? {
						...show,
						slideRIds: remapReferenceList(show.slideRIds, remap.rIdByOldRId, remap.removedRIds),
					}
				: show;
		return updateCustomShow(
			effective,
			oldShows.find(
				(node) =>
					String(node[attributeKey(node, 'id') ?? '']) === show.id ||
					String(node[attributeKey(node, 'name') ?? '']) === show.name,
			),
		);
	});
	replaceChildren(list, 'custShow', updated, 'p:custShow');
	presentation[key ?? 'p:custShowLst'] = list;
}

interface SectionLocation {
	parent: XmlObject;
	key: string;
	list: XmlObject;
}

function findSectionList(
	presentation: XmlObject,
	lookup: IPptxXmlLookupService,
): SectionLocation | undefined {
	const direct = childKey(presentation, 'sectionLst');
	if (direct) {
		return { parent: presentation, key: direct, list: presentation[direct] as XmlObject };
	}
	const extList = lookup.getChildByLocalName(presentation, 'extLst');
	for (const ext of lookup.getChildrenArrayByLocalName(extList, 'ext')) {
		const key = childKey(ext, 'sectionLst');
		if (key) {
			return { parent: ext, key, list: ext[key] as XmlObject };
		}
	}
	return undefined;
}

function updateSection(section: PptxSection): XmlObject {
	const node = cloneXmlObject(section.rawXml) ?? {};
	setAttribute(node, 'name', section.name, '@_name');
	setAttribute(node, 'id', section.id, '@_id');
	const oldSlideListKey = childKey(node, 'sldIdLst');
	const oldSlideList = oldSlideListKey ? (node[oldSlideListKey] as XmlObject) : undefined;
	const slideList = cloneXmlObject(oldSlideList) ?? {};
	replaceChildren(
		slideList,
		'sldId',
		section.slideIds.map((id) => ({ '@_id': id })),
		'p14:sldId',
	);
	replaceChildren(node, 'sldIdLst', slideList, 'p14:sldIdLst');
	const sectionPrKey = childKey(node, 'sectionPr');
	if (section.collapsed !== undefined || section.color !== undefined || sectionPrKey) {
		const sectionPr =
			cloneXmlObject(sectionPrKey ? (node[sectionPrKey] as XmlObject) : undefined) ?? {};
		if (section.collapsed === undefined) {
			delete sectionPr[attributeKey(sectionPr, 'collapsed') ?? ''];
		} else {
			setAttribute(sectionPr, 'collapsed', section.collapsed ? '1' : '0', '@_collapsed');
		}
		if (section.color === undefined) {
			delete sectionPr[attributeKey(sectionPr, 'clr') ?? ''];
		} else {
			setAttribute(sectionPr, 'clr', section.color.replace('#', ''), '@_clr');
		}
		node[sectionPrKey ?? 'p15:sectionPr'] = sectionPr;
	}
	return node;
}

export function applySections(
	presentation: XmlObject,
	sections: PptxSection[] | undefined,
	lookup: IPptxXmlLookupService,
	remap?: PptxSlideReferenceRemap,
): void {
	if (sections === undefined) {
		return;
	}
	let location = findSectionList(presentation, lookup);
	if (sections.length === 0) {
		if (location) {
			delete location.parent[location.key];
		}
		return;
	}
	if (!location) {
		const extListKey = childKey(presentation, 'extLst') ?? 'p:extLst';
		const extList = cloneXmlObject(presentation[extListKey] as XmlObject | undefined) ?? {};
		const extKey = childKey(extList, 'ext') ?? 'p:ext';
		const entries = lookup.getChildrenArrayByLocalName(extList, 'ext');
		const ext: XmlObject = { '@_uri': SECTION_EXTENSION_URI, 'p14:sectionLst': {} };
		extList[extKey] = [...entries, ext];
		presentation[extListKey] = extList;
		location = { parent: ext, key: 'p14:sectionLst', list: ext['p14:sectionLst'] as XmlObject };
	}
	const list = cloneXmlObject(location.list) ?? {};
	let effectiveSections =
		remap && remap.changed
			? sections.map((section) => ({
					...section,
					slideIds: remapReferenceList(
						section.slideIds,
						remap.sldIdByOldSldId,
						remap.removedSldIds,
					),
				}))
			: sections;
	if (remap && remap.newSldIds.size > 0) {
		const ownerBySlideId = new Map<string, number>();
		for (let sectionIndex = 0; sectionIndex < effectiveSections.length; sectionIndex++) {
			for (const slideId of effectiveSections[sectionIndex].slideIds) {
				ownerBySlideId.set(slideId, sectionIndex);
			}
		}

		// PowerPoint expects every slide in a sectioned deck to occur in exactly
		// one section. A newly inserted slide has no source section reference, so
		// inherit the preceding slide's section (or the following section when it
		// was inserted at the beginning).
		for (let index = 0; index < remap.orderedSldIds.length; index++) {
			const slideId = remap.orderedSldIds[index];
			if (!remap.newSldIds.has(slideId)) {
				continue;
			}
			let owner: number | undefined;
			for (let cursor = index - 1; cursor >= 0 && owner === undefined; cursor--) {
				owner = ownerBySlideId.get(remap.orderedSldIds[cursor]);
			}
			for (let cursor = index + 1; cursor < remap.orderedSldIds.length && owner === undefined; cursor++) {
				owner = ownerBySlideId.get(remap.orderedSldIds[cursor]);
			}
			if (owner !== undefined) {
				ownerBySlideId.set(slideId, owner);
			}
		}

		effectiveSections = effectiveSections.map((section, sectionIndex) => ({
			...section,
			slideIds: remap.orderedSldIds.filter((slideId) => ownerBySlideId.get(slideId) === sectionIndex),
		}));
	}
	replaceChildren(list, 'section', effectiveSections.map(updateSection), 'p14:section');
	location.parent[location.key] = list;
}
// Modified by Aixa Ltd from the original ChristopherVR/pptx-viewer project.

import { EMU_PER_PX } from '../../constants';
import { XmlObject, PptxElement } from '../../types';
import { cloneXmlObject } from '../../utils/clone-utils';
import { canonicalPlaceholderType } from '../../utils/placeholder-validation';
import { createShapeIdAllocator, elementShapeIds } from '../../utils/shape-ids';
import { xmlPath } from '../../utils/xml-access';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeTextEditing';
import type { PlaceholderInfo } from './PptxHandlerRuntimeTypes';
import { isEmptyGeneratedPlaceholder, markGeneratedPlaceholder } from './transient-placeholder';

/**
 * Layout-switching helpers for the PptxHandlerRuntime mixin chain.
 *
 * Provides methods that map slide elements to a new layout's placeholders
 * by type, reposition matched placeholders, preserve unmatched content, and
 * inject empty placeholders that exist only in the target layout.
 */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	// ── Placeholder info extraction ─────────────────────────────────────

	/**
	 * Read placeholder info from a `p:nvPr` XML node.
	 *
	 * This is a local helper that mirrors the logic in
	 * `PptxHandlerRuntimeElementParsing.extractPlaceholderInfo` — we
	 * duplicate it here because that mixin sits higher in the chain and
	 * is not yet available at this level.
	 */
	private readPlaceholderInfoFromNvPr(nvPr: XmlObject | undefined): PlaceholderInfo | null {
		if (!nvPr) {
			return null;
		}
		const ph = nvPr['p:ph'] as XmlObject | undefined;
		if (!ph) {
			return null;
		}

		const idx = ph['@_idx'];
		const type = ph['@_type'];
		const sz = ph['@_sz'];

		return {
			idx: idx !== undefined ? String(idx) : undefined,
			type: type !== undefined ? String(type).toLowerCase() : undefined,
			sz: sz !== undefined ? String(sz).toLowerCase() : undefined,
		};
	}

	/**
	 * Extract placeholder info from a parsed slide element's rawXml.
	 * Works for shapes (`p:nvSpPr`), pictures (`p:nvPicPr`), and
	 * graphic frames (`p:nvGraphicFramePr`).
	 */
	protected getElementPlaceholderInfo(element: PptxElement): PlaceholderInfo | null {
		const raw = element.rawXml;
		if (!raw) {
			return null;
		}

		const nvPr =
			xmlPath(raw, 'p:nvSpPr', 'p:nvPr') ??
			xmlPath(raw, 'p:nvPicPr', 'p:nvPr') ??
			xmlPath(raw, 'p:nvGraphicFramePr', 'p:nvPr');

		return this.readPlaceholderInfoFromNvPr(nvPr);
	}

	// ── Layout placeholder extraction ───────────────────────────────────

	/**
	 * Extract all placeholders from a layout's `p:spTree`, returning
	 * their placeholder info and their transform (position/size in EMU).
	 */
	protected extractLayoutPlaceholders(
		layoutXml: XmlObject,
		layoutPath?: string,
	): Array<{
		phInfo: PlaceholderInfo;
		xEmu: number;
		yEmu: number;
		cxEmu: number;
		cyEmu: number;
		shapeXml: XmlObject;
	}> {
		const spTree = xmlPath(layoutXml, 'p:sldLayout', 'p:cSld', 'p:spTree');
		if (!spTree) {
			return [];
		}

		const result: Array<{
			phInfo: PlaceholderInfo;
			xEmu: number;
			yEmu: number;
			cxEmu: number;
			cyEmu: number;
			shapeXml: XmlObject;
		}> = [];

		// Most layout placeholders are p:sp nodes, but imported and third-party
		// decks can also persist picture or graphic-frame placeholders directly.
		const shapes = [
			...(this.ensureArray(spTree['p:sp']) as XmlObject[]),
			...(this.ensureArray(spTree['p:pic']) as XmlObject[]),
			...(this.ensureArray(spTree['p:graphicFrame']) as XmlObject[]),
		];
		for (const shape of shapes) {
			const nvPr =
				xmlPath(shape, 'p:nvSpPr', 'p:nvPr') ??
				xmlPath(shape, 'p:nvPicPr', 'p:nvPr') ??
				xmlPath(shape, 'p:nvGraphicFramePr', 'p:nvPr');
			const phInfo = this.readPlaceholderInfoFromNvPr(nvPr);
			if (!phInfo) {
				continue;
			}

			// A layout placeholder commonly omits a:xfrm and inherits it from
			// the matching master placeholder. Resolve that inheritance exactly
			// as normal slide parsing does, otherwise the destination is 0x0 and
			// switching appears to leave only the layout background visible.
			const masterPath = layoutPath ? this.resolveMasterPathForLayout(layoutPath) : undefined;
			const masterContext = masterPath
				? this.findPlaceholderInShapeTree(
						xmlPath(this.masterXmlMap.get(masterPath), 'p:sldMaster', 'p:cSld', 'p:spTree'),
						phInfo,
					)
				: undefined;
			const inheritedShape = masterContext?.shape ?? masterContext?.picture;
			const resolvedShape = inheritedShape
				? (this.mergeXmlObjects(inheritedShape, shape) ?? shape)
				: shape;

			// Get transform
			const spPr = resolvedShape['p:spPr'] as XmlObject | undefined;
			const xfrm = spPr?.['a:xfrm'] as XmlObject | undefined;
			const off = xfrm?.['a:off'] as XmlObject | undefined;
			const ext = xfrm?.['a:ext'] as XmlObject | undefined;

			const xEmu = off ? Number(off['@_x'] || 0) : 0;
			const yEmu = off ? Number(off['@_y'] || 0) : 0;
			const cxEmu = ext ? Number(ext['@_cx'] || 0) : 0;
			const cyEmu = ext ? Number(ext['@_cy'] || 0) : 0;

			result.push({ phInfo, xEmu, yEmu, cxEmu, cyEmu, shapeXml: resolvedShape });
		}

		return result;
	}

	// ── Placeholder matching key ────────────────────────────────────────

	/**
	 * Build a matching key for a placeholder. Placeholders match primarily
	 * by type. When both have an idx, the idx must also match.
	 */
	protected buildPlaceholderMatchKey(phInfo: PlaceholderInfo): string {
		// Normalise missing type to "body" (the OOXML default)
		const type = phInfo.type || 'body';
		if (phInfo.idx !== undefined) {
			return `${type}:${phInfo.idx}`;
		}
		return type;
	}

	private placeholderRole(phInfo: PlaceholderInfo): 'title' | 'content' | string {
		const type = phInfo.type || 'body';
		if (type === 'title' || type === 'ctrtitle') {
			return 'title';
		}
		if (
			type === 'body' ||
			type === 'obj' ||
			type === 'subtitle' ||
			type === 'pic' ||
			type === 'chart' ||
			type === 'tbl' ||
			type === 'media'
		) {
			return 'content';
		}
		return type;
	}

	private preferredPlaceholderTypes(element: PptxElement): string[] {
		switch (element.type) {
			case 'image':
				return ['pic', 'obj'];
			case 'chart':
				return ['chart', 'obj'];
			case 'table':
				return ['tbl', 'obj'];
			case 'video':
			case 'audio':
				return ['media', 'obj'];
			case 'text':
				return ['body', 'subtitle', 'obj'];
			default:
				return ['obj', 'body'];
		}
	}

	private placeholderMatchScore(
		element: PptxElement,
		source: PlaceholderInfo,
		target: PlaceholderInfo,
	): number {
		const sourceType = source.type || 'body';
		const targetType = target.type || 'body';
		const sourceRole = this.placeholderRole(source);
		const targetRole = this.placeholderRole(target);
		if (sourceRole !== targetRole) return -1;

		let score = 0;
		if (source.idx !== undefined && target.idx === source.idx) score += 100;
		if (targetType === sourceType) score += 50;
		const preferredIndex = this.preferredPlaceholderTypes(element).indexOf(targetType);
		if (preferredIndex >= 0) score += 30 - preferredIndex * 5;
		// An object placeholder accepts every content kind, but it must rank below
		// the kind-specific picture/chart/table/media destination when both exist.
		if (targetType === 'obj') score += 10;
		return score;
	}

	// ── Core layout switching logic ─────────────────────────────────────

	/**
	 * Re-map slide elements to a new layout's placeholders.
	 *
	 * - Placeholder elements whose type matches a new-layout placeholder
	 *   get their position/size updated to the new layout's values.
	 * - Placeholder elements with no match in the new layout are preserved.
	 * - New-layout placeholders with no matching slide element produce
	 *   empty text elements that are appended to the slide.
	 * - Non-placeholder elements are left untouched.
	 *
	 * @returns The updated elements array.
	 */
	protected remapElementsToNewLayout(
		elements: PptxElement[],
		newLayoutXml: XmlObject,
		newLayoutPath: string,
	): PptxElement[] {
		const layoutPlaceholders = this.extractLayoutPlaceholders(newLayoutXml, newLayoutPath);

		// Keep placeholders as an array. A layout may legally contain multiple
		// placeholders with the same type (and occasionally an omitted idx), so
		// a Map would silently discard all but the last one.
		const targetPlaceholders: Array<{
			phInfo: PlaceholderInfo;
			xEmu: number;
			yEmu: number;
			cxEmu: number;
			cyEmu: number;
			shapeXml: XmlObject;
			matched: boolean;
		}> = [];
		for (const lp of layoutPlaceholders) {
			targetPlaceholders.push({ ...lp, matched: false });
		}

		const resultElements: PptxElement[] = [];

		for (const element of elements) {
			const metadata = element as PptxElement & {
				_layoutSwitchOriginal?: PptxElement;
			};
			// Placeholders generated for a previous layout are transient while they
			// stay empty; keeping them would duplicate boxes on the next switch. Once
			// the user has typed into one it is ordinary slide content.
			if (isEmptyGeneratedPlaceholder(element)) continue;
			// Always remap from the canonical pre-switch element. This makes A -> B
			// -> A reversible and prevents transforms/placeholder identities from
			// accumulating across repeated layout selections.
			const original = metadata._layoutSwitchOriginal ?? element;
			const originalPhInfo = this.getElementPlaceholderInfo(original);
			const baseElement = {
				// Keep all current content and styling edits. Only placeholder geometry
				// and identity come from the canonical pre-switch element.
				...element,
				...(originalPhInfo
					? { x: original.x, y: original.y, width: original.width, height: original.height }
					: {}),
				rawXml: element.rawXml ? cloneXmlObject(element.rawXml) : element.rawXml,
			} as PptxElement & { _layoutSwitchOriginal?: PptxElement };
			baseElement._layoutSwitchOriginal = metadata._layoutSwitchOriginal ?? {
				...element,
				rawXml: element.rawXml ? cloneXmlObject(element.rawXml) : element.rawXml,
			};
			const phInfo = originalPhInfo ?? this.getElementPlaceholderInfo(baseElement);

			if (!phInfo) {
				// Non-placeholder element: keep as-is
				resultElements.push(baseElement);
				continue;
			}

			// Rank every compatible destination. This is important for layouts that
			// combine several content placeholders: a picture must choose pic before
			// body, a chart must choose chart, and repeated placeholders still use idx.
			let resolvedLayoutPh: (typeof targetPlaceholders)[number] | undefined;
			let bestScore = -1;
			for (const candidate of targetPlaceholders) {
				if (candidate.matched) continue;
				const score = this.placeholderMatchScore(element, phInfo, candidate.phInfo);
				if (score > bestScore) {
					bestScore = score;
					resolvedLayoutPh = candidate;
				}
			}

			if (resolvedLayoutPh) {
				// Matched: update position and size from new layout
				resolvedLayoutPh.matched = true;

				const updatedElement = {
					...baseElement,
					rawXml: baseElement.rawXml ? cloneXmlObject(baseElement.rawXml) : baseElement.rawXml,
				} as PptxElement;
				if (resolvedLayoutPh.cxEmu > 0 && resolvedLayoutPh.cyEmu > 0) {
					updatedElement.x = Math.round(resolvedLayoutPh.xEmu / EMU_PER_PX);
					updatedElement.y = Math.round(resolvedLayoutPh.yEmu / EMU_PER_PX);
					updatedElement.width = Math.round(resolvedLayoutPh.cxEmu / EMU_PER_PX);
					updatedElement.height = Math.round(resolvedLayoutPh.cyEmu / EMU_PER_PX);
				}

				// Update the element's rawXml transform to match
				if (updatedElement.rawXml && resolvedLayoutPh.cxEmu > 0 && resolvedLayoutPh.cyEmu > 0) {
					this.updateElementRawXmlTransform(
						updatedElement.rawXml,
						resolvedLayoutPh.xEmu,
						resolvedLayoutPh.yEmu,
						resolvedLayoutPh.cxEmu,
						resolvedLayoutPh.cyEmu,
					);
				}
				if (updatedElement.rawXml) {
					this.updateElementRawXmlPlaceholder(updatedElement.rawXml, resolvedLayoutPh.phInfo);
				}

				resultElements.push(updatedElement);
			} else {
				// Never discard user content merely because the selected layout has
				// no corresponding placeholder. PowerPoint keeps such content as a
				// free-standing element; preserving it also keeps pictures and text
				// from disappearing during layout changes.
				resultElements.push(baseElement);
			}
		}

		// Add empty placeholders from the new layout that were not matched
		const allocateShapeId = createShapeIdAllocator(elementShapeIds(resultElements));
		for (const lp of targetPlaceholders) {
			if (lp.matched) {
				continue;
			}
			// Skip footers, date-time, and slide number placeholders -- they
			// are rendered from the layout/master and don't need slide-level
			// elements.
			const skipTypes = new Set(['dt', 'ftr', 'sldnum', 'hdr']);
			if (lp.phInfo.type && skipTypes.has(lp.phInfo.type)) {
				continue;
			}

			// Create an empty text element for this placeholder
			const emptyElement = this.createEmptyPlaceholderElement(
				lp.phInfo,
				lp.xEmu,
				lp.yEmu,
				lp.cxEmu,
				lp.cyEmu,
				newLayoutPath,
				lp.shapeXml,
				allocateShapeId(),
			);
			if (emptyElement) {
				resultElements.push(emptyElement);
			}
		}

		return resultElements;
	}

	private updateElementRawXmlPlaceholder(rawXml: XmlObject, phInfo: PlaceholderInfo): void {
		const nvPr =
			xmlPath(rawXml, 'p:nvSpPr', 'p:nvPr') ??
			xmlPath(rawXml, 'p:nvPicPr', 'p:nvPr') ??
			xmlPath(rawXml, 'p:nvGraphicFramePr', 'p:nvPr');
		if (!nvPr) {
			return;
		}
		const ph = (nvPr['p:ph'] as XmlObject | undefined) ?? {};
		nvPr['p:ph'] = ph;
		if (phInfo.type) {
			ph['@_type'] = canonicalPlaceholderType(phInfo.type);
		} else {
			delete ph['@_type'];
		}
		if (phInfo.idx !== undefined) {
			ph['@_idx'] = phInfo.idx;
		} else {
			delete ph['@_idx'];
		}
	}

	// ── rawXml transform update ─────────────────────────────────────────

	/**
	 * Update the transform (`a:xfrm`) inside an element's rawXml to
	 * reflect new position and size values in EMU.
	 */
	protected updateElementRawXmlTransform(
		rawXml: XmlObject,
		xEmu: number,
		yEmu: number,
		cxEmu: number,
		cyEmu: number,
	): void {
		// Find spPr in the appropriate container
		const spPr = rawXml['p:spPr'] as XmlObject | undefined;
		if (!spPr) {
			return;
		}

		let xfrm = spPr['a:xfrm'] as XmlObject | undefined;
		if (!xfrm) {
			xfrm = {};
			spPr['a:xfrm'] = xfrm;
		}

		let off = xfrm['a:off'] as XmlObject | undefined;
		if (!off) {
			off = {};
			xfrm['a:off'] = off;
		}
		off['@_x'] = String(xEmu);
		off['@_y'] = String(yEmu);

		let ext = xfrm['a:ext'] as XmlObject | undefined;
		if (!ext) {
			ext = {};
			xfrm['a:ext'] = ext;
		}
		ext['@_cx'] = String(cxEmu);
		ext['@_cy'] = String(cyEmu);
	}

	// ── Empty placeholder creation ──────────────────────────────────────

	/**
	 * Read the prompt text a layout/master placeholder shows when it is empty
	 * ("Click to add title"). Duplicated from `extractPlaceholderDefaultsFromShape`
	 * because that mixin sits higher in the chain than this one.
	 */
	private readPlaceholderPromptText(shapeXml: XmlObject | undefined): string | undefined {
		const txBody = xmlPath(shapeXml, 'p:txBody');
		if (!txBody) {
			return undefined;
		}
		const parts: string[] = [];
		for (const paragraph of this.ensureArray(txBody['a:p']) as XmlObject[]) {
			if (!paragraph) {
				continue;
			}
			const runs = [
				...(this.ensureArray(paragraph['a:r']) as XmlObject[]),
				...(this.ensureArray(paragraph['a:fld']) as XmlObject[]),
			];
			for (const run of runs) {
				if (run?.['a:t'] !== undefined) {
					parts.push(String(run['a:t']));
				}
			}
			if (paragraph['a:t'] !== undefined) {
				parts.push(String(paragraph['a:t']));
			}
		}
		const promptText = parts.join('').trim();
		return promptText.length > 0 ? promptText : undefined;
	}

	private defaultPlaceholderPromptText(type: string | undefined): string {
		switch (type) {
			case 'title':
			case 'ctrtitle':
				return 'Click to add title';
			case 'subtitle':
				return 'Click to add subtitle';
			case 'pic':
				return 'Click to add picture';
			case 'chart':
				return 'Click to add chart';
			case 'tbl':
				return 'Click to add table';
			default:
				return 'Click to add text';
		}
	}

	/**
	 * Create a minimal text element representing an empty placeholder
	 * from the new layout. The element has the correct position/size and
	 * a `rawXml` with a `p:ph` reference so that the save pipeline
	 * preserves the placeholder binding.
	 */
	/**
	 * Read the level-1 run defaults a layout/master placeholder authors, so a
	 * generated placeholder starts at the layout's own font instead of the
	 * generic viewer default (and never shifts afterwards).
	 */
	private readPlaceholderRunDefaults(shapeXml: XmlObject | undefined): {
		fontSize?: number;
		fontFamily?: string;
	} {
		const defRPr =
			xmlPath(shapeXml, 'p:txBody', 'a:lstStyle', 'a:lvl1pPr', 'a:defRPr') ??
			xmlPath(shapeXml, 'p:txBody', 'a:p', 'a:endParaRPr') ??
			xmlPath(shapeXml, 'p:txBody', 'a:p', 'a:r', 'a:rPr');
		if (!defRPr) {
			return {};
		}
		const result: { fontSize?: number; fontFamily?: string } = {};
		const hundredths = Number(defRPr['@_sz']);
		if (Number.isFinite(hundredths) && hundredths > 0) {
			result.fontSize = (hundredths / 100) * (96 / 72);
		}
		const typeface = (xmlPath(defRPr, 'a:latin') as XmlObject | undefined)?.['@_typeface'];
		// `+mj-lt` / `+mn-lt` are theme references the renderer resolves itself.
		if (typeof typeface === 'string' && typeface.length > 0 && !typeface.startsWith('+')) {
			result.fontFamily = typeface;
		}
		return result;
	}

	protected createEmptyPlaceholderElement(
		phInfo: PlaceholderInfo,
		xEmu: number,
		yEmu: number,
		cxEmu: number,
		cyEmu: number,
		_layoutPath: string,
		layoutShapeXml?: XmlObject,
		shapeId = '2',
	): PptxElement | null {
		if (cxEmu <= 0 || cyEmu <= 0) {
			return null;
		}

		const phNode: XmlObject = {};
		if (phInfo.type) {
			phNode['@_type'] = canonicalPlaceholderType(phInfo.type);
		}
		if (phInfo.idx !== undefined) {
			phNode['@_idx'] = phInfo.idx;
		}
		const runDefaults = this.readPlaceholderRunDefaults(layoutShapeXml);

		const rawXml: XmlObject = {
			'p:nvSpPr': {
				'p:cNvPr': {
					'@_id': shapeId,
					'@_name': `Placeholder ${phInfo.type || 'content'}`,
				},
				'p:cNvSpPr': {
					'a:spLocks': { '@_noGrp': '1' },
				},
				'p:nvPr': {
					'p:ph': phNode,
				},
			},
			'p:spPr': {
				'a:xfrm': {
					'a:off': { '@_x': String(xEmu), '@_y': String(yEmu) },
					'a:ext': { '@_cx': String(cxEmu), '@_cy': String(cyEmu) },
				},
			},
			'p:txBody': {
				'a:bodyPr': {},
				'a:lstStyle': {},
				'a:p': { 'a:endParaRPr': { '@_lang': 'en-US' } },
			},
		};

		const element: PptxElement = {
			type: 'text' as const,
			id: `ph-${phInfo.type || 'content'}-${phInfo.idx || '0'}-${Date.now()}`,
			shapeId,
			x: Math.round(xEmu / EMU_PER_PX),
			y: Math.round(yEmu / EMU_PER_PX),
			width: Math.round(cxEmu / EMU_PER_PX),
			height: Math.round(cyEmu / EMU_PER_PX),
			text: '',
			// Without a prompt an empty placeholder renders nothing at all, so the
			// inserted layout looks like it has no text areas.
			promptText:
				this.readPlaceholderPromptText(layoutShapeXml) ??
				this.defaultPlaceholderPromptText(phInfo.type),
			...(Object.keys(runDefaults).length > 0 ? { textStyle: runDefaults } : {}),
			rawXml,
		};
		// Lets consumers resolve the theme's major vs minor font for a box that
		// carries no authored run properties yet.
		(element as PptxElement & { placeholderType?: string }).placeholderType = phInfo.type ?? 'body';
		return markGeneratedPlaceholder(element);
	}

	/**
	 * Build empty placeholder elements for every placeholder a layout declares
	 * that the slide itself does not contain.
	 *
	 * A `.pptx` only stores the placeholders that were actually filled in, so a
	 * freshly authored slide has nothing to click into even though its layout
	 * defines a title and body. PowerPoint shows those as prompt boxes; this
	 * reproduces them for the load path the same way a layout switch does.
	 */
	protected async buildMissingLayoutPlaceholders(
		slidePath: string,
		slideElements: PptxElement[],
		layoutPath: string,
	): Promise<PptxElement[]> {
		let layoutXml = this.layoutXmlMap.get(layoutPath);
		if (!layoutXml) {
			try {
				const layoutXmlStr = await this.zip.file(layoutPath)?.async('string');
				if (!layoutXmlStr) {
					return [];
				}
				layoutXml = this.parser.parse(layoutXmlStr) as XmlObject;
				this.layoutXmlMap.set(layoutPath, layoutXml);
			} catch {
				return [];
			}
		}

		const taken = new Set<string>();
		for (const element of slideElements) {
			const phInfo = this.getElementPlaceholderInfo(element);
			if (phInfo) {
				taken.add(this.buildPlaceholderMatchKey(phInfo));
			}
		}

		const skipTypes = new Set(['dt', 'ftr', 'sldnum', 'hdr']);
		const generated: PptxElement[] = [];
		const allocateShapeId = createShapeIdAllocator(elementShapeIds(slideElements));
		for (const lp of this.extractLayoutPlaceholders(layoutXml, layoutPath)) {
			if (lp.phInfo.type && skipTypes.has(lp.phInfo.type)) {
				continue;
			}
			const key = this.buildPlaceholderMatchKey(lp.phInfo);
			if (taken.has(key)) {
				continue;
			}
			taken.add(key);
			const element = this.createEmptyPlaceholderElement(
				lp.phInfo,
				lp.xEmu,
				lp.yEmu,
				lp.cxEmu,
				lp.cyEmu,
				layoutPath,
				lp.shapeXml,
				allocateShapeId(),
			);
			if (element) {
				element.id = `${slidePath}-ph-${key}`;
				generated.push(element);
			}
		}
		return generated;
	}
}

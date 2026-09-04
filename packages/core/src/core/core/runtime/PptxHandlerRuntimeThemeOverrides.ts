import { XmlObject, PptxSlide } from '../../types';
import { safeResolveZipPath } from '../../utils/safe-path';
import type { PptxThemeFormatScheme } from '../../types';
import {
	PptxHandlerRuntime as PptxHandlerRuntimeBase,
	extractFillStyleListChildOrder,
} from './PptxHandlerRuntimeThemeFormatScheme';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/** Apply an archive theme to specific zero-based slide positions, or every slide.
	 * Slide-scoped theme overrides leave shared masters/layouts and sibling slides intact.
	 */
	public async setSlidesTheme(themePath: string, slideIndices?: readonly number[]): Promise<void> {
		const path = themePath.trim().replace(/\\/g, '/');
		const themeXml = path.startsWith('ppt/theme/') && (await this.zip.file(path)?.async('string'));
		const themeData = themeXml ? (this.parser.parse(themeXml) as XmlObject) : undefined;
		const elements = (themeData?.['a:theme'] as XmlObject | undefined)?.['a:themeElements'] as
			XmlObject | undefined;
		if (!elements) throw new Error('The selected presentation theme is unavailable.');
		const presentation = this.parser.parse(
			await this.zip.file('ppt/presentation.xml')!.async('string'),
		) as XmlObject;
		const relationships = this.parser.parse(
			await this.zip.file('ppt/_rels/presentation.xml.rels')!.async('string'),
		) as XmlObject;
		const rels = this.ensureArray(
			(relationships['Relationships'] as XmlObject)['Relationship'],
		) as XmlObject[];
		const list = ((presentation['p:presentation'] as XmlObject)['p:sldIdLst'] ?? {}) as XmlObject;
		const ids = this.ensureArray(list['p:sldId']) as XmlObject[];
		const indices = slideIndices ?? ids.map((_, i) => i);
		const targets = [...new Set(indices)].map((index) => {
			if (!Number.isInteger(index) || index < 0 || index >= ids.length) {
				throw new Error('The selected slide is unavailable.');
			}
			const rel = rels.find((r) => r['@_Id'] === ids[index]['@_r:id']);
			const target = rel && safeResolveZipPath('ppt', String(rel['@_Target'] ?? ''));
			if (!target || !this.zip.file(target)) throw new Error('The selected slide is unavailable.');
			return target;
		});
		const overrideXml = this.builder.build({
			'a:themeOverride': {
				'@_xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
				...elements,
			},
		});
		const contentTypes = this.parser.parse(
			await this.zip.file('[Content_Types].xml')!.async('string'),
		) as XmlObject;
		const types = contentTypes['Types'] as XmlObject;
		const overrides = this.ensureArray(types['Override']) as XmlObject[];
		for (const slidePath of targets) {
			if (!this.pendingSlideThemeSaveContexts.has(slidePath)) {
				this.pendingSlideThemeSaveContexts.set(slidePath, await this.loadThemeOverride(slidePath));
			}
			const relsPath = slidePath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
			const existingXml = await this.zip.file(relsPath)?.async('string');
			const data = existingXml ? (this.parser.parse(existingXml) as XmlObject) : {};
			const root = (data['Relationships'] ?? {
				'@_xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
			}) as XmlObject;
			const slideRels = this.ensureArray(root['Relationship']) as XmlObject[];
			const type =
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride';
			const old = slideRels.find((rel) => rel['@_Type'] === type);
			let overridePath = old && safeResolveZipPath('ppt/slides', String(old['@_Target'] ?? ''));
			if (!overridePath?.startsWith('ppt/theme/') || !overridePath.endsWith('.xml'))
				overridePath = undefined;
			// Overrides may be shared in imported decks. Never mutate a sibling's part.
			if (overridePath) {
				for (const file of this.zip.file(/\.rels$/)) {
					if (file.name === relsPath) continue;
					const otherData = this.parser.parse(await file.async('string')) as XmlObject;
					const otherRels = this.ensureArray(
						(otherData['Relationships'] as XmlObject)?.['Relationship'],
					) as XmlObject[];
					const base = file.name.substring(0, file.name.lastIndexOf('/_rels/'));
					if (
						otherRels.some(
							(rel) =>
								rel['@_Type'] === type &&
								safeResolveZipPath(base, String(rel['@_Target'] ?? '')) === overridePath,
						)
					) {
						overridePath = undefined;
						break;
					}
				}
			}
			if (!overridePath) {
				let suffix = 1;
				while (this.zip.file(`ppt/theme/themeOverride${suffix}.xml`)) suffix++;
				overridePath = `ppt/theme/themeOverride${suffix}.xml`;
			}
			this.zip.file(overridePath, overrideXml);
			// Theme format schemes can contain image relationships, relative to the same theme directory.
			const sourceRels = await this.zip
				.file(path.replace(/\/([^/]+)$/, '/_rels/$1.rels'))
				?.async('string');
			const overrideRelsPath = overridePath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
			if (sourceRels) this.zip.file(overrideRelsPath, sourceRels);
			else this.zip.remove(overrideRelsPath);
			let id = String(old?.['@_Id'] ?? '');
			if (!id) {
				let suffix = 1;
				while (slideRels.some((rel) => rel['@_Id'] === `rId${suffix}`)) suffix++;
				id = `rId${suffix}`;
			}
			root['Relationship'] = [
				...slideRels.filter((rel) => rel['@_Type'] !== type),
				{
					'@_Id': id,
					'@_Type': type,
					'@_Target': this.buildRelativeTargetPath(slidePath, overridePath),
				},
			];
			data['Relationships'] = root;
			this.zip.file(relsPath, this.builder.build(data));
			this.slideRelsMap
				.get(slidePath)
				?.set(id, this.buildRelativeTargetPath(slidePath, overridePath));
			this.themeOverrideCache.delete(overridePath);
			if (!overrides.some((entry) => entry['@_PartName'] === `/${overridePath}`)) {
				overrides.push({
					'@_PartName': `/${overridePath}`,
					'@_ContentType': 'application/vnd.openxmlformats-officedocument.themeOverride+xml',
				});
			}
		}
		types['Override'] = overrides;
		this.zip.file('[Content_Types].xml', this.builder.build(contentTypes));
	}

	protected override async withSlideThemeForSave(
		slide: PptxSlide,
		save: () => Promise<void>,
	): Promise<void> {
		const slideOverride = this.pendingSlideThemeSaveContexts.has(slide.id)
			? this.pendingSlideThemeSaveContexts.get(slide.id)
			: await this.loadThemeOverride(slide.id);
		if (!slideOverride && !this.pendingSlideThemeSaveContexts.has(slide.id)) {
			await save();
			return;
		}
		// Keep scheme references when saving the pre-Apply model or reopening.
		const restore = this.applyThemeOverrideState({});
		const previousMasterMap = this.currentMasterClrMap;
		const previousSlideMap = this.currentSlideClrMapOverride;
		try {
			await this.setActiveMasterForSlide(slide.id);
			this.currentSlideClrMapOverride = slide.clrMapOverride ?? null;
			const layoutPath = this.findLayoutPathForSlide(slide.id);
			const layoutOverride = layoutPath && (await this.loadThemeOverride(layoutPath));
			if (layoutOverride) this.applyThemeOverrideState(layoutOverride);
			if (slideOverride) this.applyThemeOverrideState(slideOverride);
			await save();
		} finally {
			restore();
			this.currentMasterClrMap = previousMasterMap;
			this.currentSlideClrMapOverride = previousSlideMap;
		}
	}

	/**
	 * Parse the `a:fmtScheme` element from the theme into a structured
	 * {@link PptxThemeFormatScheme}.  Each sub-list (fillStyleLst, lnStyleLst,
	 * effectStyleLst, bgFillStyleLst) contains up to three intensity levels
	 * (subtle / moderate / intense) referenced by 1-based index.
	 *
	 * When `rawXml` is supplied (the original theme/themeOverride XML
	 * source), the document order of fill children is recovered via a
	 * targeted regex scan and threaded into {@link parseFillStyleList} so
	 * mixed solid/grad/blip/grpFill lists round-trip in source order even
	 * though fast-xml-parser collapses heterogeneous siblings into typed
	 * buckets.
	 */
	protected parseFormatScheme(fmtScheme: XmlObject, rawXml?: string): PptxThemeFormatScheme {
		const name = String(fmtScheme['@_name'] || '').trim() || undefined;

		const fillOrder = extractFillStyleListChildOrder(rawXml, 'a:fillStyleLst');
		const bgFillOrder = extractFillStyleListChildOrder(rawXml, 'a:bgFillStyleLst');

		// --- Fill styles (a:fillStyleLst) ---
		const fillStyleLst = fmtScheme['a:fillStyleLst'] as XmlObject | undefined;
		const fillStyles = this.parseFillStyleList(fillStyleLst, fillOrder);

		// --- Line styles (a:lnStyleLst) ---
		const lnStyleLst = fmtScheme['a:lnStyleLst'] as XmlObject | undefined;
		const lineStyles = this.parseLineStyleList(lnStyleLst);

		// --- Effect styles (a:effectStyleLst) ---
		const effectStyleLst = fmtScheme['a:effectStyleLst'] as XmlObject | undefined;
		const effectStyles = this.parseEffectStyleList(effectStyleLst);

		// --- Background fill styles (a:bgFillStyleLst) ---
		const bgFillStyleLst = fmtScheme['a:bgFillStyleLst'] as XmlObject | undefined;
		const backgroundFillStyles = this.parseFillStyleList(bgFillStyleLst, bgFillOrder);

		return {
			name,
			fillStyles,
			lineStyles,
			effectStyles,
			backgroundFillStyles,
		};
	}

	/**
	 * Extract a colour map override from a `p:clrMapOvr` node.
	 * Returns `null` when the node is absent, empty, or specifies
	 * `a:masterClrMapping` (meaning "inherit from master").
	 */
	protected parseClrMapOverrideNode(
		clrMapOvr: XmlObject | undefined,
	): Record<string, string> | null {
		if (!clrMapOvr) {
			return null;
		}

		// <a:masterClrMapping/> means "use the master's map" — no override.
		if (clrMapOvr['a:masterClrMapping'] !== undefined) {
			return null;
		}

		const overrideNode = clrMapOvr['a:overrideClrMapping'] as XmlObject | undefined;
		if (!overrideNode) {
			return null;
		}

		const aliasKeys = [
			'bg1',
			'tx1',
			'bg2',
			'tx2',
			'accent1',
			'accent2',
			'accent3',
			'accent4',
			'accent5',
			'accent6',
			'hlink',
			'folHlink',
		];

		const overrideMap: Record<string, string> = {};
		for (const key of aliasKeys) {
			const mapped = String(overrideNode[`@_${key}`] || '')
				.trim()
				.toLowerCase();
			if (mapped) {
				overrideMap[key] = mapped;
			}
		}

		return Object.keys(overrideMap).length > 0 ? overrideMap : null;
	}

	/**
	 * Parse `p:clrMapOvr` from a slide's XML.  Returns a mapping from logical
	 * colour name (bg1, tx1, ...) to the theme colour slot it should resolve to,
	 * or `null` when the slide inherits from the master (`a:masterClrMapping`
	 * or no override present).
	 */
	protected parseSlideClrMapOverride(slideXml: XmlObject): Record<string, string> | null {
		const clrMapOvr = (slideXml['p:sld'] as XmlObject | undefined)?.['p:clrMapOvr'] as
			XmlObject | undefined;
		return this.parseClrMapOverrideNode(clrMapOvr);
	}

	/**
	 * Parse `p:clrMapOvr` from a slide layout's XML.  Layouts can override
	 * the master's colour map just like slides can.
	 */
	protected parseLayoutClrMapOverride(layoutXml: XmlObject): Record<string, string> | null {
		const clrMapOvr = (layoutXml['p:sldLayout'] as XmlObject | undefined)?.['p:clrMapOvr'] as
			XmlObject | undefined;
		return this.parseClrMapOverrideNode(clrMapOvr);
	}

	/**
	 * Apply a theme override for the duration of parsing a layout or slide.
	 * Saves the current theme state, applies the override, and returns a
	 * restore function that must be called when done.
	 */
	protected applyThemeOverrideState(override: {
		colorOverrides?: Record<string, string>;
		formatSchemeOverride?: PptxThemeFormatScheme;
	}): () => void {
		const prevColorMap = { ...this.themeColorMap };
		const prevFontMap = { ...this.themeFontMap };
		const prevFormatScheme = this.themeFormatScheme;

		// Apply colour overrides
		if (override.colorOverrides) {
			for (const [key, value] of Object.entries(override.colorOverrides)) {
				if (key.startsWith('__fontOverride_')) {
					// Font override
					const fontKey = key.replace('__fontOverride_', '');
					this.themeFontMap[fontKey] = value;
					// Also set EA and CS variants
					if (fontKey === 'mj-lt') {
						this.themeFontMap['mj-ea'] = value;
						this.themeFontMap['mj-cs'] = value;
					} else if (fontKey === 'mn-lt') {
						this.themeFontMap['mn-ea'] = value;
						this.themeFontMap['mn-cs'] = value;
					}
				} else {
					this.themeColorMap[key] = value;
				}
			}
			// Re-apply aliases
			this.themeColorMap['tx1'] = this.themeColorMap['dk1'] || prevColorMap['dk1'];
			this.themeColorMap['bg1'] = this.themeColorMap['lt1'] || prevColorMap['lt1'];
			this.themeColorMap['tx2'] = this.themeColorMap['dk2'] || prevColorMap['dk2'];
			this.themeColorMap['bg2'] = this.themeColorMap['lt2'] || prevColorMap['lt2'];
		}

		// Apply format scheme override
		if (override.formatSchemeOverride) {
			this.themeFormatScheme = override.formatSchemeOverride;
		}

		// Return restore function
		return () => {
			this.themeColorMap = prevColorMap;
			this.themeFontMap = prevFontMap;
			this.themeFormatScheme = prevFormatScheme;
		};
	}

	/**
	 * Load and apply a theme override part referenced from a layout's or
	 * slide's relationships.  Theme overrides (`themeOverride*.xml`) can
	 * replace the colour scheme, font scheme, and/or format scheme for
	 * the scope of a specific layout or slide.
	 *
	 * The method caches results so each override file is parsed at most once.
	 * Overrides are applied as temporary deltas: the caller should save and
	 * restore `themeColorMap`, `themeFontMap`, and `themeFormatScheme` around
	 * the scope where the override is active.
	 */
	protected async loadThemeOverride(partBasePath: string): Promise<{
		colorOverrides?: Record<string, string>;
		formatSchemeOverride?: PptxThemeFormatScheme;
	} | null> {
		// Resolve relationship to find theme override file
		const rels = this.slideRelsMap.get(partBasePath);
		if (!rels) {
			return null;
		}

		let overridePath: string | undefined;
		for (const [, target] of rels) {
			if (target.includes('themeOverride')) {
				// Target is relative to the part, e.g. "../theme/themeOverride1.xml"
				const partDir = partBasePath.substring(0, partBasePath.lastIndexOf('/') + 1);
				overridePath = this.resolvePath(partDir, target);
				break;
			}
		}
		if (!overridePath) {
			return null;
		}

		// Check cache
		const cached = this.themeOverrideCache.get(overridePath);
		if (cached) {
			return cached;
		}

		try {
			const overrideXml = await this.zip.file(overridePath)?.async('string');
			if (!overrideXml) {
				return null;
			}

			const overrideData = this.parser.parse(overrideXml) as XmlObject;
			const root = overrideData['a:themeOverride'] as XmlObject | undefined;
			if (!root) {
				return null;
			}

			const result: {
				colorOverrides?: Record<string, string>;
				formatSchemeOverride?: PptxThemeFormatScheme;
			} = {};

			// Colour scheme override
			const clrScheme = root['a:clrScheme'] as XmlObject | undefined;
			if (clrScheme) {
				const colorOverrides: Record<string, string> = {};
				const schemeKeys = [
					'dk1',
					'lt1',
					'dk2',
					'lt2',
					'accent1',
					'accent2',
					'accent3',
					'accent4',
					'accent5',
					'accent6',
					'hlink',
					'folHlink',
				];
				for (const key of schemeKeys) {
					const colorNode = clrScheme[`a:${key}`] as XmlObject | undefined;
					const parsed = this.parseColorChoice(colorNode);
					if (parsed) {
						colorOverrides[key] = parsed;
					}
				}
				if (Object.keys(colorOverrides).length > 0) {
					result.colorOverrides = colorOverrides;
				}
			}

			// Format scheme override
			const fmtScheme = root['a:fmtScheme'] as XmlObject | undefined;
			if (fmtScheme) {
				result.formatSchemeOverride = this.parseFormatScheme(fmtScheme);
			}

			// Font scheme override (apply to themeFontMap temporarily)
			const fontScheme = root['a:fontScheme'] as XmlObject | undefined;
			if (fontScheme) {
				const majorFontNode = fontScheme['a:majorFont'] as XmlObject | undefined;
				const minorFontNode = fontScheme['a:minorFont'] as XmlObject | undefined;
				const majorLatin = majorFontNode?.['a:latin'] as XmlObject | undefined;
				const minorLatin = minorFontNode?.['a:latin'] as XmlObject | undefined;
				const majorFont = this.normalizeTypefaceToken(String(majorLatin?.['@_typeface'] || ''));
				const minorFont = this.normalizeTypefaceToken(String(minorLatin?.['@_typeface'] || ''));
				if (!result.colorOverrides) {
					result.colorOverrides = {};
				}
				// Font overrides are stored alongside colour overrides for simplicity
				if (majorFont) {
					result.colorOverrides['__fontOverride_mj-lt'] = majorFont;
				}
				if (minorFont) {
					result.colorOverrides['__fontOverride_mn-lt'] = minorFont;
				}
				// M4: per-script font overrides on a theme override are
				// not currently exposed through this side-channel object,
				// but we capture them so they survive theme writer
				// passthrough when the override file itself lives on disk
				// (overrides are written verbatim on save and only the
				// main theme.xml is regenerated by Item 1's writer).
			}

			this.themeOverrideCache.set(overridePath, result);
			return result;
		} catch (error) {
			console.warn(`Failed to load theme override at ${overridePath}:`, error);
			return null;
		}
	}
}

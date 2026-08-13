import { hasTextProperties } from 'pptx-viewer-core';
import type { PptxElement, PptxEmbeddedFont } from 'pptx-viewer-core';
import type { ViewerFontSource } from 'pptx-viewer-shared';
import JSZip from 'jszip';
import React, { useState, useRef, useEffect } from 'react';
import { useTranslation } from 'react-i18next';
import {
	LuChevronDown,
	LuClipboardPaste,
	LuCopy,
	LuPaintbrush,
	LuScissors,
	LuUpload,
} from 'react-icons/lu';

import type { ElementClipboardPayload } from '../../types';
import { cn } from '../../utils';
import { RibbonMenu } from './RibbonMenu';
import { SlidesGroup } from './SlidesGroup';
import { gB, gL, grp, ic, sep } from './toolbar-constants';

export interface HomeSectionProps {
	canEdit: boolean;
	clipboardPayload: ElementClipboardPayload | null;
	formatPainterActive?: boolean;
	canActivateFormatPainter?: boolean;
	onCopy: () => void;
	onCut: () => void;
	onPaste: () => void;
	onToggleFormatPainter?: () => void;
	layoutOptions: Array<{ path: string; name: string }>;
	currentLayoutPath?: string;
	onInsertSlideFromLayout: (path: string, name?: string) => void;
	onApplyLayout?: (path: string) => void;
	onResetSlide?: () => void;
	onAddSection?: () => void;
	selectedElement?: PptxElement | null;
	onUpdateTextStyle?: (style: Record<string, unknown>) => void;
	themeFonts?: { heading?: string; body?: string };
	embeddedFontFamilies?: string[];
	onUploadCustomFontPackage?: (file: File, fonts: ViewerFontSource[]) => void | Promise<void>;
	onEmbedCustomFonts?: (fonts: PptxEmbeddedFont[]) => void;
}

export function extractFontInfo(
	element?: PptxElement | null,
	themeFonts?: { heading?: string; body?: string },
): { fontFamily: string; fontSize: string } {
	const placeholderType = (element as { placeholderType?: string } | null | undefined)?.placeholderType;
	const isHeading = placeholderType === 'title' || placeholderType === 'ctrTitle';
	const defaultFontFamily =
		(isHeading ? themeFonts?.heading : themeFonts?.body) ??
		themeFonts?.body ??
		themeFonts?.heading ??
		'Segoe UI';
	const defaults = { fontFamily: defaultFontFamily, fontSize: '24' };
	if (!element) {
		return defaults;
	}
	if (!hasTextProperties(element)) {
		return defaults;
	}

	const segStyle = element.textSegments?.[0]?.style;
	const textStyle = element.textStyle;

	const fontFamily = segStyle?.fontFamily ?? textStyle?.fontFamily ?? defaults.fontFamily;
	const fontSize = segStyle?.fontSize ?? textStyle?.fontSize;

	return {
		fontFamily,
		fontSize: fontSize !== undefined && fontSize !== null ? String(fontSize) : defaults.fontSize,
	};
}

const COMMON_FONTS = [
	'Abadi',
	'Aptos',
	'Aptos Display',
	'Arial',
	'Arial Black',
	'Arial Narrow',
	'Bahnschrift',
	'Baskerville',
	'Book Antiqua',
	'Bookman Old Style',
	'Calibri',
	'Calibri Light',
	'Cambria',
	'Candara',
	'Century',
	'Century Gothic',
	'Consolas',
	'Comic Sans MS',
	'Courier New',
	'Franklin Gothic Book',
	'Franklin Gothic Demi',
	'Franklin Gothic Medium',
	'Garamond',
	'Georgia',
	'Gill Sans MT',
	'Helvetica Neue',
	'Helvetica',
	'Impact',
	'Inter',
	'Lucida Console',
	'Lucida Sans Unicode',
	'Microsoft Sans Serif',
	'Montserrat',
	'Noto Sans',
	'Open Sans',
	'Palatino Linotype',
	'Poppins',
	'Roboto',
	'Rockwell',
	'Segoe UI',
	'Source Sans Pro',
	'Tahoma',
	'Times New Roman',
	'Trebuchet MS',
	'Tw Cen MT',
	'Tw Cen MT Condensed',
	'Verdana',
];

const COMMON_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 96];

const CUSTOM_FONT_FILE_EXTENSION = /\.(?:ttf|otf|woff2?)$/iu;

const getZipEntryBaseName = (path: string): string =>
	path.replace(/\\/gu, '/').split('/').filter(Boolean).pop() ?? path;

const isUsableZipFontEntry = (path: string): boolean => {
	const normalized = path.replace(/\\/gu, '/').toLowerCase();
	if (normalized.includes('__macosx/') || normalized.endsWith('/.ds_store')) return false;
	return CUSTOM_FONT_FILE_EXTENSION.test(normalized);
};

export function HomeSection(p: HomeSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const [fontMenuOpen, setFontMenuOpen] = useState(false);
	const [sizeMenuOpen, setSizeMenuOpen] = useState(false);
	const [copiedFeedback, setCopiedFeedback] = useState(false);
	const [cutFeedback, setCutFeedback] = useState(false);
	const [customFontFamilies, setCustomFontFamilies] = useState<string[]>(() => {
		if (typeof window === 'undefined') return [];
		return (
			(window as Window & { __AIXA_PPTX_CUSTOM_FONT_FAMILIES__?: string[] })
				.__AIXA_PPTX_CUSTOM_FONT_FAMILIES__ ?? []
		);
	});
	const fontMenuRef = useRef<HTMLDivElement>(null);
	const fontFileInputRef = useRef<HTMLInputElement>(null);
	const sizeMenuRef = useRef<HTMLDivElement>(null);
	const { fontFamily, fontSize } = extractFontInfo(
		p.selectedElement,
		p.themeFonts,
	);
	const themeFontEntries = [
		p.themeFonts?.heading ? { family: p.themeFonts.heading, label: 'Headings' } : undefined,
		p.themeFonts?.body ? { family: p.themeFonts.body, label: 'Body' } : undefined,
	].filter((entry): entry is { family: string; label: string } => Boolean(entry));
	const embeddedFonts = Array.from(new Set(p.embeddedFontFamilies ?? []));
	const specialFonts = new Set([
		...themeFontEntries.map((entry) => entry.family),
		...embeddedFonts,
		...customFontFamilies,
	]);
	const availableFonts = COMMON_FONTS.filter((family) => !specialFonts.has(family));

	const applyFont = (family: string) => {
		p.onUpdateTextStyle?.({ fontFamily: family });
		setFontMenuOpen(false);
	};

	type LocalViewerFontSource = ViewerFontSource & { data: ArrayBuffer };

	const fontSourceFromFile = async (file: File): Promise<LocalViewerFontSource | null> => {
		const baseName = getZipEntryBaseName(file.name);
		const stem = baseName.replace(CUSTOM_FONT_FILE_EXTENSION, '');
		const bold = /(?:^|[-_\s])(?:bold|semibold|demibold|black)(?:$|[-_\s])/iu.test(stem);
		const italic = /(?:^|[-_\s])(?:italic|oblique)(?:$|[-_\s])/iu.test(stem);
		const family = stem
			.replace(/(?:[-_\s])(?:regular|normal|book|medium|bold|semibold|demibold|black|italic|oblique)+$/iu, '')
			.replace(/[-_]+/gu, ' ')
			.trim();
		if (!family) return null;
		const normalizedName = baseName.toLowerCase();
		const format = normalizedName.endsWith('.woff2') ? 'woff2' : normalizedName.endsWith('.woff') ? 'woff' : normalizedName.endsWith('.otf') ? 'opentype' : 'truetype';
		return {
			family,
			src: URL.createObjectURL(file),
			format,
			weight: bold ? 700 : 400,
			style: italic ? 'italic' : 'normal',
			data: await file.arrayBuffer(),
		};
	};

	const registerFontSources = async (sources: LocalViewerFontSource[]): Promise<void> => {
		if (typeof FontFace === 'undefined') return;
		const loadedSources: LocalViewerFontSource[] = [];
		for (const source of sources) {
			try {
				// Loading the extracted bytes directly avoids browser differences in
				// parsing blob URLs and missing MIME types from ZIP entries.
				const face = new FontFace(source.family, source.data, {
					weight: String(source.weight ?? 400),
					style: source.style ?? 'normal',
				});
				await face.load();
				document.fonts.add(face);
				loadedSources.push(source);
			} catch (error) {
				console.warn(`[PowerPointViewer] Could not load custom font ${source.family}.`, error);
			}
		}
		if (loadedSources.length === 0) {
			throw new Error('None of the selected font files could be loaded by this browser.');
		}
		const next = Array.from(new Set([...customFontFamilies, ...loadedSources.map((source) => source.family)]));
		setCustomFontFamilies(next);
		const fontWindow = window as Window & { __AIXA_PPTX_CUSTOM_FONT_FAMILIES__?: string[] };
		fontWindow.__AIXA_PPTX_CUSTOM_FONT_FAMILIES__ = next;
		window.dispatchEvent(new CustomEvent('aixa:pptx-custom-fonts', { detail: next }));
		if (loadedSources[0]) applyFont(loadedSources[0].family);
	};

	const registerLocalFontPackage = async (file: File): Promise<void> => {
		let sources: LocalViewerFontSource[] = [];
		if (/\.zip$/iu.test(file.name)) {
			const zip = await JSZip.loadAsync(file);
			// JSZip exposes a flat map containing entries from every directory,
			// so this finds fonts no matter how deeply they are nested.
			const entries = Object.values(zip.files).filter(
				(entry) => !entry.dir && isUsableZipFontEntry(entry.name),
			);
			sources = (await Promise.all(entries.map(async (entry) => {
				const blob = await entry.async('blob');
				return fontSourceFromFile(
					new File([blob], getZipEntryBaseName(entry.name), { type: blob.type }),
				);
			}))).filter((source): source is LocalViewerFontSource => Boolean(source));
		} else {
			const source = await fontSourceFromFile(file);
			if (source) sources = [source];
		}
		if (sources.length === 0) throw new Error('No browser-compatible font files were found.');
		// The host callback is the security gate (for example an antivirus scan).
		// It must complete before untrusted bytes are registered or embedded.
		await p.onUploadCustomFontPackage?.(
			file,
			sources.map(({ data: _data, ...source }) => source),
		);
		await registerFontSources(sources);
		p.onEmbedCustomFonts?.(
			sources.map((source) => ({
				name: source.family,
				dataUrl: '',
				bold: Number(source.weight ?? 400) >= 600,
				italic: source.style === 'italic' || source.style === 'oblique',
				format: source.format,
				rawFontData: new Uint8Array(source.data.slice(0)),
			})),
		);
	};

	useEffect(() => {
		const handleCustomFonts = (event: Event) => {
			const families = (event as CustomEvent<unknown>).detail;
			const incoming = Array.isArray(families)
				? families.filter((family): family is string => typeof family === 'string' && family.trim().length > 0)
				: [];
			// Host updates may arrive after a local ZIP was loaded. Preserve the
			// locally registered families until the host persists and echoes them.
			setCustomFontFamilies((current) => Array.from(new Set([...current, ...incoming])));
		};
		window.addEventListener('aixa:pptx-custom-fonts', handleCustomFonts);
		return () => window.removeEventListener('aixa:pptx-custom-fonts', handleCustomFonts);
	}, []);

	// Close font menu on outside click
	useEffect(() => {
		if (!fontMenuOpen) {
			return;
		}
		const handler = (e: MouseEvent) => {
			if (fontMenuRef.current && !fontMenuRef.current.contains(e.target as Node)) {
				setFontMenuOpen(false);
			}
		};
		document.addEventListener('mousedown', handler);
		return () => document.removeEventListener('mousedown', handler);
	}, [fontMenuOpen]);

	// Close size menu on outside click
	useEffect(() => {
		if (!sizeMenuOpen) {
			return;
		}
		const handler = (e: MouseEvent) => {
			if (sizeMenuRef.current && !sizeMenuRef.current.contains(e.target as Node)) {
				setSizeMenuOpen(false);
			}
		};
		document.addEventListener('mousedown', handler);
		return () => document.removeEventListener('mousedown', handler);
	}, [sizeMenuOpen]);

	return (
		<>
			{/* Clipboard group */}
			<div className='flex flex-col items-center gap-0.5'>
				<div className={grp}>
					<button
						type='button'
						onClick={p.onPaste}
						disabled={!p.clipboardPayload || !p.canEdit}
						className={gB}
						title={t('pptx.arrange.paste')}
					>
						<LuClipboardPaste className={ic} />
					</button>
					<button
						type='button'
						onClick={() => {
							p.onCut();
							setCutFeedback(true);
							setTimeout(() => setCutFeedback(false), 600);
						}}
						disabled={!p.canEdit}
						className={cn(gB, cutFeedback && 'bg-green-600/20 text-green-400')}
						title={t('pptx.arrange.cut')}
					>
						<LuScissors className={ic} />
					</button>
					<button
						type='button'
						onClick={() => {
							p.onCopy();
							setCopiedFeedback(true);
							setTimeout(() => setCopiedFeedback(false), 600);
						}}
						className={cn(gB, copiedFeedback && 'bg-green-600/20 text-green-400')}
						title={t('pptx.arrange.copy')}
					>
						<LuCopy className={ic} />
					</button>
					{p.onToggleFormatPainter && (
						<button
							type='button'
							onClick={p.onToggleFormatPainter}
							disabled={
								!p.canEdit || (p.canActivateFormatPainter === false && !p.formatPainterActive)
							}
							data-testid='format-painter-toggle'
							data-active={p.formatPainterActive ? 'true' : 'false'}
							className={cn(
								gL,
								p.formatPainterActive ? 'bg-amber-600 hover:bg-amber-500 text-amber-50' : '',
							)}
							title={t('pptx.arrange.formatPainter')}
						>
							<LuPaintbrush className={ic} />
						</button>
					)}
				</div>
				<span className='text-[9px] text-muted-foreground leading-none'>
					{t('pptx.ribbon.clipboard')}
				</span>
			</div>

			{sep}

			<SlidesGroup
				canEdit={p.canEdit}
				layoutOptions={p.layoutOptions}
				currentLayoutPath={p.currentLayoutPath}
				onInsertSlideFromLayout={p.onInsertSlideFromLayout}
				onApplyLayout={p.onApplyLayout}
				onResetSlide={p.onResetSlide}
				onAddSection={p.onAddSection}
			/>

			{/* Font group */}
			<div className='flex flex-col items-center gap-0.5'>
				<div className='flex items-center gap-1'>
					<div className='relative' ref={fontMenuRef}>
						<button
							type='button'
							onClick={() => setFontMenuOpen((v) => !v)}
							className='inline-flex items-center justify-between px-2 py-1 rounded-sm border border-border/60 bg-background/60 text-[11px] text-foreground min-w-[120px] truncate hover:bg-accent/40 transition-colors cursor-pointer'
						>
							<span className='truncate'>{fontFamily}</span>
							<LuChevronDown className='w-3 h-3 ml-1 shrink-0 text-muted-foreground' />
						</button>
						{fontMenuOpen && (
							<RibbonMenu anchorRef={fontMenuRef} className='flex flex-col w-72 pt-1'>
								<div className='rounded-lg border border-border bg-popover backdrop-blur-lg shadow-2xl py-1 max-h-[420px] overflow-y-auto'>
									{themeFontEntries.length > 0 && (
										<>
											<div className='px-3 pb-1 pt-2 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground'>
												Theme fonts
											</div>
											{themeFontEntries.map(({ family, label }) => (
												<button
													key={`${label}-${family}`}
													type='button'
													className='flex w-full items-center justify-between gap-3 px-3 py-1.5 text-sm text-foreground hover:bg-muted'
													style={{ fontFamily: family }}
													onClick={() => applyFont(family)}
												>
													<span className='truncate'>{family}</span>
													<span className='shrink-0 text-[10px] text-muted-foreground'>({label})</span>
												</button>
											))}
										</>
									)}
									{embeddedFonts.length > 0 && (
										<>
											<div className='border-t border-border/60 px-3 pb-1 pt-2 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground'>
												Embedded in presentation
											</div>
											{embeddedFonts.map((family) => (
												<button
													key={`embedded-${family}`}
													type='button'
													className='flex w-full items-center gap-2 px-3 py-1.5 text-sm text-foreground hover:bg-muted'
													style={{ fontFamily: family }}
													onClick={() => applyFont(family)}
												>
													{family}
												</button>
											))}
										</>
									)}
									{customFontFamilies.length > 0 && (
										<>
											<div className='border-t border-border/60 px-3 pb-1 pt-2 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground'>
												Custom fonts
											</div>
											{customFontFamilies.map((family) => (
												<button
													key={`custom-${family}`}
													type='button'
													className='flex w-full items-center gap-2 px-3 py-1.5 text-sm text-foreground hover:bg-muted'
													style={{ fontFamily: family }}
													onClick={() => applyFont(family)}
												>
													{family}
												</button>
											))}
										</>
									)}
									<div className='border-t border-border/60 px-3 pb-1 pt-2 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground'>
										All fonts
									</div>
									{availableFonts.map((f) => (
										<button
											key={f}
											type='button'
											className='flex items-center gap-2 w-full px-3 py-1.5 text-xs text-foreground hover:bg-muted transition-colors'
											style={{ fontFamily: f }}
											onClick={() => {
											applyFont(f);
										}}
										>
											{f}
										</button>
									))}
									<button
										type='button'
										className='sticky bottom-0 flex w-full items-center gap-2 border-t border-border bg-popover px-3 py-2 text-xs font-medium text-primary hover:bg-muted'
										onClick={() => fontFileInputRef.current?.click()}
									>
										<LuUpload className='h-4 w-4' />
										Add custom font package
									</button>
									<input
										ref={fontFileInputRef}
										type='file'
										accept='.zip,.ttf,.otf,.woff,.woff2,application/zip,font/ttf,font/otf,font/woff,font/woff2'
										className='hidden'
										onChange={(event) => {
											const file = event.currentTarget.files?.[0];
											if (file) void registerLocalFontPackage(file);
											event.currentTarget.value = '';
										}}
									/>
								</div>
							</RibbonMenu>
						)}
					</div>
					<div className='relative' ref={sizeMenuRef}>
						<button
							type='button'
							onClick={() => setSizeMenuOpen((v) => !v)}
							className='inline-flex items-center justify-between px-2 py-1 rounded-sm border border-border/60 bg-background/60 text-[11px] text-foreground min-w-[50px] text-center hover:bg-accent/40 transition-colors cursor-pointer'
						>
							<span className='truncate'>{fontSize}</span>
							<LuChevronDown className='w-3 h-3 ml-1 shrink-0 text-muted-foreground' />
						</button>
						{sizeMenuOpen && (
							<RibbonMenu anchorRef={sizeMenuRef} className='flex flex-col w-48 pt-1'>
								<div className='rounded-lg border border-border bg-popover backdrop-blur-lg shadow-2xl py-1 max-h-60 overflow-y-auto'>
									{COMMON_SIZES.map((s) => (
										<button
											key={s}
											type='button'
											className='flex items-center gap-2 w-full px-3 py-1.5 text-xs text-foreground hover:bg-muted transition-colors'
											onClick={() => {
												p.onUpdateTextStyle?.({ fontSize: s });
												setSizeMenuOpen(false);
											}}
										>
											{s}
										</button>
									))}
								</div>
							</RibbonMenu>
						)}
					</div>
				</div>
				<span className='text-[9px] text-muted-foreground leading-none'>Font</span>
			</div>

			{sep}
		</>
	);
}

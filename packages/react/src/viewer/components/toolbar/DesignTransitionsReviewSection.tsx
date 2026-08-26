import React from 'react';
import { useTranslation } from 'react-i18next';
import type { PptxSlide, PptxSlideTransition, PptxTransitionType } from 'pptx-viewer-core';
import {
	LuCopy,
	LuMonitor,
	LuPaintBucket,
	LuPalette,
	LuPanelRight,
	LuPencil,
	LuPlay,
} from 'react-icons/lu';

import { cn } from '../../utils';
import { ic, ics, pill, sep } from './toolbar-constants';

/* ── Design ────────────────────────────────────────────── */

export interface DesignSectionProps {
	canEdit: boolean;
	onToggleThemeGallery: () => void;
	isThemeGalleryOpen: boolean;
	onToggleThemeEditor: () => void;
	isThemeEditorOpen: boolean;
	onOpenDocumentProperties?: () => void;
	onToggleInspector?: () => void;
	isInspectorPaneOpen?: boolean;
}

export function DesignSection(p: DesignSectionProps): React.ReactElement {
	const { t } = useTranslation();

	return (
		<>
			{/* Themes */}
			<button
				onClick={p.onToggleThemeGallery}
				disabled={!p.canEdit}
				className={cn(
					pill,
					p.isThemeGalleryOpen ? 'bg-primary hover:bg-primary/80 text-white' : '',
				)}
				title={t('pptx.ribbon.browseThemesTitle')}
			>
				<LuPalette className={ics} />
				{t('pptx.ribbon.browseThemes')}
			</button>
			<button
				onClick={p.onToggleThemeEditor}
				disabled={!p.canEdit}
				className={cn(pill, p.isThemeEditorOpen ? 'bg-primary hover:bg-primary/80 text-white' : '')}
				title={t('pptx.ribbon.editThemeTitle')}
			>
				<LuPencil className={ics} />
				{t('pptx.ribbon.editTheme')}
			</button>

			{sep}

			{/* Customize */}
			{p.onOpenDocumentProperties && (
				<button
					onClick={p.onOpenDocumentProperties}
					className={pill}
					title={t('pptx.ribbon.slideSizeTitle')}
				>
					<LuMonitor className={ics} />
					{t('pptx.ribbon.slideSize')}
				</button>
			)}
			{p.onToggleInspector && (
				<button
					onClick={p.onToggleInspector}
					className={cn(
						pill,
						p.isInspectorPaneOpen ? 'bg-primary hover:bg-primary/80 text-white' : '',
					)}
					title={t('pptx.ribbon.formatBackgroundTitle')}
				>
					<LuPaintBucket className={ics} />
					{t('pptx.ribbon.formatBackground')}
				</button>
			)}
		</>
	);
}

/* ── Transitions ───────────────────────────────────────── */

const TRANSITION_PRESETS = [
	{ value: 'none', labelKey: 'pptx.ribbon.transition.none' },
	{ value: 'fade', labelKey: 'pptx.ribbon.transition.fade' },
	{ value: 'push', labelKey: 'pptx.ribbon.transition.push' },
	{ value: 'wipe', labelKey: 'pptx.ribbon.transition.wipe' },
	{ value: 'split', labelKey: 'pptx.ribbon.transition.split' },
	{ value: 'reveal', labelKey: 'pptx.ribbon.transition.reveal' },
	{ value: 'cut', labelKey: 'pptx.ribbon.transition.cut' },
	{ value: 'cover', labelKey: 'pptx.ribbon.transition.cover' },
	{ value: 'uncover', labelKey: 'pptx.ribbon.transition.uncover' },
] as const;

const DEFAULT_TRANSITION_DURATION_MS = 500;
const MAX_TRANSITION_DURATION_MS = 10_000;
const MAX_ADVANCE_AFTER_MS = 5_999_990;

/** Parse seconds, mm:ss, or hh:mm:ss input without allowing invalid clock values. */
export function parseTransitionTimeInput(value: string): number | null {
	const trimmed = value.trim();
	if (!trimmed) return null;
	const parts = trimmed.split(':');
	if (parts.length > 3 || parts.some((part) => !/^\d+(?:\.\d+)?$/u.test(part))) {
		return null;
	}
	const values = parts.map(Number);
	if (values.some((part) => !Number.isFinite(part))) return null;
	if (parts.length > 1 && values.at(-1)! >= 60) return null;
	if (parts.length === 3 && values[1]! >= 60) return null;

	const seconds = values.reduce((total, part) => total * 60 + part, 0);
	const milliseconds = Math.round(seconds * 1000);
	return Number.isSafeInteger(milliseconds) ? milliseconds : null;
}

export function formatTransitionDuration(durationMs?: number): string {
	const safeDuration = typeof durationMs === 'number' && Number.isFinite(durationMs)
		? Math.max(0, Math.min(MAX_TRANSITION_DURATION_MS, durationMs))
		: DEFAULT_TRANSITION_DURATION_MS;
	return (safeDuration / 1000).toFixed(2).padStart(5, '0');
}

export function formatAdvanceAfter(advanceAfterMs?: number): string {
	const safeDuration = typeof advanceAfterMs === 'number' && Number.isFinite(advanceAfterMs)
		? Math.max(0, Math.min(MAX_ADVANCE_AFTER_MS, advanceAfterMs))
		: 0;
	const minutes = Math.floor(safeDuration / 60_000);
	const seconds = (safeDuration % 60_000) / 1000;
	return `${String(minutes).padStart(2, '0')}:${seconds.toFixed(2).padStart(5, '0')}`;
}

export interface TransitionsSectionProps {
	canEdit: boolean;
	activeSlide?: PptxSlide;
	onTransitionChange: (updates: Partial<PptxSlideTransition>) => void;
	onApplyTransitionToAll: () => void;
	isInspectorPaneOpen: boolean;
	onToggleInspector: () => void;
}

export function TransitionsSection(p: TransitionsSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const { activeSlide, canEdit, onApplyTransitionToAll, onTransitionChange } = p;
	const transition = activeSlide?.transition;
	const selected = transition?.type ?? 'none';
	const advanceOnClick = transition?.advanceOnClick !== false;
	const advanceAfter = transition?.advanceAfterMs !== undefined;
	const editsDisabled = !canEdit || !activeSlide;
	const [durationDraft, setDurationDraft] = React.useState(() =>
		formatTransitionDuration(transition?.durationMs),
	);
	const [advanceAfterDraft, setAdvanceAfterDraft] = React.useState(() =>
		formatAdvanceAfter(transition?.advanceAfterMs),
	);

	React.useEffect(() => {
		setDurationDraft(formatTransitionDuration(transition?.durationMs));
	}, [activeSlide?.id, transition?.durationMs]);

	React.useEffect(() => {
		setAdvanceAfterDraft(formatAdvanceAfter(transition?.advanceAfterMs));
	}, [activeSlide?.id, transition?.advanceAfterMs]);

	const commitDuration = React.useCallback(() => {
		const parsed = parseTransitionTimeInput(durationDraft);
		if (parsed === null) {
			setDurationDraft(formatTransitionDuration(transition?.durationMs));
			return;
		}
		const durationMs = Math.min(MAX_TRANSITION_DURATION_MS, parsed);
		setDurationDraft(formatTransitionDuration(durationMs));
		if (!editsDisabled && durationMs !== transition?.durationMs) {
			onTransitionChange({ durationMs });
		}
	}, [durationDraft, editsDisabled, onTransitionChange, transition?.durationMs]);

	const commitAdvanceAfter = React.useCallback(() => {
		const parsed = parseTransitionTimeInput(advanceAfterDraft);
		if (parsed === null) {
			setAdvanceAfterDraft(formatAdvanceAfter(transition?.advanceAfterMs));
			return;
		}
		const advanceAfterMs = Math.min(MAX_ADVANCE_AFTER_MS, parsed);
		setAdvanceAfterDraft(formatAdvanceAfter(advanceAfterMs));
		if (!editsDisabled && advanceAfter && advanceAfterMs !== transition?.advanceAfterMs) {
			onTransitionChange({ advanceAfterMs });
		}
	}, [
		advanceAfter,
		advanceAfterDraft,
		editsDisabled,
		onTransitionChange,
		transition?.advanceAfterMs,
	]);

	return (
		<>
			{/* Preview */}
			<button type='button' className={pill} title={t('pptx.ribbon.previewTransition')}>
				<LuPlay className={ics} />
				{t('pptx.ribbon.preview')}
			</button>

			{sep}

			{/* Transition preset gallery */}
			<div className='inline-flex items-center gap-0.5 overflow-x-auto max-w-[420px]'>
				{TRANSITION_PRESETS.map((preset) => (
					<button
						key={preset.value}
						type='button'
						onClick={() =>
							onTransitionChange({ type: preset.value as PptxTransitionType })
						}
						disabled={editsDisabled}
						className={cn(
							'flex-shrink-0 px-2 py-1 max-md:min-h-[44px] rounded border text-[11px] leading-tight transition-colors',
							selected === preset.value
								? 'border-primary bg-primary/10 text-primary font-medium'
								: 'border-border bg-muted hover:bg-accent text-foreground',
						)}
						title={t('pptx.ribbon.transitionTitle', { name: t(preset.labelKey) })}
					>
						{t(preset.labelKey)}
					</button>
				))}
			</div>

			{sep}

			{/* Duration */}
			<label className='inline-flex items-center gap-1.5 text-xs text-muted-foreground'>
				<span className='whitespace-nowrap'>{t('pptx.ribbon.duration')}</span>
				<input
					type='text'
					inputMode='decimal'
					value={durationDraft}
					onChange={(e) => setDurationDraft(e.target.value)}
					onBlur={commitDuration}
					onKeyDown={(e) => {
						if (e.key === 'Enter') e.currentTarget.blur();
						if (e.key === 'Escape') {
							setDurationDraft(formatTransitionDuration(transition?.durationMs));
							e.currentTarget.blur();
						}
					}}
					disabled={editsDisabled}
					className='w-14 px-1.5 py-1 rounded border border-border bg-muted text-xs text-foreground text-center'
					title={t('pptx.ribbon.transitionDurationTitle')}
				/>
			</label>

			{sep}

			{/* Sound */}
			<label className='inline-flex items-center gap-1.5 text-xs text-muted-foreground'>
				<span className='whitespace-nowrap'>{t('pptx.ribbon.sound')}</span>
				<select
					className='w-24 px-1.5 py-1 rounded border border-border bg-muted text-xs text-foreground'
					defaultValue='none'
					disabled
				>
					<option value='none'>{t('pptx.ribbon.soundNone')}</option>
				</select>
			</label>

			{sep}

			{/* Apply to All */}
			<button
				type='button'
				onClick={onApplyTransitionToAll}
				disabled={editsDisabled || !transition}
				className={pill}
				title={t('pptx.ribbon.applyTransitionToAll')}
			>
				<LuCopy className={ics} />
				{t('pptx.headerFooter.applyToAll')}
			</button>

			{sep}

			{/* Advance Slide group */}
			<div className='inline-flex flex-col gap-1 text-xs text-muted-foreground'>
				<span className='text-[10px] font-medium text-foreground'>
					{t('pptx.ribbon.advanceSlide')}
				</span>
				<label className='inline-flex items-center gap-1.5 cursor-pointer'>
					<input
						type='checkbox'
						checked={advanceOnClick}
						onChange={(e) => onTransitionChange({ advanceOnClick: e.target.checked })}
						disabled={editsDisabled}
						className='accent-primary h-3 w-3'
					/>
					<span className='whitespace-nowrap'>{t('pptx.ribbon.onMouseClick')}</span>
				</label>
				<label className='inline-flex items-center gap-1.5 cursor-pointer'>
					<input
						type='checkbox'
						checked={advanceAfter}
						onChange={(e) => {
							if (!e.target.checked) {
								onTransitionChange({ advanceAfterMs: undefined });
								return;
							}
							const parsed = parseTransitionTimeInput(advanceAfterDraft);
							const advanceAfterMs = Math.min(MAX_ADVANCE_AFTER_MS, parsed ?? 0);
							setAdvanceAfterDraft(formatAdvanceAfter(advanceAfterMs));
							onTransitionChange({ advanceAfterMs });
						}}
						disabled={editsDisabled}
						className='accent-primary h-3 w-3'
					/>
					<span className='whitespace-nowrap'>{t('pptx.ribbon.afterDuration')}</span>
					<input
						type='text'
						inputMode='decimal'
						value={advanceAfterDraft}
						onChange={(e) => setAdvanceAfterDraft(e.target.value)}
						onBlur={commitAdvanceAfter}
						onKeyDown={(e) => {
							if (e.key === 'Enter') e.currentTarget.blur();
							if (e.key === 'Escape') {
								setAdvanceAfterDraft(formatAdvanceAfter(transition?.advanceAfterMs));
								e.currentTarget.blur();
							}
						}}
						disabled={editsDisabled || !advanceAfter}
						className='w-16 px-1 py-0.5 rounded border border-border bg-muted text-xs text-foreground text-center disabled:opacity-50'
						title={t('pptx.ribbon.advanceAfterSeconds')}
					/>
				</label>
			</div>

			{sep}

			{/* Inspector */}
			<button
				type='button'
				onClick={p.onToggleInspector}
				className={cn(
					pill,
					p.isInspectorPaneOpen ? 'bg-primary hover:bg-primary/80 text-white' : '',
				)}
				title={t('pptx.ribbon.openInspectorTransitions')}
			>
				<LuPanelRight className={ic} />
				{t('pptx.ribbon.inspector')}
			</button>
		</>
	);
}

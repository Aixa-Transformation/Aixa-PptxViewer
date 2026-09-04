import type { PptxSlide, PptxSlideMaster } from 'pptx-viewer-core';
import React, { useRef } from 'react';
import { useTranslation } from 'react-i18next';
import { LuX } from 'react-icons/lu';

import { cn, normalizeHexColor } from '../../utils';
import { DebouncedColorInput } from './DebouncedColorInput';
import { CARD, HEADING, BTN } from './inspector-pane-constants';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

interface SlideBackgroundPanelProps {
	activeSlide: PptxSlide;
	canEdit: boolean;
	onUpdateSlide: (patch: Partial<PptxSlide>) => void;
	onApplySlideBackground?: (patch: SlideBackgroundPatch, allSlides: boolean) => void;

	/** Template-mode fields (only needed for master/layout editing) */
	editTemplateMode?: boolean;
	slideMasters?: PptxSlideMaster[];
	onSetTemplateBackground?: (path: string, color: string) => void;
	onGetTemplateBackgroundColor?: (path: string) => string | undefined;
}

export type SlideBackgroundPatch = Pick<
	PptxSlide,
	| 'backgroundColor'
	| 'backgroundImage'
	| 'backgroundGradient'
	| 'backgroundSource'
	| 'showMasterShapes'
>;

export function getDeckBackgroundPatch(activeSlide: PptxSlide): SlideBackgroundPatch {
	return {
		backgroundColor: activeSlide.backgroundColor,
		// An empty string is an intentional image removal. It lets the PPTX
		// writer distinguish a requested solid/gradient override from an
		// inherited colour that should preserve an existing image fill.
		backgroundImage: activeSlide.backgroundImage || '',
		backgroundGradient: activeSlide.backgroundGradient,
		backgroundSource:
			activeSlide.backgroundColor || activeSlide.backgroundImage || activeSlide.backgroundGradient
				? 'slide'
				: 'inherited',
		showMasterShapes: activeSlide.showMasterShapes,
	};
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function SlideBackgroundPanel({
	activeSlide,
	canEdit,
	onUpdateSlide,
	onApplySlideBackground,
	editTemplateMode,
	slideMasters,
	onSetTemplateBackground,
	onGetTemplateBackgroundColor,
}: SlideBackgroundPanelProps): React.ReactElement {
	const { t } = useTranslation();
	const bgImageInputRef = useRef<HTMLInputElement>(null);

	return (
		<>
			{/* Slide Background */}
			<div className={cn(CARD, 'space-y-2')}>
				<div className={HEADING}>Background</div>

				{/* Solid colour */}
				<label className='flex items-center gap-2 text-[11px]'>
					<span className='text-muted-foreground w-10 shrink-0'>Colour</span>
					<DebouncedColorInput
						value={normalizeHexColor(activeSlide.backgroundColor, '#ffffff')}
						disabled={!canEdit}
						className='h-6 w-8 rounded border border-border bg-muted cursor-pointer'
						onCommit={(hex) =>
							onUpdateSlide({
								backgroundColor: hex,
								backgroundImage: '',
								backgroundGradient: undefined,
								backgroundSource: 'slide',
							})
						}
					/>
					<span className='text-muted-foreground text-[10px] truncate'>
						{activeSlide.backgroundColor || 'none'}
					</span>
				</label>

				{/* Background image */}
				<div className='space-y-1'>
					<div className='flex items-center gap-2 text-[11px]'>
						<span className='text-muted-foreground w-10 shrink-0'>Image</span>
						<input
							ref={bgImageInputRef}
							type='file'
							accept='image/png,image/jpeg,image/gif,image/webp,image/svg+xml'
							className='hidden'
							disabled={!canEdit}
							onChange={(e) => {
								const file = e.target.files?.[0];
								if (!file) {
									return;
								}
								const reader = new FileReader();
								reader.onload = () => {
									if (typeof reader.result === 'string') {
										onUpdateSlide({
											backgroundImage: reader.result,
											backgroundGradient: undefined,
											backgroundSource: 'slide',
										});
									}
								};
								reader.readAsDataURL(file);
								e.target.value = '';
							}}
						/>
						<button
							type='button'
							className={cn(BTN, 'flex-1 text-center')}
							disabled={!canEdit}
							onClick={() => bgImageInputRef.current?.click()}
						>
							{activeSlide.backgroundImage
								? t('pptx.slideBackground.replaceImage')
								: t('pptx.slideBackground.chooseImage')}
						</button>
					</div>
					{activeSlide.backgroundImage && (
						<div className='relative mt-1'>
							<img
								src={activeSlide.backgroundImage}
								alt={t('pptx.slideBackground.backgroundPreview')}
								className='w-full h-16 object-cover rounded border border-border'
							/>
							<button
								type='button'
								className='absolute top-0.5 right-0.5 rounded bg-background/80 hover:bg-red-700 p-0.5 text-[10px] transition-colors'
								disabled={!canEdit}
								title={t('pptx.slideBackground.removeBackgroundImage')}
								onClick={() =>
									onUpdateSlide({
										backgroundImage: '',
										backgroundSource:
											activeSlide.backgroundColor || activeSlide.backgroundGradient
												? 'slide'
												: 'inherited',
									})
								}
							>
								<LuX className='w-3 h-3' />
							</button>
						</div>
					)}
				</div>

				{/* Clear background */}
				{(activeSlide.backgroundColor ||
					activeSlide.backgroundImage ||
					activeSlide.backgroundGradient) && (
					<button
						type='button'
						className={cn(BTN, 'w-full text-center text-red-400 hover:text-red-300')}
						disabled={!canEdit}
						onClick={() =>
							onUpdateSlide({
								backgroundColor: undefined,
								backgroundImage: undefined,
								backgroundGradient: undefined,
								backgroundSource: 'inherited',
							})
						}
					>
						{t('pptx.slideBackground.clearBackground')}
					</button>
				)}

				<label className='flex items-center gap-2 text-[11px]'>
					<input
						type='checkbox'
						checked={activeSlide.showMasterShapes === false}
						disabled={!canEdit}
						onChange={(event) =>
							onUpdateSlide({ showMasterShapes: event.target.checked ? false : true })
						}
					/>
					<span className='text-muted-foreground'>Hide background graphics</span>
				</label>

				<div className='grid grid-cols-2 gap-1.5'>
					<button
						type='button'
						className={BTN}
						disabled={!canEdit || !onApplySlideBackground}
						title='Applies the background settings to the current slide only'
						onClick={() => onApplySlideBackground?.(getDeckBackgroundPatch(activeSlide), false)}
					>
						{t('pptx.documentProperties.applyCurrentSlide')}
					</button>
					<button
						type='button'
						className={BTN}
						disabled={!canEdit || !onApplySlideBackground}
						title='Applies the background settings to every slide'
						onClick={() => onApplySlideBackground?.(getDeckBackgroundPatch(activeSlide), true)}
					>
						{t('pptx.documentProperties.applyAllSlides')}
					</button>
				</div>
			</div>

			{/* Master / Layout Background (template mode) */}
			{editTemplateMode && onSetTemplateBackground && onGetTemplateBackgroundColor && (
				<TemplateBackgroundCard
					activeSlide={activeSlide}
					slideMasters={slideMasters}
					canEdit={canEdit}
					onSetTemplateBackground={onSetTemplateBackground}
					onGetTemplateBackgroundColor={onGetTemplateBackgroundColor}
				/>
			)}
		</>
	);
}

// ---------------------------------------------------------------------------
// Template Background Card
// ---------------------------------------------------------------------------

function TemplateBackgroundCard({
	activeSlide,
	slideMasters,
	canEdit,
	onSetTemplateBackground,
	onGetTemplateBackgroundColor,
}: {
	activeSlide: PptxSlide;
	slideMasters: PptxSlideMaster[] | undefined;
	canEdit: boolean;
	onSetTemplateBackground: (path: string, color: string) => void;
	onGetTemplateBackgroundColor: (path: string) => string | undefined;
}): React.ReactElement {
	const { t } = useTranslation();
	return (
		<div className={cn(CARD, 'space-y-2')}>
			<div className={HEADING}>{t('pptx.slideBackground.templateBackgroundsHeading')}</div>

			{/* Layout background */}
			{activeSlide.layoutPath && (
				<label className='flex items-center gap-2 text-[11px]'>
					<span
						className='text-muted-foreground w-14 shrink-0 truncate'
						title={activeSlide.layoutName ?? activeSlide.layoutPath}
					>
						{t('pptx.master.layout')}
					</span>
					<DebouncedColorInput
						value={normalizeHexColor(
							onGetTemplateBackgroundColor(activeSlide.layoutPath),
							'#ffffff',
						)}
						disabled={!canEdit}
						className='h-6 w-8 rounded border border-border bg-muted cursor-pointer'
						onCommit={(hex) => onSetTemplateBackground(activeSlide.layoutPath!, hex)}
					/>
					<span className='text-muted-foreground text-[10px] truncate'>
						{activeSlide.layoutName ?? t('pptx.master.layout')}
					</span>
				</label>
			)}

			{/* Master background */}
			{(() => {
				const master = slideMasters?.find((m) =>
					m.layoutPaths?.includes(activeSlide.layoutPath ?? ''),
				);
				if (!master) {
					return null;
				}
				return (
					<label className='flex items-center gap-2 text-[11px]'>
						<span
							className='text-muted-foreground w-14 shrink-0 truncate'
							title={master.name ?? master.path}
						>
							{t('pptx.master.master')}
						</span>
						<DebouncedColorInput
							value={normalizeHexColor(onGetTemplateBackgroundColor(master.path), '#ffffff')}
							disabled={!canEdit}
							className='h-6 w-8 rounded border border-border bg-muted cursor-pointer'
							onCommit={(hex) => onSetTemplateBackground(master.path, hex)}
						/>
						<span className='text-muted-foreground text-[10px] truncate'>
							{master.name ?? t('pptx.master.master')}
						</span>
					</label>
				);
			})()}
		</div>
	);
}

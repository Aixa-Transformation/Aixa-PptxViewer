import type { PptxLayoutOption, PptxSlide } from 'pptx-viewer-core';
import React, { useState, useRef, useEffect, useCallback } from 'react';
import { useTranslation } from 'react-i18next';
import { LuChevronDown, LuFolderPlus, LuPlus, LuRotateCcw, LuLayoutGrid } from 'react-icons/lu';

import { cn } from '../../utils';
import { StaticElementRenderer } from '../StaticElementRenderer';
import { RibbonMenu } from './RibbonMenu';
import { ic, pill, sep } from './toolbar-constants';

export interface SlidesGroupProps {
	canEdit: boolean;
	layoutOptions: PptxLayoutOption[];
	currentLayoutPath?: string;
	onInsertSlideFromLayout: (path: string, name?: string) => void;
	onApplyLayout?: (path: string) => void;
	onResetSlide?: () => void;
	onAddSection?: () => void;
}

function LayoutPreview({ layout }: { layout: PptxLayoutOption }): React.ReactElement {
	const width = Math.max(layout.previewWidth ?? 960, 1);
	const height = Math.max(layout.previewHeight ?? 540, 1);
	const previewWidth = 128;
	const previewHeight = 72;
	const scale = Math.min(previewWidth / width, previewHeight / height);
	// Placeholder chrome lives inside the scaled slide canvas. Compensate for
	// that transform so its outline remains approximately 1.5px in the visible
	// thumbnail instead of shrinking to an almost invisible hairline.
	const placeholderBorderWidth = 1.5 / scale;
	const placeholderInsetWidth = 0.75 / scale;
	const slide: PptxSlide = {
		id: `layout-preview-${layout.path}`,
		rId: '',
		slideNumber: 0,
		elements: layout.previewElements ?? [],
		backgroundColor: layout.previewBackgroundColor ?? '#ffffff',
	};

	return (
		<div
			className='relative shrink-0 overflow-hidden border border-border/70 bg-white shadow-sm'
			style={{ width: previewWidth, height: previewHeight }}
		>
			<div
				className='absolute left-0 top-0 origin-top-left overflow-hidden'
				style={{
					width,
					height,
					transform: `scale(${scale})`,
					backgroundColor: layout.previewBackgroundColor ?? '#ffffff',
				}}
			>
				{(layout.previewElements ?? []).slice(0, 100).map((element, index) => (
					<StaticElementRenderer
						key={element.id}
						element={element}
						activeSlide={slide}
						allSlides={[slide]}
						zIndex={index}
					/>
				))}
				{(layout.previewPlaceholders ?? []).map((placeholder, index) =>
					placeholder.x !== undefined &&
					placeholder.y !== undefined &&
					placeholder.width !== undefined &&
					placeholder.height !== undefined ? (
						<div
							key={`${placeholder.type}-${placeholder.idx ?? index}`}
							className='absolute border-2 border-dashed border-slate-400/80 bg-white/10'
							style={{
								left: placeholder.x,
								top: placeholder.y,
								width: placeholder.width,
								height: placeholder.height,
								borderWidth: placeholderBorderWidth,
								borderColor: 'rgba(71, 85, 105, 0.92)',
								backgroundColor: 'rgba(255, 255, 255, 0.22)',
								boxShadow: `inset 0 0 0 ${placeholderInsetWidth}px rgba(255, 255, 255, 0.72)`,
							}}
						/>
					) : null,
				)}
			</div>
		</div>
	);
}

export function SlidesGroup(p: SlidesGroupProps): React.ReactElement {
	const { t } = useTranslation();
	const [newSlideMenuOpen, setNewSlideMenuOpen] = useState(false);
	const [layoutMenuOpen, setLayoutMenuOpen] = useState(false);
	const newSlideMenuRef = useRef<HTMLDivElement>(null);
	const layoutMenuRef = useRef<HTMLDivElement>(null);

	const handleNewSlide = useCallback(() => {
		if (p.layoutOptions.length > 0) {
			const first = p.layoutOptions[0];
			p.onInsertSlideFromLayout(first.path, first.name);
		}
	}, [p]);

	useEffect(() => {
		if (!newSlideMenuOpen) {
			return;
		}
		const handler = (e: MouseEvent) => {
			if (newSlideMenuRef.current && !newSlideMenuRef.current.contains(e.target as Node)) {
				setNewSlideMenuOpen(false);
			}
		};
		document.addEventListener('mousedown', handler);
		return () => document.removeEventListener('mousedown', handler);
	}, [newSlideMenuOpen]);

	useEffect(() => {
		if (!layoutMenuOpen) {
			return;
		}
		const handler = (e: MouseEvent) => {
			if (layoutMenuRef.current && !layoutMenuRef.current.contains(e.target as Node)) {
				setLayoutMenuOpen(false);
			}
		};
		document.addEventListener('mousedown', handler);
		return () => document.removeEventListener('mousedown', handler);
	}, [layoutMenuOpen]);

	return (
		<>
			<div className='flex flex-col items-center gap-0.5'>
				<div className='flex items-center gap-1'>
					{/* New Slide split button */}
					<div className='relative inline-flex items-center' ref={newSlideMenuRef}>
						<button
							type='button'
							onClick={handleNewSlide}
							disabled={!p.canEdit || p.layoutOptions.length === 0}
							className={cn(
								pill,
								'whitespace-nowrap',
								p.layoutOptions.length > 0 ? 'rounded-r-none' : '',
							)}
							title={t('pptx.home.newSlide')}
						>
							<LuPlus className={ic} />
							{t('pptx.home.newSlide')}
						</button>
						{p.layoutOptions.length > 0 && (
							<button
								type='button'
								disabled={!p.canEdit}
								className='inline-flex items-center justify-center self-stretch px-1 rounded-r bg-muted hover:bg-accent text-xs transition-colors border-l border-border/40 active:scale-95 active:opacity-80'
								title={t('pptx.home.chooseLayout')}
								onClick={() => setNewSlideMenuOpen((v) => !v)}
							>
								<LuChevronDown className='w-3 h-3' />
							</button>
						)}
						{newSlideMenuOpen && (
							<RibbonMenu anchorRef={newSlideMenuRef} className='flex flex-col w-[620px] pt-1'>
								<div className='grid grid-cols-4 gap-2 rounded-lg border border-border bg-popover backdrop-blur-lg shadow-2xl p-3 max-h-[520px] overflow-y-auto'>
									{p.layoutOptions.map((lo) => (
										<button
											key={lo.path}
											type='button'
											className='flex min-w-0 flex-col items-center gap-1 rounded p-1 text-xs text-foreground hover:bg-muted transition-colors'
											onClick={() => {
												p.onInsertSlideFromLayout(lo.path, lo.name);
												setNewSlideMenuOpen(false);
											}}
										>
											<LayoutPreview layout={lo} />
											<span className='w-full truncate text-center'>{lo.name}</span>
										</button>
									))}
								</div>
							</RibbonMenu>
						)}
					</div>

					{/* Layout button */}
					<div className='relative inline-flex items-center' ref={layoutMenuRef}>
						<button
							type='button'
							disabled={!p.canEdit || p.layoutOptions.length === 0}
							className={pill}
							title={t('pptx.master.layout')}
							onClick={() => setLayoutMenuOpen((v) => !v)}
						>
							<LuLayoutGrid className={ic} />
							{t('pptx.master.layout')}
						</button>
						{layoutMenuOpen && (
							<RibbonMenu anchorRef={layoutMenuRef} className='flex flex-col w-[620px] pt-1'>
								<div className='grid grid-cols-4 gap-2 rounded-lg border border-border bg-popover backdrop-blur-lg shadow-2xl p-3 max-h-[520px] overflow-y-auto'>
									{p.layoutOptions.map((lo) => {
										const isActive = lo.path === p.currentLayoutPath;
										return (
											<button
												key={lo.path}
												type='button'
												aria-current={isActive ? 'true' : undefined}
												className={cn(
													'relative flex min-w-0 flex-col items-center gap-1 rounded border-[3px] p-1 text-xs text-foreground transition-colors hover:bg-muted',
													isActive
														? 'border-blue-600 bg-blue-50 shadow-[0_0_0_2px_rgba(37,99,235,0.25)] dark:bg-blue-950/40'
														: 'border-transparent',
												)}
												onClick={() => {
													p.onApplyLayout?.(lo.path);
													setLayoutMenuOpen(false);
												}}
											>
												{isActive && (
													<span className='absolute right-1 top-1 z-10 rounded-full bg-blue-600 px-1.5 py-0.5 text-[9px] font-bold leading-none text-white shadow'>
														✓
													</span>
												)}
												<LayoutPreview layout={lo} />
												<span className='w-full truncate text-center'>{lo.name}</span>
											</button>
										);
									})}
								</div>
							</RibbonMenu>
						)}
					</div>

					{/* Reset button */}
					<button
						type='button'
						disabled={!p.canEdit}
						className={pill}
						title={t('pptx.sections.resetSlideTitle')}
						onClick={p.onResetSlide}
					>
						<LuRotateCcw className={ic} />
						{t('pptx.animations.reset')}
					</button>

					{/* Section button */}
					<button
						type='button'
						disabled={!p.canEdit}
						className={pill}
						title={t('pptx.sections.addSection')}
						onClick={p.onAddSection}
					>
						<LuFolderPlus className={ic} />
						{t('pptx.sections.sectionButtonLabel')}
					</button>
				</div>
				<span className='text-[9px] text-muted-foreground leading-none'>Slides</span>
			</div>

			{sep}
		</>
	);
}

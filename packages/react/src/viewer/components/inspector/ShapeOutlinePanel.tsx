import type {
	ConnectorArrowType,
	PptxTheme,
	ShapeStyle,
	StrokeDashType,
} from 'pptx-viewer-core';
import React, { useRef } from 'react';
import { useTranslation } from 'react-i18next';
import { LuBan, LuPipette } from 'react-icons/lu';

import { ARROW_SIZE_OPTIONS, CONNECTOR_ARROW_OPTIONS, STROKE_DASH_OPTIONS } from '../../constants';
import { cn, normalizeHexColor, openNativeEyeDropper } from '../../utils';
import { CARD, HEADING, INPUT } from './inspector-pane-constants';
import { ThemeColorGrid } from './ThemeColorGrid';

const STANDARD_OUTLINE_COLORS = [
	'#C00000',
	'#FF0000',
	'#FFC000',
	'#FFFF00',
	'#92D050',
	'#00B050',
	'#00B0F0',
	'#0070C0',
	'#002060',
	'#7030A0',
] as const;

const OUTLINE_WEIGHT_POINTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6] as const;
const POINT_TO_PX = 96 / 72;

const COMMON_DASH_OPTIONS = STROKE_DASH_OPTIONS.filter((option) => option.value !== 'custom');

interface ShapeOutlinePanelProps {
	style: ShapeStyle | undefined;
	theme: PptxTheme | undefined;
	canEdit: boolean;
	showArrows: boolean;
	onUpdateShapeStyle: (patch: Partial<ShapeStyle>) => void;
}

export function outlinePointsToPixels(points: number): number {
	return points * POINT_TO_PX;
}

export function outlinePixelsToPoints(pixels: number | undefined): number {
	return (pixels ?? outlinePointsToPixels(1)) / POINT_TO_PX;
}

export function createSolidOutlinePatch(
	color: string,
	style: ShapeStyle | undefined,
): Partial<ShapeStyle> {
	const currentWidth = style?.strokeWidth ?? 0;
	return {
		strokeFillMode: 'solid',
		strokeColor: normalizeHexColor(color, '#000000'),
		strokeWidth: currentWidth > 0 ? currentWidth : outlinePointsToPixels(1),
	};
}

export function createNoOutlinePatch(): Partial<ShapeStyle> {
	return {
		strokeFillMode: 'none',
		strokeColor: 'transparent',
		strokeWidth: 0,
	};
}

/**
 * PowerPoint-style outline controls backed only by schema-native line fields.
 * Every option here is written through ShapeStyle and therefore survives a
 * PPTX save/reopen instead of becoming a browser-only decoration.
 */
export function ShapeOutlinePanel({
	style,
	theme,
	canEdit,
	showArrows,
	onUpdateShapeStyle,
}: ShapeOutlinePanelProps): React.ReactElement {
	const { t } = useTranslation();
	const moreColorInputRef = useRef<HTMLInputElement>(null);
	const selectedColor = normalizeHexColor(style?.strokeColor, '#000000');
	const noOutline =
		style?.strokeFillMode === 'none' ||
		style?.strokeColor === 'transparent' ||
		style?.strokeWidth === 0;
	const weightPoints = outlinePixelsToPoints(style?.strokeWidth);

	const setColor = (color: string): void => {
		onUpdateShapeStyle(createSolidOutlinePatch(color, style));
	};

	const pickFromScreen = async (): Promise<void> => {
		const color = await openNativeEyeDropper();
		if (color) {
			setColor(color);
		}
	};

	return (
		<div className={cn(CARD, 'space-y-2')} data-testid='shape-outline-panel'>
			<div className={HEADING}>{t('pptx.drawing.shapeOutline', 'Shape Outline')}</div>

			<ThemeColorGrid
				theme={theme}
				selectedColor={noOutline ? undefined : selectedColor}
				disabled={!canEdit}
				onSelectColor={setColor}
			/>

			<div className='space-y-1'>
				<div className='text-[11px] uppercase tracking-wide text-muted-foreground'>
					{t('pptx.shapeOutline.standardColors', 'Standard Colors')}
				</div>
				<div className='grid grid-cols-10 gap-1'>
					{STANDARD_OUTLINE_COLORS.map((color) => (
						<button
							key={color}
							type='button'
							disabled={!canEdit}
							data-pptx-compact
							className={cn(
								'h-5 rounded-sm border transition-colors disabled:cursor-not-allowed disabled:opacity-40',
								!noOutline && selectedColor.toUpperCase() === color
									? 'border-primary ring-1 ring-primary'
									: 'border-border hover:border-foreground',
							)}
							style={{ backgroundColor: color }}
							title={color}
							aria-label={`Outline colour ${color}`}
							onClick={() => setColor(color)}
						/>
					))}
				</div>
			</div>

			<div className='divide-y divide-border rounded border border-border'>
				<button
					type='button'
					disabled={!canEdit}
					className={cn(
						'flex w-full items-center gap-2 px-2 py-1.5 text-left text-[11px] hover:bg-accent disabled:cursor-not-allowed disabled:opacity-40',
						noOutline && 'bg-accent',
					)}
					onClick={() => onUpdateShapeStyle(createNoOutlinePatch())}
				>
					<LuBan className='h-3.5 w-3.5' />
					{t('pptx.shapeOutline.noOutline', 'No Outline')}
				</button>

				<input
					ref={moreColorInputRef}
					type='color'
					value={selectedColor}
					disabled={!canEdit}
					className='hidden'
					onChange={(event) => setColor(event.target.value)}
				/>
				<button
					type='button'
					disabled={!canEdit}
					className='flex w-full items-center gap-2 px-2 py-1.5 text-left text-[11px] hover:bg-accent disabled:cursor-not-allowed disabled:opacity-40'
					onClick={() => moreColorInputRef.current?.click()}
				>
					<span className='h-3.5 w-3.5 rounded-full border border-border bg-[conic-gradient(red,yellow,lime,aqua,blue,magenta,red)]' />
					{t('pptx.shapeOutline.moreColors', 'More Outline Colors…')}
				</button>

				<button
					type='button'
					disabled={!canEdit}
					className='flex w-full items-center gap-2 px-2 py-1.5 text-left text-[11px] hover:bg-accent disabled:cursor-not-allowed disabled:opacity-40'
					onClick={() => void pickFromScreen()}
				>
					<LuPipette className='h-3.5 w-3.5' />
					{t('pptx.shapeOutline.eyedropper', 'Eyedropper')}
				</button>
			</div>

			<label className='grid grid-cols-[5rem,1fr] items-center gap-2 text-[11px]'>
				<span className='text-muted-foreground'>{t('pptx.shapeOutline.weight', 'Weight')}</span>
				<select
					aria-label='Outline weight'
					disabled={!canEdit || noOutline}
					className={INPUT}
					value={String(
						OUTLINE_WEIGHT_POINTS.reduce((nearest, candidate) =>
							Math.abs(candidate - weightPoints) < Math.abs(nearest - weightPoints)
								? candidate
								: nearest,
						),
					)}
					onChange={(event) =>
						onUpdateShapeStyle({ strokeWidth: outlinePointsToPixels(Number(event.target.value)) })
					}
				>
					{OUTLINE_WEIGHT_POINTS.map((points) => (
						<option key={points} value={points}>
							{points} pt
						</option>
					))}
				</select>
			</label>

			<label className='grid grid-cols-[5rem,1fr] items-center gap-2 text-[11px]'>
				<span className='text-muted-foreground'>{t('pptx.shapeOutline.dashes', 'Dashes')}</span>
				<select
					aria-label='Outline dashes'
					disabled={!canEdit || noOutline}
					className={INPUT}
					value={style?.strokeDash === 'custom' ? 'solid' : style?.strokeDash || 'solid'}
					onChange={(event) =>
						onUpdateShapeStyle({ strokeDash: event.target.value as StrokeDashType })
					}
				>
					{COMMON_DASH_OPTIONS.map((option) => (
						<option key={option.value} value={option.value}>
							{t(option.i18nKey)}
						</option>
					))}
				</select>
			</label>

			{showArrows && (
				<div className='grid grid-cols-2 gap-1.5 border-t border-border pt-2 text-[11px]'>
					<ArrowSelect
						label={t('pptx.connectorArrows.startArrow')}
						value={style?.connectorStartArrow || 'none'}
						disabled={!canEdit || noOutline}
						onChange={(connectorStartArrow) => onUpdateShapeStyle({ connectorStartArrow })}
					/>
					<ArrowSelect
						label={t('pptx.connectorArrows.endArrow')}
						value={style?.connectorEndArrow || 'none'}
						disabled={!canEdit || noOutline}
						onChange={(connectorEndArrow) => onUpdateShapeStyle({ connectorEndArrow })}
					/>
					<ArrowSizeSelect
						label={t('pptx.connectorArrows.startWidth')}
						value={style?.connectorStartArrowWidth || 'med'}
						disabled={
							!canEdit ||
							noOutline ||
							!style?.connectorStartArrow ||
							style.connectorStartArrow === 'none'
						}
						onChange={(connectorStartArrowWidth) =>
							onUpdateShapeStyle({ connectorStartArrowWidth })
						}
					/>
					<ArrowSizeSelect
						label={t('pptx.connectorArrows.endWidth')}
						value={style?.connectorEndArrowWidth || 'med'}
						disabled={
							!canEdit ||
							noOutline ||
							!style?.connectorEndArrow ||
							style.connectorEndArrow === 'none'
						}
						onChange={(connectorEndArrowWidth) =>
							onUpdateShapeStyle({ connectorEndArrowWidth })
						}
					/>
				</div>
			)}
		</div>
	);
}

function ArrowSelect({
	label,
	value,
	disabled,
	onChange,
}: {
	label: string;
	value: ConnectorArrowType;
	disabled: boolean;
	onChange: (value: ConnectorArrowType) => void;
}): React.ReactElement {
	const { t } = useTranslation();
	return (
		<label className='flex flex-col gap-1'>
			<span className='text-muted-foreground'>{label}</span>
			<select
				className={INPUT}
				disabled={disabled}
				value={value}
				onChange={(event) => onChange(event.target.value as ConnectorArrowType)}
			>
				{CONNECTOR_ARROW_OPTIONS.map((option) => (
					<option key={option.value} value={option.value}>
						{t(option.i18nKey)}
					</option>
				))}
			</select>
		</label>
	);
}

function ArrowSizeSelect({
	label,
	value,
	disabled,
	onChange,
}: {
	label: string;
	value: 'sm' | 'med' | 'lg';
	disabled: boolean;
	onChange: (value: 'sm' | 'med' | 'lg') => void;
}): React.ReactElement {
	const { t } = useTranslation();
	return (
		<label className='flex flex-col gap-1'>
			<span className='text-muted-foreground'>{label}</span>
			<select
				className={INPUT}
				disabled={disabled}
				value={value}
				onChange={(event) => onChange(event.target.value as 'sm' | 'med' | 'lg')}
			>
				{ARROW_SIZE_OPTIONS.map((option) => (
					<option key={option.value} value={option.value}>
						{t(option.i18nKey)}
					</option>
				))}
			</select>
		</label>
	);
}

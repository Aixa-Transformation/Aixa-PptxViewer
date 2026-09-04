import type { ComponentType, Ref } from 'react';

export type ViewerMode = 'preview' | 'edit' | 'present' | 'master';

export interface PowerPointViewerHandle {
	getContent(): Promise<Uint8Array>;
	setActiveSlideIndex(index: number): void;
	goTo(index: number): void;
	goPrev(): void;
	goNext(): void;
	addSlide(afterIndex?: number): void;
	deleteSlides(indexes: number[]): void;
}

export interface PowerPointViewerProps {
	ref?: Ref<PowerPointViewerHandle>;
	content: Uint8Array;
	canEdit?: boolean;
	singleSlideOnly?: boolean;
	/** Zero-based slide index controlled by the host application. */
	activeSlideIndex?: number;
	className?: string;
	theme?: unknown;
	onContentChange?: (content: Uint8Array) => void;
	onDirtyChange?: (dirty: boolean) => void;
	onDeckWideChange?: (change: { type: 'background' | 'theme' }) => void;
	/** Persist an explicit Apply action; use these bytes or call getContent(). */
	onSaveRequest?: (content?: Uint8Array) => void;
}

export const PowerPointViewer: ComponentType<PowerPointViewerProps>;
export const SlideCanvas: ComponentType<Record<string, unknown>>;
export function useViewerBuildingBlocks(input: Record<string, unknown>): {
	canvasProps: Record<string, unknown>;
	toolbarProps: Record<string, unknown>;
	loading: boolean;
	error?: string | null;
};
export const vermilionDarkTheme: unknown;
export const vermilionLightTheme: unknown;

export class Presentation {
	static load(buffer: ArrayBuffer): Promise<Presentation>;
	readonly slideCount: number;
	removeSlide(index: number): this;
	merge(
		source: Presentation,
		options?: { slideIndices?: number[]; insertAt?: number; keepSourceTheme?: boolean },
	): number;
	save(): Promise<Uint8Array>;
	dispose(): void;
}

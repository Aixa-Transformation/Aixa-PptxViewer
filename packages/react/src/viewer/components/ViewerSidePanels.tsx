/**
 * ViewerSidePanels: Inspector pane, selection pane, theme editor panel,
 * and theme gallery that appear alongside the slide canvas.
 */
import { themeColorSchemesEqual } from 'pptx-viewer-core';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import type { ToolbarActionId } from 'pptx-viewer-shared';
import type { PptxAiBridge, PptxAiConfig } from 'pptx-viewer-shared/ai';
import { flushSync } from 'react-dom';

import { ViewerInspector, SelectionPane } from '.';
import type { AiPanelController } from '../hooks/ai/useAiPanelController';
import type { UseCommentsResult } from '../hooks/useComments-helpers';
import type { EditorHistoryResult } from '../hooks/useEditorHistory';
import type { ElementManipulationHandlers } from '../hooks/useElementManipulation';
import type { ElementOperations } from '../hooks/useElementOperations';
import type { PropertyHandlersResult } from '../hooks/usePropertyHandlers';
import type { ThemeHandlersResult } from '../hooks/useThemeHandlers';
import type { ViewerState } from '../hooks/useViewerState';
import type { CanvasSize } from '../types';
import type { ViewerMode } from '../types-core';
import { AiChatPanelLazy } from './ai';
import { ThemeEditorPanel } from './inspector/ThemeEditorPanel';
import { MobileDismissSheet } from './mobile/MobileDismissSheet';
import { ResizeHandle } from './ResizeHandle';
import type { ThemeDefinition } from './toolbar/ThemeGallery';
import { BUILT_IN_THEMES, ThemeGallery } from './toolbar/ThemeGallery';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

export interface ViewerSidePanelsProps {
	isMobile: boolean;
	hiddenActions?: readonly ToolbarActionId[];
	mode: ViewerMode;
	canEdit: boolean;
	activeSlide: PptxSlide | undefined;
	masterPseudoSlide: PptxSlide | undefined;
	slides: PptxSlide[];
	canvasSize: CanvasSize;
	activeSlideIndex: number;
	selectedElement: PptxElement | null;
	state: ViewerState;
	comments: UseCommentsResult;
	ops: ElementOperations;
	manipulation: ElementManipulationHandlers;
	propertyHandlers: PropertyHandlersResult;
	themeHandlers: ThemeHandlersResult;
	history: EditorHistoryResult;
	/** Width of the right inspector panel in pixels. */
	panelWidth?: number;
	/** Callback to resize the right panel. */
	onResizeRight?: (delta: number) => void;
	/** Reports an operation that changes the appearance of every slide. */
	onDeckWideChange?: (change: { type: 'background' | 'theme' }) => void;
	onSaveRequest?: (content?: Uint8Array) => void;
	/** AI assistant config (present only when the host passes the `ai` prop). */
	aiConfig?: PptxAiConfig;
	/** Bridge exposing the live deck to the AI core. */
	aiBridge?: PptxAiBridge;
	/** AI panel open state + focus/prefill controller (present when `ai` is set). */
	aiPanel?: AiPanelController;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function ViewerSidePanels(props: ViewerSidePanelsProps) {
	const {
		isMobile,
		mode,
		canEdit,
		activeSlide,
		masterPseudoSlide,
		slides,
		canvasSize,
		activeSlideIndex,
		selectedElement,
		state: s,
		comments,
		ops,
		manipulation,
		propertyHandlers,
		themeHandlers,
		history,
		panelWidth,
		onResizeRight,
		onDeckWideChange,
		onSaveRequest,
		aiConfig,
		aiBridge,
		aiPanel,
		hiddenActions,
	} = props;

	const effectiveSlide = mode === 'master' ? masterPseudoSlide : activeSlide;
	const currentBuiltInTheme =
		BUILT_IN_THEMES.find((candidate) =>
			themeColorSchemesEqual(s.theme?.colorScheme, candidate.colorScheme),
		) ?? null;

	return (
		<>
			{!isMobile &&
				(mode === 'edit' || mode === 'master') &&
				s.isInspectorPaneOpen &&
				onResizeRight && <ResizeHandle direction='horizontal' onResize={onResizeRight} />}
			<ViewerInspector
				isMobile={isMobile}
				hiddenActions={hiddenActions}
				isOpen={(mode === 'edit' || mode === 'master') && s.isInspectorPaneOpen}
				canEdit={canEdit}
				mode={mode}
				activeSlide={effectiveSlide}
				slides={slides}
				canvasSize={canvasSize}
				selectedElement={selectedElement}
				effectiveSelectedIds={s.effectiveSelectedIds}
				tableEditorState={s.tableEditorState}
				sidebarPanelMode={s.sidebarPanelMode}
				activeSlideIndex={activeSlideIndex}
				comments={comments}
				onSetSidebarPanelMode={s.setSidebarPanelMode}
				onClose={() => s.setIsInspectorPaneOpen(false)}
				onUpdateElementStyle={ops.updateSelectedShapeStyle}
				onUpdateTextStyle={ops.updateSelectedTextStyle}
				onUpdateElement={ops.updateSelectedElement}
				onApplySelection={ops.applySelection}
				onSetCanvasSize={s.setCanvasSize}
				onMoveLayer={manipulation.handleMoveLayer}
				onMoveLayerToEdge={manipulation.handleMoveLayerToEdge}
				onDeleteElement={manipulation.handleDelete}
				onUpdateSlide={propertyHandlers.handleUpdateSlide}
				onApplySlideBackground={(updates, allSlides) => {
					if (!canEdit) return;
					// Hosts may serialize immediately in onSaveRequest; commit the slide state first.
					flushSync(() => {
						if (allSlides) propertyHandlers.handleUpdateAllSlidesBackground(updates);
						else propertyHandlers.handleUpdateSlide(updates);
					});
					if (allSlides) onDeckWideChange?.({ type: 'background' });
					onSaveRequest?.();
				}}
				presentationProperties={s.presentationProperties}
				onUpdatePresentationProperties={propertyHandlers.handleUpdatePresentationProperties}
				editTemplateMode={s.editTemplateMode}
				slideMasters={s.slideMasters}
				themeOptions={s.themeOptions}
				notesMaster={s.notesMaster}
				handoutMaster={s.handoutMaster}
				notesCanvasSize={s.notesCanvasSize}
				coreProperties={s.coreProperties}
				appProperties={s.appProperties}
				customProperties={s.customProperties}
				tagCollections={s.tagCollections}
				onUpdateTagCollections={s.setTagCollections}
				onUpdateCoreProperties={propertyHandlers.handleUpdateCoreProperties}
				onUpdateAppProperties={propertyHandlers.handleUpdateAppProperties}
				onUpdateCustomProperties={propertyHandlers.handleUpdateCustomProperties}
				onApplyTheme={async (path, allSlides) => {
					if (!canEdit) return;
					const updated = await themeHandlers.handleApplyTheme(path, allSlides);
					if (!updated) return;
					if (allSlides) onDeckWideChange?.({ type: 'theme' });
					onSaveRequest?.(updated);
				}}
				onSetTemplateBackground={themeHandlers.handleSetTemplateBackground}
				onGetTemplateBackgroundColor={themeHandlers.handleGetTemplateBackgroundColor}
				mediaDataUrls={s.mediaDataUrls}
				theme={s.theme}
				panelWidth={panelWidth}
			/>

			{s.isSelectionPaneOpen && (mode === 'edit' || mode === 'master') && (
				<MobileDismissSheet
					onClose={() => s.setIsSelectionPaneOpen(false)}
					className='absolute right-0 top-0 z-30 h-full          '
				>
					<SelectionPane
						slides={slides}
						activeSlideIndex={activeSlideIndex}
						selectedElementId={s.selectedElementId}
						selectedElementIds={s.selectedElementIds}
						canEdit={canEdit}
						setSelectedElementId={s.setSelectedElementId}
						setSelectedElementIds={s.setSelectedElementIds}
						setSlides={s.setSlides}
						markDirty={history.markDirty}
						onClose={() => s.setIsSelectionPaneOpen(false)}
					/>
				</MobileDismissSheet>
			)}

			{s.isThemeEditorOpen && mode === 'edit' && (
				<MobileDismissSheet
					onClose={() => s.setIsThemeEditorOpen(false)}
					className='absolute right-0 top-0 z-30 h-full w-72 overflow-y-auto border-l border-border bg-card p-2.5 shadow-xl         '
				>
					<ThemeEditorPanel
						theme={s.theme}
						canEdit={canEdit}
						onUpdateColorScheme={themeHandlers.handleUpdateThemeColorScheme}
						onUpdateFontScheme={themeHandlers.handleUpdateThemeFontScheme}
						onUpdateThemeName={themeHandlers.handleUpdateThemeName}
						onApplyToPresentation={themeHandlers.handleApplyThemeToPresentation}
						onClose={() => s.setIsThemeEditorOpen(false)}
					/>
				</MobileDismissSheet>
			)}

			{aiPanel?.isOpen && aiConfig && aiBridge && (mode === 'edit' || mode === 'master') && (
				<AiChatPanelLazy
					bridge={aiBridge}
					config={aiConfig}
					aiPanel={aiPanel}
					onClose={aiPanel.close}
					panelWidth={panelWidth}
				/>
			)}

			<ThemeGallery
				open={s.isThemeGalleryOpen}
				currentTheme={currentBuiltInTheme}
				canEdit={canEdit}
				onClose={() => s.setIsThemeGalleryOpen(false)}
				onApplyTheme={(theme: ThemeDefinition) => {
					themeHandlers.handleApplyThemeData(
						theme.colorScheme,
						{
							majorFont: {
								latin: theme.fontScheme.majorFont,
								eastAsia: theme.fontScheme.majorFont,
								complexScript: theme.fontScheme.majorFont,
							},
							minorFont: {
								latin: theme.fontScheme.minorFont,
								eastAsia: theme.fontScheme.minorFont,
								complexScript: theme.fontScheme.minorFont,
							},
						},
						theme.name,
					);
				}}
			/>
		</>
	);
}

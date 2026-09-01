/**
 * ViewerInspector: Wrapper that wires InspectorPane props.
 *
 * Accepts a simplified set of props and derives the verbose prop
 * bag required by InspectorPane internally, keeping the main
 * orchestrator lean.
 */
import type {
	PptxAppProperties,
	PptxCoreProperties,
	PptxCustomProperty,
	PptxElement,
	PptxHandoutMaster,
	PptxNotesMaster,
	PptxSlide,
	PptxSlideMaster,
	PptxPresentationProperties,
	PptxTagCollection,
	PptxTheme,
	PptxThemeOption,
	ShapeStyle,
	TextStyle,
} from 'pptx-viewer-core';
import { isActionHidden, type ToolbarActionId } from 'pptx-viewer-shared';
import { useTranslation } from 'react-i18next';

import type { UseCommentsResult } from '../hooks/useComments-helpers';
import type { CanvasSize, ViewerMode, TableCellEditorState } from '../types';
import { cn } from '../utils';
import type { InspectorTab } from './inspector/inspector-pane-types';
import { InspectorPane } from './InspectorPane';

/* ------------------------------------------------------------------ */
/*  Props                                                             */
/* ------------------------------------------------------------------ */

export interface ViewerInspectorProps {
	isMobile: boolean;
	hiddenActions?: readonly ToolbarActionId[];
	isOpen: boolean;
	canEdit: boolean;
	mode: ViewerMode;
	activeSlide: PptxSlide | undefined;
	slides: PptxSlide[];
	canvasSize: CanvasSize;
	selectedElement: PptxElement | null;
	effectiveSelectedIds: string[];
	tableEditorState?: TableCellEditorState | null;
	sidebarPanelMode: string;
	activeSlideIndex: number;
	comments: UseCommentsResult;
	onSetSidebarPanelMode: (mode: string) => void;
	onClose: () => void;
	onUpdateElementStyle: (updates: Partial<ShapeStyle>) => void;
	onUpdateTextStyle: (updates: Partial<TextStyle>) => void;
	onUpdateElement: (updates: Partial<PptxElement>) => void;
	onApplySelection: (id: string | null) => void;
	onSetCanvasSize: React.Dispatch<React.SetStateAction<CanvasSize>>;
	onMoveLayer: (direction: string) => void;
	onMoveLayerToEdge: (direction: string) => void;
	onDeleteElement: () => void;
	onUpdateSlide: (updates: Partial<PptxSlide>) => void;
	onUpdateAllSlidesBackground: (
		updates: Pick<PptxSlide, 'backgroundColor' | 'backgroundImage' | 'backgroundGradient'>,
	) => void;
	presentationProperties: PptxPresentationProperties;
	onUpdatePresentationProperties: (updates: Partial<PptxPresentationProperties>) => void;
	notesMaster?: PptxNotesMaster;
	handoutMaster?: PptxHandoutMaster;
	notesCanvasSize?: CanvasSize;
	coreProperties?: PptxCoreProperties;
	appProperties?: PptxAppProperties;
	customProperties: PptxCustomProperty[];
	themeOptions: PptxThemeOption[];
	onUpdateCoreProperties: (updates: Partial<PptxCoreProperties>) => void;
	onUpdateAppProperties: (updates: Partial<PptxAppProperties>) => void;
	onUpdateCustomProperties: (next: PptxCustomProperty[]) => void;
	tagCollections?: PptxTagCollection[];
	onUpdateTagCollections?: (next: PptxTagCollection[]) => void;
	onApplyTheme: (themePath: string, applyToAllMasters: boolean) => void;
	editTemplateMode?: boolean;
	slideMasters?: PptxSlideMaster[];
	onSetTemplateBackground?: (path: string, backgroundColor: string) => void;
	onGetTemplateBackgroundColor?: (path: string) => string | undefined;
	mediaDataUrls?: Map<string, string>;
	theme?: PptxTheme;
	/** Width of the panel in pixels (for resizable panels). */
	panelWidth?: number;
}

/* ------------------------------------------------------------------ */
/*  Component                                                         */
/* ------------------------------------------------------------------ */

export function ViewerInspector({
	isMobile,
	isOpen,
	canEdit,
	mode,
	activeSlide,
	slides,
	canvasSize,
	selectedElement,
	effectiveSelectedIds,
	tableEditorState,
	sidebarPanelMode,
	activeSlideIndex,
	comments,
	onSetSidebarPanelMode,
	onClose,
	onUpdateElementStyle,
	onUpdateTextStyle,
	onUpdateElement,
	onApplySelection,
	onSetCanvasSize,
	onMoveLayer,
	onMoveLayerToEdge,
	onDeleteElement,
	onUpdateSlide,
	onUpdateAllSlidesBackground,
	presentationProperties,
	onUpdatePresentationProperties,
	notesMaster,
	handoutMaster,
	notesCanvasSize,
	coreProperties,
	appProperties,
	customProperties,
	themeOptions,
	onUpdateCoreProperties,
	onUpdateAppProperties,
	onUpdateCustomProperties,
	tagCollections,
	onUpdateTagCollections,
	onApplyTheme,
	editTemplateMode,
	slideMasters,
	onSetTemplateBackground,
	onGetTemplateBackgroundColor,
	mediaDataUrls,
	theme,
	panelWidth,
	hiddenActions,
}: ViewerInspectorProps): React.ReactElement | null {
	const { t } = useTranslation();

	if (!isOpen) {
		return null;
	}

	// Allow the comments tab to render even when no element is selected,
	// because comments belong to the slide rather than a single element.
	const requestedTab = sidebarPanelMode as InspectorTab;
	const tab =
		(requestedTab === 'elements' && isActionHidden('inspectorElements', hiddenActions)) ||
		(requestedTab === 'comments' && isActionHidden('inspectorComments', hiddenActions))
			? 'properties'
			: requestedTab;
	if (!selectedElement && tab !== 'comments' && tab !== 'properties') {
		return null;
	}

	const slideId = activeSlide?.id ?? '';
	const commentDraft = comments.commentDraftBySlideId[slideId] ?? '';
	const editingCommentId = comments.editingCommentIdBySlideId[slideId] ?? null;

	return (
		<div
			data-pptx-inspector
			role='complementary'
			aria-label={t('pptx.inspector.properties')}
			className={cn(
				'h-full flex-shrink-0',
				isMobile && 'absolute inset-x-0 bottom-0 top-auto z-30',
			)}
			style={!isMobile && panelWidth ? { width: panelWidth } : undefined}
		>
			<InspectorPane
				isMobile={isMobile}
				hiddenActions={hiddenActions}
				isOpen={isOpen}
				canEdit={canEdit}
				mode={mode}
				activeSlide={activeSlide}
				slides={slides}
				canvasSize={canvasSize}
				selectedElement={selectedElement}
				selectedElementIds={effectiveSelectedIds}
				tableEditorState={tableEditorState}
				activeTab={tab}
				onSetActiveTab={(nextTab) => onSetSidebarPanelMode(nextTab)}
				onClose={onClose}
				onUpdateElementStyle={onUpdateElementStyle}
				onUpdateTextStyle={onUpdateTextStyle}
				onUpdateElement={onUpdateElement}
				onUpdateSlide={onUpdateSlide}
				onUpdateAllSlidesBackground={onUpdateAllSlidesBackground}
				onSelectElement={(id) => onApplySelection(id)}
				onMoveLayer={onMoveLayer}
				onMoveLayerToEdge={onMoveLayerToEdge}
				onDeleteElement={onDeleteElement}
				comments={activeSlide?.comments ?? []}
				presentationProperties={presentationProperties}
				onUpdatePresentationProperties={onUpdatePresentationProperties}
				notesMaster={notesMaster}
				handoutMaster={handoutMaster}
				notesCanvasSize={notesCanvasSize}
				coreProperties={coreProperties}
				appProperties={appProperties}
				customProperties={customProperties}
				themeOptions={themeOptions}
				onUpdateCoreProperties={onUpdateCoreProperties}
				onUpdateAppProperties={onUpdateAppProperties}
				onUpdateCustomProperties={onUpdateCustomProperties}
				tagCollections={tagCollections}
				onUpdateTagCollections={onUpdateTagCollections}
				onApplyTheme={onApplyTheme}
				commentDraft={commentDraft}
				editingCommentId={editingCommentId}
				commentEditDraft={
					editingCommentId ? (comments.commentEditDraftByCommentId[editingCommentId] ?? '') : ''
				}
				onSetCommentDraft={(draft) => comments.handleCommentDraftChange(slideId, draft)}
				onAddComment={() => comments.handleAddSlideComment(activeSlideIndex)}
				onDeleteComment={(id) => comments.handleDeleteSlideComment(activeSlideIndex, id)}
				onStartEditComment={(id) => comments.handleStartCommentEdit(slideId, id)}
				onSaveEditComment={(id) => comments.handleSaveCommentEdit(activeSlideIndex, id)}
				onCancelEditComment={() => comments.handleCancelCommentEdit(slideId)}
				onSetCommentEditDraft={(draft) =>
					comments.handleSetCommentEditDraft(
						comments.editingCommentIdBySlideId[slideId] ?? '',
						draft,
					)
				}
				onToggleCommentResolved={(id) => comments.handleToggleCommentResolved(activeSlideIndex, id)}
				onStartReply={(id) => comments.handleStartReply(activeSlideIndex, id)}
				onCancelReply={comments.handleCancelReply}
				onReplyDraftChange={(commentId, draft) => comments.handleReplyDraftChange(commentId, draft)}
				onSubmitReply={(commentId) => comments.handleSubmitReply(activeSlideIndex, commentId)}
				replyingToCommentId={comments.replyingToCommentId}
				replyDraftByCommentId={comments.replyDraftByCommentId}
				onUpdateCanvasSize={onSetCanvasSize}
				editTemplateMode={editTemplateMode}
				slideMasters={slideMasters}
				onSetTemplateBackground={onSetTemplateBackground}
				onGetTemplateBackgroundColor={onGetTemplateBackgroundColor}
				mediaDataUrls={mediaDataUrls}
				theme={theme}
				panelWidth={panelWidth}
			/>
		</div>
	);
}

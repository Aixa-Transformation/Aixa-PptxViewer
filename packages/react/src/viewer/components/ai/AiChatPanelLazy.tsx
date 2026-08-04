import type { PptxAiBridge, PptxAiConfig } from 'pptx-viewer-shared/ai';
import type { ReactElement } from 'react';

import type { AiPanelController } from '../../hooks/ai/useAiPanelController';

/** Optional AI integration point. AI is disabled in this host application. */
export interface AiChatPanelLazyProps {
  bridge: PptxAiBridge;
  config: PptxAiConfig;
  aiPanel: AiPanelController;
  onClose: () => void;
  panelWidth?: number;
}

export function AiChatPanelLazy(_props: AiChatPanelLazyProps): ReactElement | null {
  return null;
}

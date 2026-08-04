import type { PptxAiChatStore } from 'pptx-viewer-shared/ai';
import type { ReactElement } from 'react';

export interface SettingsAiTabProps {
  store?: PptxAiChatStore;
}

/** AI settings are not enabled in this host application. */
export function SettingsAiTab(_props: SettingsAiTabProps): ReactElement | null {
  return null;
}

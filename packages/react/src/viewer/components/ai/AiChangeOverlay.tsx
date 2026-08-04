import type { ReactElement } from 'react';
import type { AiChangeBatch } from 'pptx-viewer-shared/ai';

export interface AiChangeOverlayProps {
  batch: AiChangeBatch | null;
  activeSlideIndex: number;
}

/** AI editing is disabled in this host application. */
export function AiChangeOverlay(_props: AiChangeOverlayProps): ReactElement | null {
  return null;
}

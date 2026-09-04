import type React from 'react';

export interface MobileDismissSheetProps {
	/** Retained for API compatibility with the former responsive sheet. */
	onClose: () => void;
	/** Classes for the desktop panel container. */
	className?: string;
	children: React.ReactNode;
}

/**
 * Desktop-only panel wrapper. It intentionally adds no backdrop, blur, or
 * swipe gesture on narrow screens.
 */
export function MobileDismissSheet({
	onClose: _onClose,
	className,
	children,
}: MobileDismissSheetProps): React.ReactElement {
	return <div className={className}>{children}</div>;
}

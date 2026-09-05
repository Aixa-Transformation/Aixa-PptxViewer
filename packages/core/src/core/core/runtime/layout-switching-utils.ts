/** Template artwork is inherited, not slide-owned, and must retain identity. */
export function isInheritedTemplateElementId(elementId: string): boolean {
	return elementId.startsWith('layout-') || elementId.startsWith('master-');
}

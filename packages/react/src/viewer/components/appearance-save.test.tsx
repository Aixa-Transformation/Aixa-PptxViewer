import { describe, expect, it, vi } from 'vitest';
import type { ReactElement } from 'react';
import { ViewerSidePanels, type ViewerSidePanelsProps } from './ViewerSidePanels';
import { ViewerInspector } from '.';

vi.mock('.', () => ({ ViewerInspector: vi.fn(), SelectionPane: vi.fn() }));

function harness(canEdit = true) {
	let revision = 0;
	const current = vi.fn(() => {
		revision++;
	});
	const all = vi.fn(() => {
		revision++;
	});
	const deckChange = vi.fn();
	const save = vi.fn(() => {
		expect(revision).toBeGreaterThan(0);
	});
	const theme = vi.fn(async () => {
		revision++;
		return new Uint8Array([1, 2, 3]);
	});
	const props = {
		isMobile: false,
		mode: 'edit',
		canEdit,
		slides: [],
		activeSlideIndex: 1,
		state: { isInspectorPaneOpen: true },
		comments: {},
		ops: {},
		manipulation: {},
		history: {},
		propertyHandlers: { handleUpdateSlide: current, handleUpdateAllSlidesBackground: all },
		themeHandlers: { handleApplyTheme: theme },
		onDeckWideChange: deckChange,
		onSaveRequest: save,
	} as unknown as ViewerSidePanelsProps;
	const tree = ViewerSidePanels(props);
	const inspector = (tree.props.children as ReactElement[]).find(
		(child) => child?.type === ViewerInspector,
	)!;
	const callbacks = inspector.props as {
		onApplySlideBackground: (patch: object, all: boolean) => void;
		onApplyTheme: (path: string, all: boolean) => Promise<void>;
	};
	return { callbacks, current, all, deckChange, save, theme };
}

describe('Apply save coordination', () => {
	it('commits current background before save and saves every explicit click even when already dirty', () => {
		const { callbacks, current, all, save, deckChange } = harness();
		callbacks.onApplySlideBackground({ backgroundColor: '#123456' }, false);
		callbacks.onApplySlideBackground({ backgroundColor: '#123456' }, false);
		expect(current).toHaveBeenCalledTimes(2);
		expect(all).not.toHaveBeenCalled();
		expect(save).toHaveBeenCalledTimes(2);
		expect(deckChange).not.toHaveBeenCalled();
	});
	it('notifies all-slide invalidation before requesting save', () => {
		const { callbacks, current, all, save, deckChange } = harness();
		callbacks.onApplySlideBackground({ backgroundImage: 'data:image/png;base64,AA' }, true);
		expect(current).not.toHaveBeenCalled();
		expect(all).toHaveBeenCalledOnce();
		expect(deckChange).toHaveBeenCalledWith({ type: 'background' });
		expect(deckChange.mock.invocationCallOrder[0]).toBeLessThan(save.mock.invocationCallOrder[0]);
	});
	it('delivers the completed theme package, with all-slide notification only for all', async () => {
		const { callbacks, theme, save, deckChange } = harness();
		await callbacks.onApplyTheme('theme2', false);
		expect(theme).toHaveBeenLastCalledWith('theme2', false);
		expect(save).toHaveBeenCalledWith(new Uint8Array([1, 2, 3]));
		expect(deckChange).not.toHaveBeenCalled();
		await callbacks.onApplyTheme('theme2', true);
		expect(deckChange).toHaveBeenCalledWith({ type: 'theme' });
	});
	it('never requests save for a failed theme Apply or a read-only viewer', async () => {
		const editable = harness();
		editable.theme.mockRejectedValueOnce(new Error('failed'));
		await expect(editable.callbacks.onApplyTheme('theme2', true)).rejects.toThrow('failed');
		expect(editable.save).not.toHaveBeenCalled();
		expect(editable.deckChange).not.toHaveBeenCalled();
		const readonly = harness(false);
		readonly.callbacks.onApplySlideBackground({}, true);
		await readonly.callbacks.onApplyTheme('theme2', true);
		expect(readonly.save).not.toHaveBeenCalled();
		expect(readonly.theme).not.toHaveBeenCalled();
	});
});

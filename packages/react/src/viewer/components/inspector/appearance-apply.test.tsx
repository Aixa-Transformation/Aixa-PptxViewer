// @vitest-environment happy-dom
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { SlideBackgroundPanel } from './SlideBackgroundPanel';
import { ThemeSelectorCard } from './PresentationSettingsCards';

vi.mock('react-i18next', () => ({
	useTranslation: () => ({
		t: (key: string) =>
			({
				'pptx.documentProperties.applyCurrentSlide': 'Apply to Current Slide',
				'pptx.documentProperties.applyAllSlides': 'Apply to All Slides',
			})[key] ?? key,
	}),
}));

const roots: ReturnType<typeof createRoot>[] = [];
afterEach(() => {
	act(() => roots.splice(0).forEach((root) => root.unmount()));
	document.body.innerHTML = '';
});
function mount(node: React.ReactNode) {
	const container = document.createElement('div');
	document.body.append(container);
	const root = createRoot(container);
	roots.push(root);
	act(() => root.render(node));
	return container;
}
async function click(container: HTMLElement, text: string) {
	const button = [...container.querySelectorAll('button')].find((b) => b.textContent === text)!;
	expect(button).toBeDefined();
	await act(async () => button.click());
}

describe('appearance Apply buttons', () => {
	it('background applies current/all, including image and graphics visibility, without editing masters', async () => {
		const apply = vi.fn();
		const master = vi.fn();
		const update = vi.fn();
		const panel = mount(
			<SlideBackgroundPanel
				activeSlide={{
					id: 'second',
					elements: [],
					backgroundImage: 'data:image/png;base64,AAAA',
					showMasterShapes: false,
				}}
				canEdit
				onUpdateSlide={update}
				onApplySlideBackground={apply}
				onSetTemplateBackground={master}
			/>,
		);
		await click(panel, 'Apply to Current Slide');
		expect(apply).toHaveBeenLastCalledWith(
			expect.objectContaining({
				backgroundImage: 'data:image/png;base64,AAAA',
				showMasterShapes: false,
			}),
			false,
		);
		await click(panel, 'Apply to All Slides');
		expect(apply).toHaveBeenLastCalledWith(expect.anything(), true);
		expect(master).not.toHaveBeenCalled();
		expect(update).not.toHaveBeenCalled();
	});
	it('theme applies the chosen theme using slide scope; read-only controls cannot apply', async () => {
		const apply = vi.fn();
		const props = {
			themeOptions: [{ name: 'Theme', path: 'ppt/theme/theme2.xml' }],
			selectedThemePath: 'ppt/theme/theme2.xml',
			setSelectedThemePath: vi.fn(),
			onApplyTheme: apply,
		};
		const panel = mount(<ThemeSelectorCard {...props} canEdit />);
		await click(panel, 'Apply to Current Slide');
		expect(apply).toHaveBeenLastCalledWith('ppt/theme/theme2.xml', false);
		await click(panel, 'Apply to All Slides');
		expect(apply).toHaveBeenLastCalledWith('ppt/theme/theme2.xml', true);
		apply.mockClear();
		const readonly = mount(<ThemeSelectorCard {...props} canEdit={false} />);
		await click(readonly, 'Apply to Current Slide');
		await click(readonly, 'Apply to All Slides');
		expect(apply).not.toHaveBeenCalled();
	});
	it('shows theme errors and permits retry', async () => {
		const apply = vi.fn().mockRejectedValue(new Error('Theme unavailable'));
		const panel = mount(
			<ThemeSelectorCard
				themeOptions={[]}
				selectedThemePath="missing"
				setSelectedThemePath={vi.fn()}
				onApplyTheme={apply}
				canEdit
			/>,
		);
		await click(panel, 'Apply to Current Slide');
		expect(panel.querySelector('[role="alert"]')?.textContent).toBe('Theme unavailable');
		apply.mockResolvedValue(undefined);
		await click(panel, 'Apply to All Slides');
		expect(panel.querySelector('[role="alert"]')).toBeNull();
	});
});

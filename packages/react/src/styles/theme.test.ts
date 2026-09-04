import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

describe('PowerPoint canvas typography isolation', () => {
	it('disables browser text autosizing at the viewer root', () => {
		const css = readFileSync(new URL('./theme.css', import.meta.url), 'utf8');
		const viewerRule = css.match(/\[data-pptx-viewer\]\s*\{(?<body>[^}]*)\}/u);

		expect(viewerRule?.groups?.body).toContain('-webkit-text-size-adjust: none !important');
		expect(viewerRule?.groups?.body).toContain('text-size-adjust: none !important');
	});
});

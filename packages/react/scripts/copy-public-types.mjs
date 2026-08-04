import { cp, mkdir } from 'node:fs/promises';

await mkdir(new URL('../dist/viewer/', import.meta.url), { recursive: true });
await Promise.all([
  cp(new URL('../public-types/index.d.ts', import.meta.url), new URL('../dist/index.d.ts', import.meta.url)),
  cp(new URL('../public-types/i18n.d.ts', import.meta.url), new URL('../dist/i18n.d.ts', import.meta.url)),
  cp(new URL('../public-types/internals.d.ts', import.meta.url), new URL('../dist/internals.d.ts', import.meta.url)),
  cp(
    new URL('../public-types/presentation.d.ts', import.meta.url),
    new URL('../dist/presentation.d.ts', import.meta.url)
  ),
  cp(
    new URL('../public-types/viewer/index.d.ts', import.meta.url),
    new URL('../dist/viewer/index.d.ts', import.meta.url)
  ),
]);

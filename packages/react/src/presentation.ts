// Browser-safe deck manipulation entry point. Keeping this separate from the
// React viewer entry prevents consumers that only need Presentation from
// loading UI chunks and their CommonJS compatibility shims.
export { Presentation } from 'pptx-viewer-core';

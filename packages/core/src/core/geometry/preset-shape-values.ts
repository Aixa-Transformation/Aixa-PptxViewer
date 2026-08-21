/**
 * DrawingML `ST_ShapeType` values accepted by `a:prstGeom/@prst`.
 *
 * Keep this list aligned with ECMA-376 and Open XML SDK's
 * `DocumentFormat.OpenXml.Drawing.ShapeTypeValues`. Internal editor sentinels
 * such as `custom` must never be serialized as preset geometry.
 */
export const DRAWINGML_PRESET_SHAPE_VALUES: ReadonlySet<string> = new Set([
	'line',
	'lineInv',
	'triangle',
	'rtTriangle',
	'rect',
	'diamond',
	'parallelogram',
	'trapezoid',
	'nonIsoscelesTrapezoid',
	'pentagon',
	'hexagon',
	'heptagon',
	'octagon',
	'decagon',
	'dodecagon',
	'star4',
	'star5',
	'star6',
	'star7',
	'star8',
	'star10',
	'star12',
	'star16',
	'star24',
	'star32',
	'roundRect',
	'round1Rect',
	'round2SameRect',
	'round2DiagRect',
	'snipRoundRect',
	'snip1Rect',
	'snip2SameRect',
	'snip2DiagRect',
	'plaque',
	'ellipse',
	'teardrop',
	'homePlate',
	'chevron',
	'pieWedge',
	'pie',
	'blockArc',
	'donut',
	'noSmoking',
	'rightArrow',
	'leftArrow',
	'upArrow',
	'downArrow',
	'stripedRightArrow',
	'notchedRightArrow',
	'bentUpArrow',
	'leftRightArrow',
	'upDownArrow',
	'leftUpArrow',
	'leftRightUpArrow',
	'quadArrow',
	'leftArrowCallout',
	'rightArrowCallout',
	'upArrowCallout',
	'downArrowCallout',
	'leftRightArrowCallout',
	'upDownArrowCallout',
	'quadArrowCallout',
	'bentArrow',
	'uturnArrow',
	'circularArrow',
	'leftCircularArrow',
	'leftRightCircularArrow',
	'curvedRightArrow',
	'curvedLeftArrow',
	'curvedUpArrow',
	'curvedDownArrow',
	'swooshArrow',
	'cube',
	'can',
	'lightningBolt',
	'heart',
	'sun',
	'moon',
	'smileyFace',
	'irregularSeal1',
	'irregularSeal2',
	'foldedCorner',
	'bevel',
	'frame',
	'halfFrame',
	'corner',
	'diagStripe',
	'chord',
	'arc',
	'leftBracket',
	'rightBracket',
	'leftBrace',
	'rightBrace',
	'bracketPair',
	'bracePair',
	'straightConnector1',
	'bentConnector2',
	'bentConnector3',
	'bentConnector4',
	'bentConnector5',
	'curvedConnector2',
	'curvedConnector3',
	'curvedConnector4',
	'curvedConnector5',
	'callout1',
	'callout2',
	'callout3',
	'accentCallout1',
	'accentCallout2',
	'accentCallout3',
	'borderCallout1',
	'borderCallout2',
	'borderCallout3',
	'accentBorderCallout1',
	'accentBorderCallout2',
	'accentBorderCallout3',
	'wedgeRectCallout',
	'wedgeRoundRectCallout',
	'wedgeEllipseCallout',
	'cloudCallout',
	'cloud',
	'ribbon',
	'ribbon2',
	'ellipseRibbon',
	'ellipseRibbon2',
	'leftRightRibbon',
	'verticalScroll',
	'horizontalScroll',
	'wave',
	'doubleWave',
	'plus',
	'flowChartProcess',
	'flowChartDecision',
	'flowChartInputOutput',
	'flowChartPredefinedProcess',
	'flowChartInternalStorage',
	'flowChartDocument',
	'flowChartMultidocument',
	'flowChartTerminator',
	'flowChartPreparation',
	'flowChartManualInput',
	'flowChartManualOperation',
	'flowChartConnector',
	'flowChartPunchedCard',
	'flowChartPunchedTape',
	'flowChartSummingJunction',
	'flowChartOr',
	'flowChartCollate',
	'flowChartSort',
	'flowChartExtract',
	'flowChartMerge',
	'flowChartOfflineStorage',
	'flowChartOnlineStorage',
	'flowChartMagneticTape',
	'flowChartMagneticDisk',
	'flowChartMagneticDrum',
	'flowChartDisplay',
	'flowChartDelay',
	'flowChartAlternateProcess',
	'flowChartOffpageConnector',
	'actionButtonBlank',
	'actionButtonHome',
	'actionButtonHelp',
	'actionButtonInformation',
	'actionButtonForwardNext',
	'actionButtonBackPrevious',
	'actionButtonEnd',
	'actionButtonBeginning',
	'actionButtonReturn',
	'actionButtonDocument',
	'actionButtonSound',
	'actionButtonMovie',
	'gear6',
	'gear9',
	'funnel',
	'mathPlus',
	'mathMinus',
	'mathMultiply',
	'mathDivide',
	'mathEqual',
	'mathNotEqual',
	'cornerTabs',
	'squareTabs',
	'plaqueTabs',
	'chartX',
	'chartStar',
	'chartPlus',
]);

/** Editor-friendly aliases whose OOXML names differ from their UI names. */
const PRESET_SHAPE_ALIASES: Readonly<Record<string, string>> = {
	cylinder: 'can',
	rightTriangle: 'rtTriangle',
	cross: 'plus',
	flowChartData: 'flowChartInputOutput',
	flowChartDirectData: 'flowChartMagneticDisk',
	flowChartSequentialAccessStorage: 'flowChartMagneticTape',
	flowChartStoredData: 'flowChartOfflineStorage',
	actionButtonBackOrPrevious: 'actionButtonBackPrevious',
	actionButtonForwardOrNext: 'actionButtonForwardNext',
};

/**
 * Return a schema-valid preset name, or `rect` for sentinels and unknown data.
 */
export function normalizeDrawingmlPresetShape(shapeType: string | undefined): string {
	const value = shapeType?.trim();
	if (!value) {
		return 'rect';
	}
	const normalized = PRESET_SHAPE_ALIASES[value] ?? value;
	return DRAWINGML_PRESET_SHAPE_VALUES.has(normalized) ? normalized : 'rect';
}

/**
 * Repair preset geometry attributes anywhere inside a parsed DrawingML tree.
 * This is primarily used for preserved group `rawXml` created by older
 * versions, where nested children do not pass through the normal element
 * writer during a save.
 */
export function sanitizeDrawingmlPresetGeometryTree(value: unknown): void {
	if (!value || typeof value !== 'object') {
		return;
	}
	if (Array.isArray(value)) {
		for (const item of value) {
			sanitizeDrawingmlPresetGeometryTree(item);
		}
		return;
	}
	const node = value as Record<string, unknown>;
	for (const [key, child] of Object.entries(node)) {
		const localName = key.slice(key.indexOf(':') + 1).split('#pptx-order-', 1)[0];
		if (localName === 'prstGeom' && child && typeof child === 'object' && !Array.isArray(child)) {
			const geometry = child as Record<string, unknown>;
			const preset = typeof geometry['@_prst'] === 'string' ? geometry['@_prst'] : undefined;
			geometry['@_prst'] = normalizeDrawingmlPresetShape(preset);
		}
		sanitizeDrawingmlPresetGeometryTree(child);
	}
}

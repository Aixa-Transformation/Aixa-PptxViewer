import type { XmlObject } from '../../types';
import { xmlAttr, xmlChild, xmlPath } from '../../utils/xml-access';

/**
 * Return true when a layout placeholder contributes visible artwork of its
 * own, rather than only text-style/geometry defaults for a slide placeholder.
 */
export function shouldRenderLayoutPlaceholderArtwork(shape: XmlObject): boolean {
	const ph = xmlPath(shape, 'p:nvSpPr', 'p:nvPr', 'p:ph');
	if (!ph) {
		return false;
	}
	const type = String(xmlAttr(ph, 'type') ?? 'obj')
		.trim()
		.toLowerCase();
	if (type === 'pic' || type === 'sldimg') {
		return true;
	}

	const spPr = xmlChild(shape, 'p:spPr');
	if (!spPr) {
		return false;
	}
	if (
		xmlChild(spPr, 'a:blipFill') ||
		xmlChild(spPr, 'a:solidFill') ||
		xmlChild(spPr, 'a:gradFill') ||
		xmlChild(spPr, 'a:pattFill') ||
		xmlChild(spPr, 'a:grpFill') ||
		xmlChild(spPr, 'a:effectLst') ||
		xmlChild(spPr, 'a:effectDag')
	) {
		return true;
	}
	const line = xmlChild(spPr, 'a:ln');
	return Boolean(line && !xmlChild(line, 'a:noFill'));
}

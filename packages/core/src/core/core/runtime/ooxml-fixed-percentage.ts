/**
 * Parse an OOXML ST_Percentage value into its fixed 1/1000-percent integer.
 *
 * Transitional packages commonly store 20% as `20000`, while Strict OOXML
 * may use the lexical form `20%`. The in-memory colour model consistently
 * uses the former 0..100000 representation.
 */
export function parseOoxmlFixedPercentage(value: unknown): number | undefined {
	const raw = String(value ?? '').trim();
	if (!raw) {
		return undefined;
	}
	const parsed = raw.endsWith('%')
		? Number.parseFloat(raw.slice(0, -1)) * 1000
		: Number.parseInt(raw, 10);
	return Number.isFinite(parsed) && parsed !== 0 ? Math.round(parsed) : undefined;
}

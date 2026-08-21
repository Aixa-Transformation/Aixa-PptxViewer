/**
 * EOT (Embedded OpenType) Container Parser
 *
 * Parses EOT containers that may appear in PPTX embedded font parts
 * (.fntdata / .odttf files). EOT is a wrapper format around font data
 * defined by the W3C: https://www.w3.org/Submission/EOT/
 *
 * Some PPTX producers (e.g. Google Slides) embed fonts in EOT format
 * rather than using the simple OOXML XOR obfuscation.
 */

import { decompressEotFont } from 'mtx-decompressor';

/* ------------------------------------------------------------------ */
/*  Constants                                                         */
/* ------------------------------------------------------------------ */

/** Offset of the EOT magic number (0x504C = "LP") in the header. */
const EOT_MAGIC_OFFSET = 34;

/** Expected magic value at the magic offset (little-endian uint16). */
const EOT_MAGIC = 0x504c;

/** Flag: font data is MicroType Express (MTX / BSGP) compressed. */
const TTEMBED_TTCOMPRESSED = 0x0004;

/** Flag: font data is XOR-encrypted using the embedding page URL. */
const TTEMBED_XORENCRYPTDATA = 0x10000000;

/* ------------------------------------------------------------------ */
/*  Binary read helpers                                               */
/* ------------------------------------------------------------------ */

function readUint32LE(data: Uint8Array, offset: number): number {
	return (
		data[offset] |
		(data[offset + 1] << 8) |
		(data[offset + 2] << 16) |
		((data[offset + 3] << 24) >>> 0)
	);
}

function readUint16LE(data: Uint8Array, offset: number): number {
	return data[offset] | (data[offset + 1] << 8);
}

function readUtf16LE(data: Uint8Array, offset: number, byteLength: number): string {
	const chars: string[] = [];
	for (let i = 0; i < byteLength; i += 2) {
		if (offset + i + 1 >= data.length) {
			break;
		}
		const code = data[offset + i] | (data[offset + i + 1] << 8);
		if (code === 0) {
			break;
		}
		chars.push(String.fromCharCode(code));
	}
	return chars.join('');
}

/* ------------------------------------------------------------------ */
/*  EOT header structure                                              */
/* ------------------------------------------------------------------ */

export interface EotHeader {
	/** Total size of the EOT container in bytes. */
	eotSize: number;
	/** Size of the embedded font data in bytes. */
	fontDataSize: number;
	/** EOT format version (typically 0x00020001 or 0x00020002). */
	version: number;
	/** Embedding flags. */
	flags: number;
	/** Whether the font data is MTX/BSGP compressed. */
	isCompressed: boolean;
	/** Whether the font data is XOR-encrypted (URL-based key). */
	isXorEncrypted: boolean;
	/** Font family name from the EOT header. */
	familyName: string;
	/** Font style name from the EOT header. */
	styleName: string;
	/** Font version string from the EOT header. */
	versionName: string;
	/** Full font name from the EOT header. */
	fullName: string;
	/** Byte offset where the font data begins within the container. */
	fontDataOffset: number;
}

export interface EotBuildOptions {
	familyName: string;
	styleName?: string;
	fullName?: string;
	weight?: number;
	italic?: boolean;
}

function readUint16BE(data: Uint8Array, offset: number): number {
	return ((data[offset] ?? 0) << 8) | (data[offset + 1] ?? 0);
}

function readUint32BE(data: Uint8Array, offset: number): number {
	return (
		(((data[offset] ?? 0) << 24) >>> 0) |
		((data[offset + 1] ?? 0) << 16) |
		((data[offset + 2] ?? 0) << 8) |
		(data[offset + 3] ?? 0)
	) >>> 0;
}

function findSfntTable(data: Uint8Array, tag: string): { offset: number; length: number } | null {
	if (data.length < 12) return null;
	const numTables = readUint16BE(data, 4);
	for (let i = 0; i < numTables; i++) {
		const record = 12 + i * 16;
		if (record + 16 > data.length) break;
		const recordTag = String.fromCharCode(
			data[record]!,
			data[record + 1]!,
			data[record + 2]!,
			data[record + 3]!,
		);
		if (recordTag !== tag) continue;
		const offset = readUint32BE(data, record + 8);
		const length = readUint32BE(data, record + 12);
		return offset + length <= data.length ? { offset, length } : null;
	}
	return null;
}

/* ------------------------------------------------------------------ */
/*  Public API                                                        */
/* ------------------------------------------------------------------ */

/**
 * Check whether a binary buffer is an EOT container.
 * Verifies the magic number 0x504C at the documented offset.
 */
export function isEotFormat(data: Uint8Array): boolean {
	if (data.length < 36) {
		return false;
	}
	const magic = readUint16LE(data, EOT_MAGIC_OFFSET);
	return magic === EOT_MAGIC;
}

/**
 * Wrap an individual TrueType/OpenType sfnt in an uncompressed EOT container.
 * PresentationML `.fntdata` parts written by PowerPoint use EOT rather than
 * the GUID-obfuscated WordprocessingML representation.
 */
export function createEotFromSfnt(
	fontData: Uint8Array,
	options: EotBuildOptions,
): Uint8Array {
	const encodeName = (value: string): Uint8Array => {
		const bytes = new Uint8Array((value.length + 1) * 2);
		const view = new DataView(bytes.buffer);
		for (let i = 0; i < value.length; i++) {
			view.setUint16(i * 2, value.charCodeAt(i), true);
		}
		return bytes;
	};
	const names = [
		encodeName(options.familyName),
		encodeName(options.styleName ?? ''),
		encodeName(''),
		encodeName(options.fullName ?? options.familyName),
	];
	// EOTPrefix is 82 bytes. Each name is encoded as uint16 byteLength,
	// UTF-16LE bytes, then a uint16 terminator. A final empty RootString
	// (uint16 zero) precedes the raw SFNT payload for version 0x00020001.
	const headerSize = 82 + names.reduce((sum, name) => sum + 2 + name.length, 0) + 2;
	const result = new Uint8Array(headerSize + fontData.length);
	const view = new DataView(result.buffer);
	const os2 = findSfntTable(fontData, 'OS/2');
	const head = findSfntTable(fontData, 'head');
	view.setUint32(0, result.length, true);
	view.setUint32(4, fontData.length, true);
	view.setUint32(8, 0x00020001, true);
	view.setUint32(12, 0, true); // uncompressed, not XOR-encrypted
	if (os2 && os2.length >= 42) {
		// EOT mirrors the identifying OS/2 fields. Desktop PowerPoint validates
		// these values before accepting an embedded font; zero-filled metadata is
		// tolerated by browsers but causes Office to silently substitute a font.
		result.set(fontData.slice(os2.offset + 32, os2.offset + 42), 16); // PANOSE
	}
	result[26] = 1; // DEFAULT_CHARSET
	result[27] = options.italic ? 1 : 0;
	view.setUint32(
		28,
		options.weight ?? (os2 && os2.length >= 6 ? readUint16BE(fontData, os2.offset + 4) : 400),
		true,
	);
	view.setUint16(
		32,
		os2 && os2.length >= 10 ? readUint16BE(fontData, os2.offset + 8) : 0,
		true,
	);
	view.setUint16(34, EOT_MAGIC, true);
	if (os2 && os2.length >= 58) {
		for (let i = 0; i < 4; i++) {
			view.setUint32(36 + i * 4, readUint32BE(fontData, os2.offset + 42 + i * 4), true);
		}
	}
	if (os2 && os2.length >= 86) {
		view.setUint32(52, readUint32BE(fontData, os2.offset + 78), true);
		view.setUint32(56, readUint32BE(fontData, os2.offset + 82), true);
	}
	if (head && head.length >= 12) {
		view.setUint32(60, readUint32BE(fontData, head.offset + 8), true);
	}

	let offset = 82;
	for (const name of names) {
		view.setUint16(offset, Math.max(0, name.length - 2), true);
		offset += 2;
		result.set(name, offset);
		offset += name.length;
	}
	view.setUint16(offset, 0, true); // empty RootString
	offset += 2;
	result.set(fontData, offset);
	return result;
}

/**
 * Parse the EOT container header and return its metadata.
 * Returns `null` if the data is not a valid EOT container.
 *
 * EOT header layout (W3C Submission):
 * ```
 * Offset  Size   Field
 * 0       4      EOTSize
 * 4       4      FontDataSize
 * 8       4      Version
 * 12      4      Flags
 * 16      10     PANOSE
 * 26      1      Charset
 * 27      1      Italic
 * 28      4      Weight
 * 32      2      fsType
 * 34      2      MagicNumber (0x504C)
 * 36      16     UnicodeRange (4 × uint32)
 * 52      8      CodePageRange (2 × uint32)
 * 60      4      CheckSumAdjustment
 * 64      16     Reserved (4 × uint32)
 * 80      ...    Variable-length name strings + font data
 * ```
 */
export function parseEotHeader(data: Uint8Array): EotHeader | null {
	if (!isEotFormat(data)) {
		return null;
	}
	if (data.length < 82) {
		return null;
	}

	const eotSize = readUint32LE(data, 0);
	const fontDataSize = readUint32LE(data, 4);
	const version = readUint32LE(data, 8);
	const flags = readUint32LE(data, 12);

	// --- Variable-length name strings start at offset 80 ---
	let offset = 80;

	// Helper: read a padded name string (2-byte padding + 2-byte size + data)
	const readNameString = (): string => {
		if (offset + 4 > data.length) {
			return '';
		}
		/* const padding = */ readUint16LE(data, offset);
		offset += 2;
		const size = readUint16LE(data, offset);
		offset += 2;
		if (size === 0 || offset + size > data.length) {
			offset += size;
			return '';
		}
		const str = readUtf16LE(data, offset, size);
		offset += size;
		return str;
	};

	const familyName = readNameString();
	const styleName = readNameString();
	const versionName = readNameString();
	const fullName = readNameString();
	// EOT 0x00020001 also carries RootString immediately after the four
	// identifying names. For unrestricted fonts this is commonly empty, but
	// its two-byte terminator still precedes the SFNT payload.
	if (
		version >= 0x00020001 &&
		offset + 4 <= data.length &&
		readUint16LE(data, offset) === 0 &&
		readUint16LE(data, offset + 2) === 0
	) {
		readNameString();
	}

	// Version 0x00020002+ has additional fields after the four name strings.
	// Per W3C EOT spec: RootString, RootStringChecksum, EUDCCodePage,
	// Padding4+SignatureSize+Signature, EUDCFlags, EUDCFontSize+EUDCFontData.
	if (version >= 0x00020002) {
		// RootStringChecksum (4 bytes) + EUDCCodePage (4 bytes)
		offset += 8;

		// Padding4 (2 bytes) + SignatureSize (2 bytes)
		if (offset + 4 <= data.length) {
			const signatureSize = readUint16LE(data, offset + 2);
			offset += 4 + signatureSize;
		}

		// EUDCFlags (4 bytes) + EUDCFontSize (4 bytes)
		if (offset + 8 <= data.length) {
			const eudcFontSize = readUint32LE(data, offset + 4);
			offset += 8 + eudcFontSize;
		}
	}

	return {
		eotSize,
		fontDataSize,
		version,
		flags,
		isCompressed: (flags & TTEMBED_TTCOMPRESSED) !== 0,
		isXorEncrypted: (flags & TTEMBED_XORENCRYPTDATA) !== 0,
		familyName,
		styleName,
		versionName,
		fullName,
		fontDataOffset: offset,
	};
}

/**
 * Extract the raw font binary (TrueType / OpenType) from an EOT container.
 *
 * - If the embedded font data is **uncompressed**, returns the raw TTF/OTF.
 * - If the font data is **MTX/BSGP compressed**, decompresses it using
 *   the MicroType Express decompressor and returns the reconstructed TTF.
 *
 * @returns An object with the extracted `fontData` and parsed `header`,
 *          or `null` if extraction is not possible.
 */
export function extractFontFromEot(
	data: Uint8Array,
): { fontData: Uint8Array; header: EotHeader } | null {
	const header = parseEotHeader(data);
	if (!header) {
		return null;
	}

	const { fontDataOffset, fontDataSize } = header;

	// Bounds check
	if (fontDataOffset + fontDataSize > data.length) {
		return null;
	}

	const fontData = data.slice(fontDataOffset, fontDataOffset + fontDataSize);

	// Use the EOT header flags to determine if data is compressed
	if (header.isCompressed) {
		try {
			const decompressed = decompressEotFont(
				fontData,
				/* compressed */ true,
				/* encrypted */ header.isXorEncrypted,
			);
			return { fontData: decompressed, header };
		} catch (e) {
			console.warn(`[pptx-viewer] MTX decompression failed for font "${header.familyName}":`, e);
			return null;
		}
	}

	return { fontData, header };
}

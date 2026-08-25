const PERMALINK_KEY = 'code';
const PERMALINK_PREFIX_GZIP = 'gz:';

// Stays well below the ~2000 character URL length that every browser and proxy handles.
const MAX_PERMALINK_URL_LENGTH = 1800;

function bytesToBase64(bytes: Uint8Array): string {
	let binary = '';
	for (const byte of bytes) {
		binary += String.fromCharCode(byte);
	}
	return btoa(binary);
}

function base64ToBytes(encoded: string): Uint8Array {
	const binary = atob(encoded);
	const bytes = new Uint8Array(binary.length);
	for (let i = 0; i < binary.length; i++) {
		bytes[i] = binary.charCodeAt(i);
	}
	return bytes;
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
	return bytes.buffer.slice(
		bytes.byteOffset,
		bytes.byteOffset + bytes.byteLength
	) as ArrayBuffer;
}

async function pipe(bytes: Uint8Array, stream: ReadableWritablePair) {
	const piped = new Blob([toArrayBuffer(bytes)]).stream().pipeThrough(stream);
	return new Uint8Array(await new Response(piped).arrayBuffer());
}

async function encode(source: string): Promise<string> {
	const compressed = await pipe(
		new TextEncoder().encode(source),
		new CompressionStream('gzip')
	);
	return PERMALINK_PREFIX_GZIP + bytesToBase64(compressed);
}

async function decode(encoded: string): Promise<string> {
	if (!encoded.startsWith(PERMALINK_PREFIX_GZIP)) {
		throw new Error('Unsupported permalink format.');
	}
	const uncompressed = await pipe(
		base64ToBytes(encoded.slice(PERMALINK_PREFIX_GZIP.length)),
		new DecompressionStream('gzip')
	);
	return new TextDecoder().decode(uncompressed);
}

/** Reads the source stored in the current URL hash, or null when there is none. */
export async function readPermalink(): Promise<string | null> {
	const encoded = new URLSearchParams(
		globalThis.location.hash.replace(/^#/, '')
	).get(PERMALINK_KEY);

	if (!encoded) return null;

	try {
		return await decode(encoded);
	} catch {
		return null;
	}
}

/** Stores the source in the URL hash and returns the shareable URL. */
export async function writePermalink(source: string): Promise<string> {
	if (!('CompressionStream' in globalThis)) {
		throw new Error('This browser does not support permalinks.');
	}

	const url = new URL(globalThis.location.href);
	url.hash = new URLSearchParams({
		[PERMALINK_KEY]: await encode(source),
	}).toString();

	if (url.href.length > MAX_PERMALINK_URL_LENGTH) {
		throw new Error(
			'Code is too long to share as a link, even compressed.'
		);
	}

	history.replaceState(null, '', url);
	return url.href;
}

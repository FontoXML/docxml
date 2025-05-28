/**
 * Converts the given number (must be an integer) to hexadecimal (4 bytes).
 * @param n The number.
 * @returns The hexadecimal value.
 */
export function toHex(n: number): string {
	return n.toString(16).padStart(8, '0').toUpperCase();
}

/**
 * Parses a hexadecimal number (must be a string & 4 bytes hex) to an integer.
 * @param hex The hexadecimal value.
 * @returns The number.
 */
export function fromHex(hex: string): number {
	return parseInt(hex, 16);
}

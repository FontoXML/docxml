import preact from '@preact/preset-vite';
import { readFileSync } from 'node:fs';
import { defineConfig } from 'vite';

function getVersion(): string {
	try {
		const denoJsonPath = new URL('../deno.json', import.meta.url);
		const denoJson = JSON.parse(readFileSync(denoJsonPath, 'utf8')) as {
			version?: string;
		};
		return denoJson.version || 'dev';
	} catch {
		return 'dev';
	}
}

export default defineConfig({
	base: './',
	plugins: [preact()],
	define: {
		__DOCXML_VERSION__: JSON.stringify(getVersion()),
	},
});

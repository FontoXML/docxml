import { execSync } from 'node:child_process';
import preact from '@preact/preset-vite';
import { defineConfig } from 'vite';

function getVersion(): string {
	try {
		return execSync('git describe --tags --abbrev=0').toString().trim();
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

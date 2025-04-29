/**
 * Running this script rebuilds it as a NodeJS compatible npm package.
 */

import { build, emptyDir } from  'dnt'; 

const VERSION = Deno.args[0];
if (!VERSION) {
	throw new Error('Please specify a version, eg. "deno task dnt 1.0.0".');
}

const msg = ` Creating npm package for version ${VERSION} `;
console.log('-'.repeat(msg.length) + '\n' + msg + '\n' + '-'.repeat(msg.length));

await emptyDir('./npm');

await build({
	entryPoints: ['./mod.ts'],
	outDir: './npm',
	// typeCheck: true,
	test: false,
	// declaration: true,
	skipSourceOutput: true,
	shims: {
		deno: true,
		custom: [
			{
				package: { name: 'crypto' },
				globalNames: [{ name: 'crypto', exportName: 'default' }],
			},
		],
	},
	package: {
		name: 'docxml',
		version: VERSION,
		description: 'TypeScript (component) library for building and parsing a DOCX file',
		author: {
			name: 'Fonto BV',
			email: 'gabe.webb@rws.com',
			url: 'https://github.com/fontoxml',
		},
		contributors: [],
		homepage: 'https://github.com/fontoxml/docxml',
		repository: {
			type: 'git',
			url: 'git+https://github.com/fontoxml/docxml.git',
		},
		bugs: {
			url: 'https://github.com/wvbe/docxml/issues',
		},
		license: 'none',
		keywords: ['ooxml', 'docx', 'components', 'deno', 'node', 'jsx'],
		type: 'module',
		main: 'script/mod.js',
		module: 'esm/mod.js',
		typings: 'types/mod.d.ts',
	},
	mappings: {
		'https://esm.sh/fontoxpath@3.28.2?pin=v121': {
			name: 'fontoxpath',
			version: '3.28.2',
		},
		'https://esm.sh/slimdom@4.3.5?pin=v121': {
			name: 'slimdom',
			version: '4.3.5',
		},
	},
});

await Deno.copyFile('README.md', 'npm/README.md');

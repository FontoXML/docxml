/// <reference lib="deno.ns" />
import ts from 'typescript';

const outDir = new URL('../public/docxml/', import.meta.url);
const outDtsPath = new URL('docxml.d.ts', outDir);
const outVersionPath = new URL('version.json', outDir);
const entryFilePath = new URL('../../mod.ts', import.meta.url).pathname;

function emitSingleDts(): string {
	const outputPath = '/out/docxml.d.ts';

	const options: ts.CompilerOptions = {
		declaration: true,
		emitDeclarationOnly: true,
		module: ts.ModuleKind.System,
		outFile: outputPath,
		target: ts.ScriptTarget.ES2020,
		moduleResolution: ts.ModuleResolutionKind.NodeNext,
		allowImportingTsExtensions: true,
		skipLibCheck: true,
		noLib: true,
	};

	const host = ts.createCompilerHost(options);

	let outDts = '';
	host.writeFile = (fileName, text) => {
		if (fileName === outputPath) {
			outDts = text;
		}
	};

	const program = ts.createProgram([entryFilePath], options, host);

	program.emit();

	if (!outDts) {
		throw new Error('No declaration output was generated.');
	}

	return [
		'// Generated file. Do not edit manually.',
		outDts,
		"declare module 'docxml' {",
		"  export * from 'mod';",
		"  export { default } from 'mod';",
		'}',
		'',
	].join('\n');
}

async function main() {
	await Deno.mkdir(outDir, { recursive: true });

	console.log('Generating single docxml.d.ts from local sources...');
	const dts = emitSingleDts();

	await Deno.writeTextFile(outDtsPath, dts);
	await Deno.writeTextFile(outVersionPath, 'local');

	console.log(`Generated ${outDtsPath.pathname} (${dts.length} chars)`);
}

await main();

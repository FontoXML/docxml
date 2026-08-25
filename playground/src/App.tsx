import 'monaco-editor/esm/vs/basic-languages/typescript/typescript.contribution';
import 'monaco-editor/esm/vs/editor/editor.all';
import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';
import 'monaco-editor/esm/vs/language/typescript/monaco.contribution';
import { useEffect, useRef, useState } from 'preact/hooks';
import ts from 'typescript';
import { examples } from './examples.ts';
import { readPermalink, writePermalink } from './permalink.ts';

// Pinned so the executed code matches the types shipped with this build.
const DOCXML_RUNTIME_URL = `https://esm.sh/jsr/@fontoxml/docxml@${__DOCXML_VERSION__}?bundle`;
const DOCXML_TYPES_URL = new URL('../docxml/docxml.d.ts', import.meta.url);

// Matches the module specifier of `from 'docxml'`, `import('docxml')` and friends.
const DOCXML_SPECIFIER = /(\bfrom\s*|\bimport\s*\(\s*)(['"])docxml\2/g;

const INITIAL_SOURCE = examples[0]?.source ?? ''; // The "Hello world" example.
const MODEL_URI = monaco.Uri.parse('file:///playground/main.ts');

// Workers for Monaco
(self as unknown as { MonacoEnvironment: unknown }).MonacoEnvironment = {
	getWorker(_workerId: string, label: string) {
		if (label === 'typescript' || label === 'javascript') {
			return new Worker(
				new URL('./workers/ts.worker.ts', import.meta.url),
				{ type: 'module' }
			);
		}
		return new Worker(
			new URL('./workers/editor.worker.ts', import.meta.url),
			{ type: 'module' }
		);
	},
};

// docxml reads Deno.cwd() even in the browser.
if (!(globalThis as unknown as { Deno?: unknown }).Deno) {
	(globalThis as unknown as { Deno: unknown }).Deno = { cwd: () => '/' };
}

async function fetchText(url: URL | string): Promise<string> {
	const response = await fetch(url);
	if (!response.ok) {
		throw new Error(`${response.status} ${response.statusText}`);
	}
	return response.text();
}

/** Feeds the declaration bundle generated at build time to Monaco. */
async function loadDocxmlTypes() {
	const bundledDts = await fetchText(DOCXML_TYPES_URL);
	const tsDefaults = monaco.languages.typescript.typescriptDefaults;

	tsDefaults.setCompilerOptions({
		target: monaco.languages.typescript.ScriptTarget.ES2020,
		module: monaco.languages.typescript.ModuleKind.ESNext,
		moduleResolution:
			monaco.languages.typescript.ModuleResolutionKind.NodeJs,
		allowSyntheticDefaultImports: true,
		esModuleInterop: true,
		allowImportingTsExtensions: true,
		strict: true,
		baseUrl: 'file:///',
	});
	tsDefaults.setEagerModelSync(true);
	tsDefaults.addExtraLib(bundledDts, 'file:///docxml/docxml-bundle.d.ts');
}

/** Returns the messages Monaco reports for the editor contents. */
async function getDiagnostics(): Promise<string[]> {
	const worker = await (
		await monaco.languages.typescript.getTypeScriptWorker()
	)(MODEL_URI);
	const uri = MODEL_URI.toString();

	const diagnostics = [
		...(await worker.getSyntacticDiagnostics(uri)),
		...(await worker.getSemanticDiagnostics(uri)),
	];

	return diagnostics.map((diagnostic) =>
		typeof diagnostic.messageText === 'string'
			? diagnostic.messageText
			: diagnostic.messageText.messageText
	);
}

async function normalizeResult(result: unknown): Promise<Uint8Array> {
	if (result instanceof Uint8Array) return result;
	if (result instanceof ArrayBuffer) return new Uint8Array(result);
	if (result instanceof Blob) {
		return new Uint8Array(await result.arrayBuffer());
	}

	const obj = result as {
		asUint8Array?: () => Uint8Array;
		toArchive?: () => Promise<{ asUint8Array: () => Uint8Array }>;
	};
	if (typeof obj?.asUint8Array === 'function') return obj.asUint8Array();
	if (typeof obj?.toArchive === 'function') {
		return (await obj.toArchive()).asUint8Array();
	}

	throw new Error(
		'Result must be Docx, Archive, Uint8Array, Blob, or ArrayBuffer.'
	);
}

function download(blob: Blob, fileName: string) {
	const url = URL.createObjectURL(blob);
	const anchor = document.createElement('a');
	anchor.href = url;
	anchor.download = fileName;
	document.body.append(anchor);
	anchor.click();
	anchor.remove();
	// Safari aborts the download when the URL is revoked synchronously.
	setTimeout(() => URL.revokeObjectURL(url));
}

function messageOf(err: unknown): string {
	return err instanceof Error ? err.message : String(err);
}

export function App() {
	const editorRef = useRef<HTMLDivElement | null>(null);
	const instanceRef = useRef<monaco.editor.IStandaloneCodeEditor | null>(
		null
	);
	const [status, setStatus] = useState('Loading types...');
	const [hasError, setHasError] = useState(false);
	const [isBusy, setIsBusy] = useState(false);

	function report(message: string, isFailure = false) {
		setStatus(message);
		setHasError(isFailure);
	}

	useEffect(() => {
		if (!editorRef.current) return;

		const model = monaco.editor.createModel(
			INITIAL_SOURCE,
			'typescript',
			MODEL_URI
		);
		const editor = monaco.editor.create(editorRef.current, {
			model,
			theme: 'vs',
			automaticLayout: true,
			minimap: { enabled: false },
			fontSize: 13,
		});
		instanceRef.current = editor;

		let disposed = false;
		(async () => {
			try {
				await loadDocxmlTypes();
				const shared = await readPermalink();
				if (disposed) return;
				if (shared) model.setValue(shared);
				report('Ready.');
			} catch (err) {
				if (!disposed) report(`Types failed: ${messageOf(err)}`, true);
			}
		})();

		return () => {
			disposed = true;
			model.dispose();
			editor.dispose();
			instanceRef.current = null;
		};
	}, []);

	async function generateDocx() {
		const source = instanceRef.current?.getValue();
		if (!source) return;

		setIsBusy(true);
		report('Compiling...');

		try {
			const [firstError] = await getDiagnostics();
			if (firstError) throw new Error(firstError);

			const jsSource = ts
				.transpileModule(source, {
					compilerOptions: {
						target: ts.ScriptTarget.ES2020,
						module: ts.ModuleKind.ES2022,
					},
				})
				.outputText.replace(
					DOCXML_SPECIFIER,
					`$1$2${DOCXML_RUNTIME_URL}$2`
				);

			const url = URL.createObjectURL(
				new Blob([jsSource], { type: 'text/javascript' })
			);

			try {
				const mod = await import(/* @vite-ignore */ url);
				if (typeof mod.default !== 'function') {
					throw new Error('Module must export a default function.');
				}
				report('Generating DOCX...');
				const data = await normalizeResult(await mod.default());
				download(
					new Blob([Uint8Array.from(data)], {
						type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
					}),
					'playground-output.docx'
				);
				report(`Done (${data.byteLength} bytes).`);
			} finally {
				URL.revokeObjectURL(url);
			}
		} catch (err) {
			report(messageOf(err), true);
		} finally {
			setIsBusy(false);
		}
	}

	async function copyPermalink() {
		const source = instanceRef.current?.getValue();
		if (!source) return;

		try {
			await navigator.clipboard.writeText(await writePermalink(source));
			report('Permalink copied.');
		} catch (err) {
			report(messageOf(err), true);
		}
	}

	function loadExample(id: string) {
		const example = examples.find((entry) => entry.id === id);
		if (example) instanceRef.current?.setValue(example.source);
	}

	return (
		<div class='layout'>
			<header class='header'>
				<h1 class='headerTitle'>Docxml Playground</h1>
				<select
					class='exampleSelect'
					onChange={(e) => loadExample(e.currentTarget.value)}
				>
					<option value=''>Load example…</option>
					{examples.map((example) => (
						<option key={example.id} value={example.id}>
							{example.label}
						</option>
					))}
				</select>
			</header>

			<div ref={editorRef} class='editor' />

			<div class='statusRow'>
				<p class={`status${hasError ? ' error' : ''}`}>{status}</p>
				<button
					type='button'
					class='secondary'
					onClick={() => void copyPermalink()}
				>
					Copy permalink
				</button>
				<button
					type='button'
					class='secondary'
					onClick={() => {
						const source = instanceRef.current?.getValue();
						if (source) {
							download(
								new Blob([source], { type: 'text/plain' }),
								'playground.ts'
							);
						}
					}}
				>
					Download .ts
				</button>
				<button
					type='button'
					disabled={isBusy}
					onClick={() => void generateDocx()}
				>
					{isBusy ? 'Generating...' : 'Generate DOCX'}
				</button>
			</div>

			<footer class='footer'>
				<span class='footerVersion'>{__DOCXML_VERSION__}</span>
				<span>·</span>
				<a
					href='https://github.com/fontoxml/docxml'
					target='_blank'
					rel='noopener'
				>
					GitHub
				</a>
				<span>·</span>
				<a href='https://fontoxml.com' target='_blank' rel='noopener'>
					FontoXML
				</a>
			</footer>
		</div>
	);
}

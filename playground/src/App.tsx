import 'monaco-editor/esm/vs/basic-languages/typescript/typescript.contribution';
import 'monaco-editor/esm/vs/editor/editor.all';
import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';
import 'monaco-editor/esm/vs/language/typescript/monaco.contribution';
import { useEffect, useRef, useState } from 'preact/hooks';
import ts from 'typescript';
import { examples } from './examples.ts';

const DOCXML_RUNTIME_URL = 'https://esm.sh/jsr/@fontoxml/docxml?bundle';
const DOCXML_TYPES_URL = new URL('../docxml/docxml.d.ts', import.meta.url);

const INITIAL_SOURCE = examples[0]?.source ?? ''; // The "Hello world" example.

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

// Polyfill Deno.cwd for docxml browser compat
if (!(globalThis as unknown as { Deno?: unknown }).Deno) {
	(globalThis as unknown as { Deno: unknown }).Deno = { cwd: () => '/' };
}

/**
 * Loads the precompiled docxml declaration bundle generated at build time
 * and feeds it to Monaco's TypeScript language service.
 */
async function loadDocxmlTypes() {
	const bundledDts = await fetch(DOCXML_TYPES_URL).then((r) => r.text());

	// Configure TS compiler options in Monaco
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

	tsDefaults.setDiagnosticsOptions({
		noSemanticValidation: false,
		noSyntaxValidation: false,
	});

	tsDefaults.setEagerModelSync(true);

	tsDefaults.addExtraLib(bundledDts, 'file:///docxml/docxml-bundle.d.ts');
}

async function normalizeResult(result: unknown): Promise<Uint8Array> {
	if (result instanceof Uint8Array) return result;
	if (result instanceof ArrayBuffer) return new Uint8Array(result);
	if (result instanceof Blob)
		return new Uint8Array(await result.arrayBuffer());

	const obj = result as Record<string, unknown>;
	if (typeof obj?.asUint8Array === 'function') {
		return (obj as { asUint8Array: () => Uint8Array }).asUint8Array();
	}
	if (typeof obj?.toArchive === 'function') {
		const archive = await (
			obj as {
				toArchive: () => Promise<{ asUint8Array: () => Uint8Array }>;
			}
		).toArchive();
		return archive.asUint8Array();
	}

	throw new Error(
		'Result must be Docx, Archive, Uint8Array, Blob, or ArrayBuffer.'
	);
}

function downloadBlob(blob: Blob, fileName: string) {
	const url = URL.createObjectURL(blob);
	const a = document.createElement('a');
	a.href = url;
	a.download = fileName;
	a.click();
	URL.revokeObjectURL(url);
}

function downloadDocx(data: Uint8Array, fileName: string) {
	downloadBlob(
		new Blob([Uint8Array.from(data)], {
			type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		}),
		fileName
	);
}

export function App() {
	const editorRef = useRef<HTMLDivElement | null>(null);
	const instanceRef = useRef<monaco.editor.IStandaloneCodeEditor | null>(
		null
	);
	const [status, setStatus] = useState('Loading types...');
	const [hasError, setHasError] = useState(false);
	const [isGenerating, setIsGenerating] = useState(false);

	useEffect(() => {
		if (!editorRef.current) return;

		const model = monaco.editor.createModel(
			INITIAL_SOURCE,
			'typescript',
			monaco.Uri.parse('file:///playground/main.ts')
		);

		instanceRef.current = monaco.editor.create(editorRef.current, {
			model,
			theme: 'vs',
			automaticLayout: true,
			minimap: { enabled: false },
			fontSize: 13,
		});

		loadDocxmlTypes()
			.then(() => {
				setStatus('Ready.');
				setHasError(false);
			})
			.catch((err) => {
				setStatus(
					`Types failed: ${err instanceof Error ? err.message : err}`
				);
				setHasError(true);
			});

		return () => {
			model.dispose();
			instanceRef.current?.dispose();
		};
	}, []);

	async function generateDocx() {
		const editor = instanceRef.current;
		if (!editor) return;

		setIsGenerating(true);
		setHasError(false);
		setStatus('Compiling...');

		try {
			const source = editor.getValue();
			const jsSource = ts
				.transpileModule(source, {
					compilerOptions: {
						target: ts.ScriptTarget.ES2020,
						module: ts.ModuleKind.ES2022,
						strict: true,
					},
				})
				.outputText.replaceAll(
					`from 'docxml'`,
					`from '${DOCXML_RUNTIME_URL}'`
				)
				.replaceAll(`from "docxml"`, `from "${DOCXML_RUNTIME_URL}"`);

			const blob = new Blob([jsSource], { type: 'text/javascript' });
			const url = URL.createObjectURL(blob);

			try {
				const mod = await import(/* @vite-ignore */ url);
				if (typeof mod.default !== 'function') {
					throw new Error('Module must export a default function.');
				}
				setStatus('Generating DOCX...');
				const result = await mod.default();
				const data = await normalizeResult(result);
				downloadDocx(data, 'playground-output.docx');
				setStatus(`Done (${data.byteLength} bytes).`);
				setHasError(false);
			} finally {
				URL.revokeObjectURL(url);
			}
		} catch (err) {
			setStatus(err instanceof Error ? err.message : String(err));
			setHasError(true);
		} finally {
			setIsGenerating(false);
		}
	}

	function loadExample(id: string) {
		const example = examples.find((e) => e.id === id);
		if (example && instanceRef.current) {
			instanceRef.current.getModel()?.setValue(example.source);
		}
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
					{examples.map((ex) => (
						<option key={ex.id} value={ex.id}>
							{ex.label}
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
					onClick={() => {
						const source = instanceRef.current?.getValue();
						if (!source) return;
						downloadBlob(
							new Blob([source], { type: 'text/plain' }),
							'playground.ts'
						);
					}}
				>
					Download .ts
				</button>
				<button
					type='button'
					disabled={isGenerating}
					onClick={() => void generateDocx()}
				>
					{isGenerating ? 'Generating...' : 'Generate DOCX'}
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

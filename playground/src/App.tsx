import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';
import 'monaco-editor/esm/vs/language/typescript/monaco.contribution';
import { useEffect, useRef, useState } from 'preact/hooks';
import ts from 'typescript';
// @ts-ignore Vite resolves worker query imports during bundling.
import EditorWorker from 'monaco-editor/esm/vs/editor/editor.worker?worker';
// @ts-ignore Vite resolves worker query imports during bundling.
import TypeScriptWorker from 'monaco-editor/esm/vs/language/typescript/ts.worker?worker';
import { playgroundExamples } from './examples.ts';

const DOCXML_PACKAGE_NAME = '@fontoxml/docxml';
const DOCXML_META_URL = 'https://jsr.io/@fontoxml/docxml/meta.json';
const DOCXML_RUNTIME_URL = 'https://esm.sh/jsr/@fontoxml/docxml?bundle';
const docxmlSourceCache = new Map<string, string>();

const monacoEnvironmentTarget = globalThis as typeof globalThis & {
	MonacoEnvironment?: {
		getWorker: (_workerId: string, label: string) => Worker;
	};
};

monacoEnvironmentTarget.MonacoEnvironment = {
	getWorker(_workerId: string, label: string) {
		if (label === 'typescript' || label === 'javascript') {
			return new TypeScriptWorker();
		}

		return new EditorWorker();
	},
};

function ensureBrowserDenoCompat() {
	const denoTarget = globalThis as typeof globalThis & {
		Deno?: {
			cwd?: () => string;
		};
	};

	if (!denoTarget.Deno) {
		denoTarget.Deno = {
			cwd: () => '/',
		};
		return;
	}

	if (typeof denoTarget.Deno.cwd !== 'function') {
		denoTarget.Deno.cwd = () => '/';
	}
}

async function fetchJsrSourceText(version: string, filePath: string) {
	const cacheKey = `${version}:${filePath}`;
	if (docxmlSourceCache.has(cacheKey)) {
		return docxmlSourceCache.get(cacheKey) ?? '';
	}

	const response = await fetch(
		`https://jsr.io/${DOCXML_PACKAGE_NAME}/${version}${filePath}?raw=1`
	);

	if (!response.ok) {
		throw new Error(`Failed to load JSR source file: ${filePath}`);
	}

	const text = await response.text();
	const contentType = response.headers.get('content-type') ?? '';

	if (contentType.includes('text/html')) {
		const document = new DOMParser().parseFromString(text, 'text/html');
		const codeElement = document.querySelector('pre');
		if (codeElement?.textContent) {
			const extracted = codeElement.textContent;
			docxmlSourceCache.set(cacheKey, extracted);
			return extracted;
		}
	}

	if (!text.trim()) {
		throw new Error(
			`Could not extract source from JSR file page: ${filePath}`
		);
	}

	docxmlSourceCache.set(cacheKey, text);
	return text;
}

function toMonacoFilePath(version: string, filePath: string) {
	return `file:///jsr/${DOCXML_PACKAGE_NAME}/${version}${filePath}`;
}

async function installMonacoTypes() {
	const metaResponse = await fetch(DOCXML_META_URL);
	if (!metaResponse.ok) {
		throw new Error('Could not load JSR metadata for docxml.');
	}

	const meta = (await metaResponse.json()) as {
		latest?: string;
		versions?: Record<string, { manifest?: Record<string, unknown> }>;
	};
	const version = meta.latest;

	if (!version) {
		throw new Error('Could not resolve the latest JSR package version.');
	}

	const versionMetaResponse = await fetch(
		`https://jsr.io/${DOCXML_PACKAGE_NAME}/${version}_meta.json`
	);

	if (!versionMetaResponse.ok) {
		throw new Error('Could not load the latest JSR package manifest.');
	}

	const versionMeta = (await versionMetaResponse.json()) as {
		manifest?: Record<string, unknown>;
	};

	const manifest = versionMeta.manifest ?? meta.versions?.[version]?.manifest;

	if (!manifest) {
		throw new Error('Could not resolve the latest JSR package manifest.');
	}

	monaco.languages.typescript.typescriptDefaults.setCompilerOptions({
		allowSyntheticDefaultImports: true,
		esModuleInterop: true,
		allowImportingTsExtensions: true,
		strict: true,
	});
	monaco.languages.typescript.typescriptDefaults.setDiagnosticsOptions({
		noSemanticValidation: false,
		noSyntaxValidation: false,
	});
	monaco.languages.typescript.typescriptDefaults.setEagerModelSync(true);

	const sourceFiles = Object.keys(manifest).filter(
		(filePath) => filePath.endsWith('.ts') || filePath.endsWith('.tsx')
	);

	const sourceEntries = await Promise.all(
		sourceFiles.map(async (filePath) => ({
			filePath,
			text: await fetchJsrSourceText(version, filePath),
		}))
	);

	for (const { filePath, text } of sourceEntries) {
		monaco.languages.typescript.typescriptDefaults.addExtraLib(
			text,
			toMonacoFilePath(version, filePath)
		);
	}

	monaco.languages.typescript.typescriptDefaults.addExtraLib(
		[
			`declare module 'docxml' {`,
			`\texport * from '${toMonacoFilePath(version, '/mod.ts')}';`,
			`\texport { default } from '${toMonacoFilePath(version, '/mod.ts')}';`,
			'}',
			'',
		].join('\n'),
		'file:///docxml-playground/docxml-module.d.ts'
	);
}

function fallbackUuid() {
	const bytes = new Uint8Array(16);

	if (globalThis.crypto?.getRandomValues) {
		globalThis.crypto.getRandomValues(bytes);
	} else {
		for (let index = 0; index < bytes.length; index++) {
			bytes[index] = Math.floor(Math.random() * 256);
		}
	}

	bytes[6] = (bytes[6] & 0x0f) | 0x40;
	bytes[8] = (bytes[8] & 0x3f) | 0x80;

	const hex = Array.from(bytes, (value) =>
		value.toString(16).padStart(2, '0')
	);
	return `${hex.slice(0, 4).join('')}-${hex.slice(4, 6).join('')}-${hex
		.slice(6, 8)
		.join('')}-${hex.slice(8, 10).join('')}-${hex.slice(10, 16).join('')}`;
}

function ensureRandomUuid() {
	const cryptoTarget = globalThis.crypto as Crypto & {
		randomUUID?: () => string;
	};

	if (typeof cryptoTarget?.randomUUID === 'function') {
		return;
	}

	if (!globalThis.crypto) {
		(globalThis as typeof globalThis & { crypto: Crypto }).crypto = {
			randomUUID: fallbackUuid,
		} as unknown as Crypto;
		return;
	}

	Object.defineProperty(globalThis.crypto, 'randomUUID', {
		value: fallbackUuid,
		configurable: true,
		writable: true,
	});
}

async function normalizeResult(result: unknown) {
	if (result instanceof Uint8Array) {
		return result;
	}

	if (result instanceof ArrayBuffer) {
		return new Uint8Array(result);
	}

	if (result instanceof Blob) {
		return new Uint8Array(await result.arrayBuffer());
	}

	if (
		result &&
		typeof result === 'object' &&
		'asUint8Array' in result &&
		typeof (result as { asUint8Array: () => Uint8Array }).asUint8Array ===
			'function'
	) {
		return (result as { asUint8Array: () => Uint8Array }).asUint8Array();
	}

	if (
		result &&
		typeof result === 'object' &&
		'toArchive' in result &&
		typeof (
			result as {
				toArchive: () => Promise<{ asUint8Array: () => Uint8Array }>;
			}
		).toArchive === 'function'
	) {
		const archive = await (
			result as {
				toArchive: () => Promise<{ asUint8Array: () => Uint8Array }>;
			}
		).toArchive();
		return archive.asUint8Array();
	}

	throw new Error('Result type is not supported for DOCX download');
}

function downloadFile(data: Uint8Array, fileName: string) {
	const view = new Uint8Array(data);
	const stableBuffer = view.buffer.slice(
		view.byteOffset,
		view.byteOffset + view.byteLength
	);

	const blob = new Blob([stableBuffer], {
		type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
	});

	const url = URL.createObjectURL(blob);
	const anchor = document.createElement('a');
	anchor.href = url;
	anchor.download = fileName;
	anchor.click();
	URL.revokeObjectURL(url);
}

function transpileTypeScript(source: string) {
	return ts.transpileModule(source, {
		compilerOptions: {
			target: ts.ScriptTarget.ES2020,
			module: ts.ModuleKind.ES2022,
			strict: true,
		},
		fileName: 'playground.ts',
		reportDiagnostics: false,
	}).outputText;
}

function buildRunnableModuleSource(source: string) {
	const hasDefaultExport = /\bexport\s+default\b/m.test(source);
	const hasTopLevelModuleSyntax =
		/^\s*import\s/m.test(source) || /^\s*export\s/m.test(source);

	if (hasDefaultExport) {
		return source;
	}

	if (hasTopLevelModuleSyntax) {
		throw new Error(
			'This code has module syntax. Export a default function to run it.'
		);
	}

	const symbolNames = [
		'Docx',
		'Paragraph',
		'Text',
		'Section',
		'Table',
		'Row',
		'Cell',
		'Comment',
		'CommentRangeStart',
		'CommentRangeEnd',
		'cm',
		'pt',
		'inch',
	];

	return `
import DocxDefault, * as __docxml from '${DOCXML_RUNTIME_URL}';
const Docx = __docxml.Docx ?? DocxDefault;
const { ${symbolNames.filter((name) => name !== 'Docx').join(', ')} } = __docxml;

export default async function __playgroundRun() {
${source}
}

function registerFallbackCompletions() {
	const sharedSymbols = [
		'Docx',
		'Paragraph',
		'Text',
		'Section',
		'Table',
		'Row',
		'Cell',
		'Comment',
		'CommentRangeStart',
		'CommentRangeEnd',
		'cm',
		'pt',
		'inch',
	];

	return monaco.languages.registerCompletionItemProvider('typescript', {
		triggerCharacters: ['.', "'", '"'],
		provideCompletionItems(model, position) {
			const currentLine = model.getLineContent(position.lineNumber);
			const range = new monaco.Range(
				position.lineNumber,
				position.column,
				position.lineNumber,
				position.column
			);

			const suggestions: monaco.languages.CompletionItem[] = [
				{
					label: "import docxml",
					kind: monaco.languages.CompletionItemKind.Snippet,
					insertText:
						"import Docx, { Paragraph, Text } from 'docxml';\n\nexport default function buildDocx() {\n\tconst docx = Docx.fromNothing();\n\n\t$0\n\n\treturn docx;\n}\n",
					insertTextRules:
						monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
					range,
				},
			];

			for (const symbol of sharedSymbols) {
				suggestions.push({
					label: symbol,
					kind: monaco.languages.CompletionItemKind.Class,
					insertText: symbol,
					range,
				});
			}

			if (/from\s+['\"]?$/.test(currentLine)) {
				suggestions.push({
					label: 'docxml',
					kind: monaco.languages.CompletionItemKind.Module,
					insertText: 'docxml',
					range,
				});
			}

			return { suggestions };
		},
	});
}
`;
}

export function App() {
	const editorElementRef = useRef<HTMLDivElement | null>(null);
	const editorRef = useRef<monaco.editor.IStandaloneCodeEditor | null>(null);
	const modelRef = useRef<monaco.editor.ITextModel | null>(null);
	const completionProviderRef = useRef<monaco.IDisposable | null>(null);
	const [status, setStatus] = useState('Ready.');
	const [hasError, setHasError] = useState(false);
	const [selectedExampleId, setSelectedExampleId] = useState('');
	const [isGenerating, setIsGenerating] = useState(false);

	useEffect(() => {
		ensureBrowserDenoCompat();

		if (!editorElementRef.current) {
			setStatus('Editor container not found.');
			setHasError(true);
			return;
		}

		modelRef.current = monaco.editor.createModel(
			playgroundExamples[0]?.source ?? '',
			'typescript',
			monaco.Uri.parse('file:///playground/main.ts')
		);

		editorRef.current = monaco.editor.create(editorElementRef.current, {
			model: modelRef.current,
			theme: 'vs',
			automaticLayout: true,
			minimap: { enabled: false },
			fontSize: 13,
			suggestOnTriggerCharacters: true,
			quickSuggestions: {
				other: true,
				comments: false,
				strings: true,
			},
		});

		completionProviderRef.current = registerFallbackCompletions();

		setStatus('Loading docxml types...');
		setHasError(false);

		void installMonacoTypes()
			.then(() => {
				setStatus('Ready.');
				setHasError(false);
			})
			.catch((error) => {
				const message =
					error instanceof Error ? error.message : String(error);
				setStatus(`Types unavailable: ${message}`);
				setHasError(true);
			});

		return () => {
			completionProviderRef.current?.dispose();
			editorRef.current?.dispose();
			modelRef.current?.dispose();
			completionProviderRef.current = null;
			editorRef.current = null;
			modelRef.current = null;
		};
	}, []);

	async function runCurrentEditor() {
		const editor = editorRef.current;
		if (!editor) {
			setStatus('Editor is not ready yet.');
			setHasError(true);
			return;
		}

		const source = editor.getValue();

		setStatus('Compiling...');
		setHasError(false);
		setIsGenerating(true);
		ensureRandomUuid();

		try {
			const runnableSource = buildRunnableModuleSource(source);
			const jsSource = transpileTypeScript(runnableSource);
			const runtimeSource = jsSource
				.replaceAll(`from 'docxml'`, `from '${DOCXML_RUNTIME_URL}'`)
				.replaceAll(`from "docxml"`, `from "${DOCXML_RUNTIME_URL}"`);

			const moduleBlob = new Blob([runtimeSource], {
				type: 'text/javascript',
			});
			const moduleUrl = URL.createObjectURL(moduleBlob);

			try {
				const run = (await import(moduleUrl)) as {
					default?: () => unknown;
				};
				setStatus('Generating DOCX...');

				if (typeof run.default !== 'function') {
					throw new Error(
						'The module must export a default function.'
					);
				}

				const result = await run.default();
				const data = await normalizeResult(result);

				downloadFile(data, 'playground-output.docx');
				setStatus(`Document generated (${data.byteLength} bytes).`);
				setHasError(false);
			} finally {
				URL.revokeObjectURL(moduleUrl);
			}
		} catch (error) {
			const message =
				error instanceof Error ? error.message : String(error);
			setStatus(message);
			setHasError(true);
		} finally {
			setIsGenerating(false);
		}
	}

	function loadSelectedExample() {
		if (!selectedExampleId) {
			setStatus('Select an example to load.');
			setHasError(false);
			return;
		}

		const editor = editorRef.current;
		if (!editor) {
			setStatus('Editor is not ready yet.');
			setHasError(true);
			return;
		}

		const example = playgroundExamples.find(
			(entry) => entry.id === selectedExampleId
		);

		if (!example) {
			setStatus('Selected example was not found.');
			setHasError(true);
			return;
		}

		editor.getModel()?.setValue(example.source);
		setStatus(`Loaded example: ${example.label}`);
		setHasError(false);
	}

	return (
		<div class='layout'>
			<header class='header'>
				<div>
					<h1 class='title'>Docxml Playground</h1>
					<p class='subtitle'>
						Write TypeScript and generate your DOCX in the browser.
					</p>
				</div>
				<div class='controls'>
					<label class='control' for='example-select'>
						Example
						<select
							id='example-select'
							value={selectedExampleId}
							onChange={(event) =>
								setSelectedExampleId(event.currentTarget.value)
							}
						>
							<option value=''>Select an example…</option>
							{playgroundExamples.map((example) => (
								<option key={example.id} value={example.id}>
									{example.label}
									{example.sourcePath
										? ` (${example.sourcePath})`
										: ''}
								</option>
							))}
						</select>
					</label>
					<button type='button' onClick={loadSelectedExample}>
						Load
					</button>
					<button
						type='button'
						class='primary'
						disabled={isGenerating}
						onClick={() => void runCurrentEditor()}
					>
						{isGenerating ? 'Generating...' : 'Generate DOCX'}
					</button>
				</div>
			</header>

			<section class='editor-shell'>
				<div ref={editorElementRef} class='editor' />
				<div class={`status${hasError ? ' error' : ''}`}>{status}</div>
			</section>

			<p class='hint'>
				Use plain TypeScript. Available symbols include Docx, Paragraph,
				Text, Section, Table, Row, Cell, Comment, CommentRangeStart,
				CommentRangeEnd, cm, pt, and inch. Return Docx, Archive,
				Uint8Array, ArrayBuffer, or Blob.
			</p>
		</div>
	);
}

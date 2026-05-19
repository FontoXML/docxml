import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';
import 'monaco-editor/esm/vs/language/typescript/monaco.contribution';
import { useEffect, useRef, useState } from 'preact/hooks';
import ts from 'typescript';

const DOCXML_PACKAGE_NAME = '@fontoxml/docxml';
const DOCXML_META_URL = 'https://jsr.io/@fontoxml/docxml/meta.json';
const DOCXML_RUNTIME_URL = 'https://esm.sh/jsr/@fontoxml/docxml?bundle';
const PLAYGROUND_FILE = 'file:///playground/main.ts';

const INITIAL_SOURCE = `import Docx, { Paragraph, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.set(
		new Paragraph({}, new Text({}, 'Hello from the docxml playground.'))
	);

	return docx;
}
`;

const docxmlSourceCache = new Map<string, string>();

const monacoEnvironmentTarget = globalThis as typeof globalThis & {
	MonacoEnvironment?: {
		getWorker: (_workerId: string, label: string) => Worker;
	};
};

monacoEnvironmentTarget.MonacoEnvironment = {
	getWorker(_workerId: string, label: string) {
		if (label === 'typescript' || label === 'javascript') {
			return new Worker(
				new URL('./workers/ts.worker.ts', import.meta.url),
				{
					type: 'module',
				}
			);
		}

		return new Worker(
			new URL('./workers/editor.worker.ts', import.meta.url),
			{
				type: 'module',
			}
		);
	},
};

function ensureBrowserPolyfills() {
	const denoTarget = globalThis as typeof globalThis & {
		Deno?: {
			cwd?: () => string;
		};
	};

	if (!denoTarget.Deno) {
		denoTarget.Deno = {
			cwd: () => '/',
		};
	} else if (typeof denoTarget.Deno.cwd !== 'function') {
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
		throw new Error(`No se pudo cargar el archivo JSR: ${filePath}`);
	}

	const text = await response.text();
	if (!text.trim()) {
		throw new Error(`Archivo vacio en JSR: ${filePath}`);
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
		throw new Error('No se pudo cargar metadata de docxml en JSR.');
	}

	const meta = (await metaResponse.json()) as {
		latest?: string;
	};

	if (!meta.latest) {
		throw new Error('No se pudo resolver la ultima version de docxml.');
	}

	const version = meta.latest;
	const versionMetaResponse = await fetch(
		`https://jsr.io/${DOCXML_PACKAGE_NAME}/${version}_meta.json`
	);

	if (!versionMetaResponse.ok) {
		throw new Error(
			'No se pudo cargar el manifiesto de la ultima version.'
		);
	}

	const versionMeta = (await versionMetaResponse.json()) as {
		manifest?: Record<string, unknown>;
	};

	if (!versionMeta.manifest) {
		throw new Error('No se pudo resolver el manifiesto de la version.');
	}

	monaco.languages.typescript.typescriptDefaults.setCompilerOptions({
		target: monaco.languages.typescript.ScriptTarget.ES2020,
		module: monaco.languages.typescript.ModuleKind.ESNext,
		moduleResolution:
			monaco.languages.typescript.ModuleResolutionKind.NodeJs,
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

	const sourceFiles = Object.keys(versionMeta.manifest).filter(
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
		'file:///playground/docxml.d.ts'
	);
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

	throw new Error(
		'El resultado debe ser Docx, Archive, Uint8Array, Blob o ArrayBuffer.'
	);
}

function downloadFile(data: Uint8Array, fileName: string) {
	const safeBytes = Uint8Array.from(data);
	const blob = new Blob([safeBytes], {
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
		fileName: 'main.ts',
		reportDiagnostics: false,
	}).outputText;
}

export function App() {
	const editorElementRef = useRef<HTMLDivElement | null>(null);
	const editorRef = useRef<monaco.editor.IStandaloneCodeEditor | null>(null);
	const [status, setStatus] = useState('Inicializando...');
	const [hasError, setHasError] = useState(false);
	const [isGenerating, setIsGenerating] = useState(false);

	useEffect(() => {
		ensureBrowserPolyfills();

		if (!editorElementRef.current) {
			setStatus('No se encontro el contenedor del editor.');
			setHasError(true);
			return;
		}

		const model = monaco.editor.createModel(
			INITIAL_SOURCE,
			'typescript',
			monaco.Uri.parse(PLAYGROUND_FILE)
		);

		editorRef.current = monaco.editor.create(editorElementRef.current, {
			model,
			theme: 'vs',
			automaticLayout: true,
			minimap: { enabled: false },
			fontSize: 13,
		});

		setStatus('Cargando tipos desde JSR...');
		setHasError(false);

		void installMonacoTypes()
			.then(() => {
				setStatus('Listo.');
				setHasError(false);
			})
			.catch((error) => {
				const message =
					error instanceof Error ? error.message : String(error);
				setStatus(`Tipos no disponibles: ${message}`);
				setHasError(true);
			});

		return () => {
			model.dispose();
			editorRef.current?.dispose();
			editorRef.current = null;
		};
	}, []);

	async function generateDocx() {
		const editor = editorRef.current;
		if (!editor) {
			setStatus('El editor aun no esta listo.');
			setHasError(true);
			return;
		}

		setIsGenerating(true);
		setHasError(false);
		setStatus('Compilando...');

		try {
			const source = editor.getValue();
			const jsSource = transpileTypeScript(source)
				.replaceAll(`from 'docxml'`, `from '${DOCXML_RUNTIME_URL}'`)
				.replaceAll(
					`from \"docxml\"`,
					`from \"${DOCXML_RUNTIME_URL}\"`
				);

			const moduleBlob = new Blob([jsSource], {
				type: 'text/javascript',
			});
			const moduleUrl = URL.createObjectURL(moduleBlob);

			try {
				const mod = (await import(moduleUrl)) as {
					default?: () => unknown | Promise<unknown>;
				};

				if (typeof mod.default !== 'function') {
					throw new Error(
						'El modulo debe exportar una funcion default.'
					);
				}

				setStatus('Generando DOCX...');
				const result = await mod.default();
				const data = await normalizeResult(result);
				downloadFile(data, 'playground-output.docx');
				setStatus(`Documento generado (${data.byteLength} bytes).`);
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

	return (
		<div class='layout'>
			<header class='header'>
				<h1>Docxml Playground</h1>
				<p>
					Playground minimo: escribe TypeScript, importa{' '}
					<code>docxml</code> y genera un .docx.
				</p>
			</header>

			<div ref={editorElementRef} class='editor' />

			<div class='actions'>
				<button
					type='button'
					class='primary'
					disabled={isGenerating}
					onClick={() => void generateDocx()}
				>
					{isGenerating ? 'Generando...' : 'Generar DOCX'}
				</button>
			</div>

			<p class={`status${hasError ? ' error' : ''}`}>{status}</p>
		</div>
	);
}

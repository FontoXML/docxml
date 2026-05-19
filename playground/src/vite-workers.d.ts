declare module 'monaco-editor/esm/vs/editor/editor.worker?worker' {
	const EditorWorker: {
		new (): Worker;
	};
	export default EditorWorker;
}

declare module 'monaco-editor/esm/vs/language/typescript/ts.worker?worker' {
	const TypeScriptWorker: {
		new (): Worker;
	};
	export default TypeScriptWorker;
}

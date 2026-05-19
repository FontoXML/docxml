export type PlaygroundExample = {
	id: string;
	label: string;
	sourcePath: string;
	source: string;
};

export const playgroundExamples: PlaygroundExample[] = [
	{
		id: 'hello',
		label: 'Simple paragraph',
		sourcePath: '',
		source: `import Docx, { Paragraph, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.set(
		new Paragraph({}, new Text({}, 'Hello from the docxml playground.'))
	);

	return docx;
}
`,
	},
];

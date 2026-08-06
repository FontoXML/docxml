// deno-lint-ignore-file jsx-key
/** @jsx Docx.jsx */

import Docx, { Paragraph, StructuredDocument, Text } from '../mod.ts';

await Docx.fromJsx([
	<StructuredDocument appearance='hidden'>
		<Paragraph>
			<Text>Content</Text>
		</Paragraph>
	</StructuredDocument>,
]).toFile('structured-document.docx');

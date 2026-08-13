// deno-lint-ignore-file jsx-key
/** @jsx Docx.jsx */

import Docx, { Paragraph, StructuredDocument, Text } from '../mod.ts';

await Docx.fromJsx([
	<StructuredDocument alias='test' appearance='hidden'>
		<Paragraph>
			<Text>Content</Text>
		</Paragraph>
		<StructuredDocument alias='nested' lock='sdtLocked'>
			<Paragraph>
				<Text>More Content</Text>
			</Paragraph>
		</StructuredDocument>
	</StructuredDocument>,
]).toFile('structured-document.docx');

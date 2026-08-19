// deno-lint-ignore-file jsx-key
/** @jsx Docx.jsx */

import Docx, { Paragraph, StructuredDocument, Text } from '../mod.ts';

await Docx.fromJsx([
	<StructuredDocument alias='test' appearance='hidden' tag='tag1'>
		<Paragraph>
			<Text>Content</Text>
		</Paragraph>
		<StructuredDocument lock='sdtLocked' tag='tag2'>
			<Paragraph>
				<Text>More Content</Text>
			</Paragraph>
		</StructuredDocument>
	</StructuredDocument>,
]).toFile('structured-document.docx');

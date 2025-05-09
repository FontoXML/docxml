/** @jsx  Docx.jsx */
import Docx, { FootnoteReference, Paragraph } from '../mod.ts';
import { Text } from '../src/components/Text.ts';
import { FootnoteType } from '../src/files/FootnotesXml.ts';

const docxFile = Docx.fromNothing();
const footnote = docxFile.document.footnotes.add(
	new Paragraph({}, new Text({}, 'Hello, this is a footnote.')),
	'separator'
);

docxFile.document.set(
	<Paragraph>
		This is my first paragraph of text.{' '}
		<FootnoteReference id={footnote.id} />
	</Paragraph>
);

await docxFile.toFile('footnotes.docx');

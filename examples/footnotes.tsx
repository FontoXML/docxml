/** @jsx  Docx.jsx */
import Docx, { FootnoteReference, Paragraph } from '../mod.ts';
import { Text } from '../src/components/Text.ts';

const docxFile = Docx.fromNothing();
const footnote1 = docxFile.document.footnotes.add(
	new Paragraph({}, new Text({}, 'Hello, this is a footnote.')),
	'normal'
);
const footnote2 = docxFile.document.footnotes.add(
	new Paragraph({}, new Text({}, 'And this is an additional footnote.')),
	'normal'
);

docxFile.document.settings.set('footnoteProperties', {
	numberingFormat: 'lowerRoman',
	position: 'beneathText',
	restart: 'page',
});

docxFile.document.set(
	<Paragraph>
		This is my first paragraph of text.{' '}
		<FootnoteReference id={footnote1.id} />
		<FootnoteReference id={footnote2.id} />
	</Paragraph>
);

await docxFile.toFile('footnotes.docx');
